"""Finance entrypoint: machine-local binding; legacy profiles remain explicit."""
import argparse
import json
import sys
from pathlib import Path

from setup_workspace import BINDING, SetupRequired, load_binding, require_account, setup_message


def main(argv=None):
    arguments = list(sys.argv[1:] if argv is None else argv)
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--config')
    options, _ = parser.parse_known_args(arguments)
    if options.config is None:
        try:
            root = load_binding(BINDING)
            options.config = str(root/'.运行记录/财务配置.json')
            if not Path(options.config).is_file(): raise SetupRequired('account_setup_required', root)
        except SetupRequired as error:
            print(json.dumps(setup_message(error.status, error.root), ensure_ascii=False))
            return 3
        arguments = ['--config', options.config] + arguments
    config = json.loads(Path(options.config).read_text(encoding='utf-8-sig'))
    if config.get('workspace_root'):
        try:
            require_account(config, Path(config['workspace_root']))
        except SetupRequired as error:
            print(json.dumps(setup_message(error.status, error.root), ensure_ascii=False))
            return 3
        if not any(arg == '--result' or arg.startswith('--result=') for arg in arguments):
            name = '历史建档结果.json' if '--index-history' in arguments else '最近结果.json'
            arguments += ['--result', str(Path(config['workspace_root'])/'.运行记录'/name)]
        from run_workspace import main as workspace_main
        return workspace_main(arguments)
    from run_invoice_workflow import main as legacy_main
    previous = sys.argv
    try:
        sys.argv = [previous[0]] + arguments
        return legacy_main(expected_department='finance')
    finally:
        sys.argv = previous

if __name__ == '__main__':
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'): stream.reconfigure(encoding='utf-8')
    raise SystemExit(main())
