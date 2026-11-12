"""Remove only unchanged inbox inputs whose durable destinations are verified."""
import json
import os
from pathlib import Path

from archive_index import ArchiveIndex, identities
from run_invoice_workflow import digest
from workspace_runtime import runtime_exclusive


def capture_inputs(inputs, root):
    receipts={}
    for raw in inputs:
        path=Path(os.path.abspath(raw))
        # Never follow a junction/symlink and remove its target or an attachment.
        if not path.is_relative_to(root) or path.resolve()!=path:continue
        try:
            if path.is_file():receipts[str(path)]=digest(path.read_bytes())
        except OSError:pass  # The workflow reports inaccessible inputs itself.
    return receipts


def cleanup_completed_inputs(paths, ledger, receipts):
    result={'cleaned_input_files':0,'retained_input_files':0,'errors':[]}
    if not receipts:return result
    index_path=paths['state_path'].parent/'归档身份索引.json'
    with runtime_exclusive(paths['archive_dir']/'.finance-archive.lock'), runtime_exclusive(paths['state_path'].with_suffix('.lock')), runtime_exclusive(ledger.with_suffix('.lock')), runtime_exclusive(index_path.with_suffix('.lock')):
        state=json.loads(paths['state_path'].read_text(encoding='utf-8'))
        index=ArchiveIndex(paths['archive_dir'],index_path,state['scope'])
        queue_path=paths['state_path'].parent/'待复核队列.json'
        queue=json.loads(queue_path.read_text(encoding='utf-8'))
        if queue.get('scope')!=state['scope']:raise RuntimeError('复核记录与工作区不一致，未清理待处理原件')
        for raw,expected in receipts.items():
            path=Path(raw)
            try:
                if not path.exists():continue  # An earlier completed invocation may have removed it.
                if not path.is_relative_to(paths['input_dir']) or path.resolve()!=path:
                    raise RuntimeError('原件路径已变化，保留文件')
                if digest(path.read_bytes())!=expected:raise RuntimeError('原件在处理后被修改，保留新内容')
                record=state['sources'].get(expected,{})
                if record.get('status')=='complete':
                    if not ledger.is_file() or digest(ledger.read_bytes())!=state.get('ledger_sha256'):
                        raise RuntimeError('当前台账未确认保存或被修改，保留原件')
                    batch=state['batches'].get(record.get('batch_id'),{})
                    manifest=batch.get('identity_manifest') or {}
                    invoice_ids={invoice_id for invoice_id,sources in manifest.get('source_documents',{}).items()
                        if any(row['source_sha256']==expected for row in sources)}
                    if not invoice_ids or not invoice_ids<=set(batch.get('exported_ids',[])):
                        raise RuntimeError('整份原件尚未确认全部入账，保留原件')
                    metadata=identities([manifest['archive_identities'][invoice_id] for invoice_id in invoice_ids])
                    known=index.source(expected)
                    if not known or known['identities']!=metadata:
                        raise RuntimeError('完整发票集合的归档原件未确认，保留输入')
                elif record.get('status')=='exception':
                    target=Path(record.get('exception_path') or '__missing__').resolve()
                    if target==paths['exception_dir'] or not target.is_relative_to(paths['exception_dir']):
                        raise RuntimeError('异常保存路径不合法，保留原件')
                    if not target.is_file() or digest(target.read_bytes())!=expected:
                        raise RuntimeError('异常原件未保存完整，保留输入')
                    note=target.with_suffix(target.suffix+'.原因.txt')
                    if not note.is_file() or note.read_bytes()!=(record['reason']+'\n').encode('utf-8'):
                        raise RuntimeError('异常原因未保存完整，保留原件')
                    if not any(row['source_sha256']==expected and Path(row['local_path']).resolve()==target
                            and row.get('reason')==record['reason'] for row in queue['items']):
                        raise RuntimeError('异常复核队列未保存完整，保留原件')
                else:raise RuntimeError('处理尚未完成，保留原件以便续跑')
                # Recheck immediately before removing this exact inbox file.
                if path.resolve()!=path or digest(path.read_bytes())!=expected:
                    raise RuntimeError('清理前原件发生变化，保留文件')
                path.unlink();result['cleaned_input_files']+=1
            except Exception as error:
                result['retained_input_files']+=1
                result['errors'].append({'source':path.name,'error':'待处理原件未清理：'+str(error)})
    return result
