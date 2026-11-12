"""Keep bookkeeping internals out of invoice and workbook directories."""
from __future__ import annotations

import hashlib
import os
from contextlib import contextmanager
from pathlib import Path


def hide_internal(path):
    """Set the Windows hidden bit without altering file bytes or other flags."""
    path = Path(path)
    if os.name != 'nt' or not path.exists(): return
    import ctypes
    from ctypes import wintypes
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    get = kernel.GetFileAttributesW
    get.argtypes = [wintypes.LPCWSTR]; get.restype = wintypes.DWORD
    set_attributes = kernel.SetFileAttributesW
    set_attributes.argtypes = [wintypes.LPCWSTR, wintypes.DWORD]
    set_attributes.restype = wintypes.BOOL
    flags = get(str(path))
    if flags != 0xffffffff and not flags & 2:
        set_attributes(str(path), flags | 2)


def prepare_runtime(root):
    folder = Path(root)/'.运行记录'
    folder.mkdir(parents=True, exist_ok=True)
    hide_internal(folder)
    return folder


def application_lock_path(logical_path):
    # The same canonical resource needs the same lock even for different roots,
    # state files, accounts or skill installations.
    resource = os.path.normcase(str(Path(logical_path).resolve()))
    key = hashlib.sha256(resource.encode('utf-8')).hexdigest()
    appdata = Path(os.environ.get('LOCALAPPDATA') or Path.home()/'.local/share')
    return appdata/'InvoiceLedgerSkill/locks'/(key+'.lock')


@contextmanager
def runtime_exclusive(logical_path):
    from run_invoice_workflow import exclusive
    logical_path = Path(logical_path)
    with exclusive(application_lock_path(logical_path)):
        # Existing guards retain their inode so a running older installation
        # cannot bypass the new lock. Do not create new guards in business dirs.
        old_guard = Path(str(logical_path)+'.guard')
        if old_guard.exists() or logical_path.exists():
            hide_internal(old_guard)
            with exclusive(logical_path):
                hide_internal(old_guard)
                yield
        else:
            yield
