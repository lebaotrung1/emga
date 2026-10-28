"""Prepare the local interpreter without depending on the system pip installation."""
import argparse
import json
import os
from pathlib import Path
import subprocess
import sys
import uuid
import venv
from datetime import datetime

MIN_PYTHON = (3, 12)
MODULES = {'flask':'flask','waitress':'waitress','openpyxl':'openpyxl','python-docx':'docx'}

class SetupError(RuntimeError):
    pass

def env_python(env):
    return env / ('Scripts/python.exe' if os.name=='nt' else 'bin/python')

def requirements_at(root):
    requirements=[]
    for line in (root/'requirements.txt').read_text(encoding='utf-8').splitlines():
        line=line.strip()
        if not line or line.startswith('#'):continue
        if line.count('==')!=1:raise SetupError('requirements.txt can dung dinh dang ten==phien_ban.')
        name,version=line.split('==')
        if name.casefold() not in MODULES:raise SetupError('Thu vien chua co trong danh sach kiem tra: '+name)
        requirements.append((name,version,MODULES[name.casefold()]))
    if not requirements:raise SetupError('requirements.txt khong co thu vien.')
    return requirements

def dependencies_ok(python,requirements):
    code=("import importlib, importlib.metadata, json; specs=json.loads("+repr(json.dumps(requirements))+"); "
          "[(importlib.import_module(module), version == importlib.metadata.version(name) or "
          "(_ for _ in ()).throw(RuntimeError('version mismatch'))) for name,version,module in specs]")
    try:
        result=subprocess.run([str(python),'-I','-c',code],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
        return result.returncode==0
    except OSError:return False

def pip_ok(python):
    try:
        result=subprocess.run([str(python),'-I','-m','pip','--version'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
        return result.returncode==0
    except OSError:return False

def unique_backup(env,kind):
    return env.with_name(env.name+'-'+kind+'-'+datetime.now().strftime('%Y%m%d-%H%M%S')+'-'+uuid.uuid4().hex[:6])

def safe_to_replace(env):
    if env.is_symlink() or (env.exists() and not env.is_dir()):
        raise SetupError('.venv phai la thu muc thuong; khong tu thay the lien ket hoac tep.')
    configured=os.environ.get('EMGA_DATA_DIR')
    if configured:
        data=Path(configured).expanduser().resolve()
        if data==env.resolve() or env.resolve() in data.parents:
            raise SetupError('EMGA_DATA_DIR dang nam trong .venv. Sao luu va chuyen du lieu ra ngoai .venv truoc khi sua moi truong.')

def rebuild_environment(env):
    safe_to_replace(env)
    backup=None
    if env.exists():
        backup=unique_backup(env,'backup')
        try:env.rename(backup)
        except OSError as e:
            raise SetupError('Khong doi ten duoc .venv. Hay dong cua so Emga/Python cu roi thu lai.') from e
        print('Da giu moi truong cu tai: '+backup.name,flush=True)
    try:
        print('Dang tao moi moi truong thu vien tu Python da cai...',flush=True)
        venv.EnvBuilder(with_pip=True).create(env)
        if not pip_ok(env_python(env)):
            raise SetupError('pip trong moi truong vua tao van bi loi. Hay dung Modify/Repair trong bo cai Python tu python.org.')
    except Exception:
        # Keep every old file and the failed attempt, and restore the previous path.
        if env.exists():env.rename(unique_backup(env,'failed'))
        if backup is not None:backup.rename(env)
        raise
    return backup

def prepare(root,repair=False):
    root=Path(root).resolve();env=root/'.venv';python=env_python(env)
    requirements=requirements_at(root)
    if env.is_symlink():raise SetupError('Khong tu sua .venv la lien ket. Hay dung thu muc moi truong rieng cua ung dung.')
    if not repair and python.is_file() and dependencies_ok(python,requirements):
        print('Thu vien da san sang. Khong can chay pip hay tai lai.',flush=True)
        return python
    fresh=False
    if repair or not python.is_file() or not pip_ok(python):
        print('Moi truong thu vien chua du hoac pip bi loi; se tao lai, giu ban cu.',flush=True)
        rebuild_environment(env);fresh=True
    print('Dang cai thu vien. Lan cai/sua nay can ket noi Internet...',flush=True)
    command=[str(python),'-I','-m','pip','install','--disable-pip-version-check']
    if not fresh:command.append('--force-reinstall')
    command.extend(['-r',str(root/'requirements.txt')])
    result=subprocess.run(command,cwd=str(root))
    if result.returncode:
        raise SetupError('Chua cai xong thu vien (pip ma loi '+str(result.returncode)+'). Xem thong bao phia tren de phan biet loi ket noi hay thu vien; khong doi du lieu/tai khoan.')
    if not dependencies_ok(python,requirements):
        raise SetupError('Thu vien sau khi cai chua vuot qua kiem tra. Chay Sua_moi_truong.bat; neu van loi, gui nguyen thong bao de kiem tra.')
    return python

def main(argv=None):
    parser=argparse.ArgumentParser(description='Emga local environment launcher')
    parser.add_argument('--repair',action='store_true',help='Keep the previous environment and create a clean one')
    parser.add_argument('--check',action='store_true',help='Prepare/check dependencies without starting the server')
    args=parser.parse_args(argv)
    if sys.version_info<MIN_PYTHON:
        print('Can Python 3.12 tro len tu python.org. Python hien tai: '+sys.version.split()[0])
        return 1
    root=Path(__file__).resolve().parent
    try:
        python=prepare(root,args.repair)
        if args.check:
            print('Moi truong da san sang. Chay Chay_Emga.bat de mo phan mem.')
            return 0
        env=os.environ.copy();env['EMGA_HOST']='0.0.0.0';env.setdefault('EMGA_PORT','8765')
        print('Mo trinh duyet tren may chu: http://127.0.0.1:'+env['EMGA_PORT'],flush=True)
        print('May LAN dung IP may chu va cong tren. Ctrl+C de dung.',flush=True)
        return subprocess.call([str(python),str(root/'app.py')],cwd=str(root),env=env)
    except KeyboardInterrupt:return 130
    except (SetupError,OSError,RuntimeError,subprocess.SubprocessError) as e:
        print('\nKhong khoi dong duoc: '+str(e),file=sys.stderr,flush=True)
        return 1

if __name__=='__main__':sys.exit(main())
