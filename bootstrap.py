"""Prepare the local interpreter without depending on the system pip installation."""
import argparse
import hashlib
import struct
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

DIAGNOSTIC_LOG = None

def emit(message='',file=None,flush=True):
    message=str(message)
    stream=file or sys.stdout
    try:print(message,file=stream,flush=flush)
    except UnicodeEncodeError:
        print(message.encode(stream.encoding or 'utf-8',errors='replace').decode(stream.encoding or 'utf-8'),file=stream,flush=flush)
    if DIAGNOSTIC_LOG is not None:
        try:
            with DIAGNOSTIC_LOG.open('a',encoding='utf-8') as log:log.write(message+'\n')
        except OSError:pass

def choose_environment(root,windows=None):
    root=Path(root).resolve()
    windows=os.name=='nt' if windows is None else windows
    if not windows:return root/'.venv'
    # Keep Windows dependency paths short even if the ZIP was extracted twice.
    base=os.environ.get('EMGA_RUNTIME_DIR')
    if base:
        base=Path(base).expanduser()
        if not base.is_absolute():raise SetupError('EMGA_RUNTIME_DIR can la duong dan tuyet doi.')
    else:
        base=Path(os.environ.get('LOCALAPPDATA',str(Path.home()/'AppData'/'Local')))/'EmgaRuntime'
    fingerprint=hashlib.sha256((root/'requirements.txt').read_bytes()).hexdigest()[:10]
    name='py'+str(sys.version_info.major)+str(sys.version_info.minor)+'-'+str(struct.calcsize('P')*8)+'-'+fingerprint
    return base/name

def run_probe(command,report=False):
    try:
        result=subprocess.run(command,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,encoding='utf-8',errors='replace')
        if report:
            emit('Kiem tra: '+subprocess.list2cmdline(command))
            emit(result.stdout or '(khong co noi dung)')
            emit('Ma thoat: '+str(result.returncode))
        return result.returncode==0
    except OSError as e:
        if report:emit('Khong chay duoc chuong trinh: '+str(e))
        return False

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

def dependencies_ok(python,requirements,report=False):
    code=("import importlib, importlib.metadata, json; specs=json.loads("+repr(json.dumps(requirements))+"); "
          "[(importlib.import_module(module), version == importlib.metadata.version(name) or "
          "(_ for _ in ()).throw(RuntimeError('version mismatch'))) for name,version,module in specs]")
    return run_probe([str(python),'-I','-c',code],report)

def pip_ok(python,report=False):
    return run_probe([str(python),'-I','-m','pip','--version'],report)

def unique_backup(env,kind):
    return env.with_name(env.name+'-'+kind+'-'+datetime.now().strftime('%Y%m%d-%H%M%S')+'-'+uuid.uuid4().hex[:6])

def safe_to_replace(env):
    if env.is_symlink() or (env.exists() and not env.is_dir()):
        raise SetupError('Thu muc thu vien phai la thu muc thuong; khong tu thay the lien ket hoac tep.')
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
            raise SetupError('Khong doi ten duoc thu muc thu vien. Hay dong cua so Emga/Python cu roi thu lai.') from e
        emit('Da giu moi truong cu tai: '+backup.name,flush=True)
    try:
        emit('Dang tao moi moi truong thu vien tu Python da cai...',flush=True)
        venv.EnvBuilder(with_pip=True).create(env)
        if not pip_ok(env_python(env),report=True):
            raise SetupError('pip cua moi truong moi chua chay duoc. Loi goc o phia tren; hay gui tep emga-khoi-dong.log de chan doan, chua can cai lai Python.')
    except Exception as e:
        emit('Tao moi truong that bai: '+str(e))
        output=getattr(e,'output',None)
        if output:emit(output.decode('utf-8',errors='replace') if isinstance(output,bytes) else output)
        # Keep every old file and the failed attempt, and restore the previous path.
        if env.exists():env.rename(unique_backup(env,'failed'))
        if backup is not None:backup.rename(env)
        raise
    return backup

def prepare(root,repair=False):
    root=Path(root).resolve();env=choose_environment(root);python=env_python(env)
    emit('Thu muc thu vien dang dung: '+str(env))
    requirements=requirements_at(root)
    if env.is_symlink():raise SetupError('Khong tu sua .venv la lien ket. Hay dung thu muc moi truong rieng cua ung dung.')
    if not repair and python.is_file() and dependencies_ok(python,requirements):
        emit('Thu vien da san sang. Khong can chay pip hay tai lai.',flush=True)
        return python
    fresh=False
    if repair or not python.is_file() or not pip_ok(python):
        emit('Moi truong thu vien chua du hoac pip bi loi; se tao lai, giu ban cu.',flush=True)
        rebuild_environment(env);fresh=True
    emit('Dang cai thu vien. Lan cai/sua nay can ket noi Internet...',flush=True)
    command=[str(python),'-I','-m','pip','install','--disable-pip-version-check']
    if not fresh:command.append('--force-reinstall')
    command.extend(['-r',str(root/'requirements.txt')])
    result=subprocess.run(command,cwd=str(root))
    if result.returncode:
        raise SetupError('Chua cai xong thu vien (pip ma loi '+str(result.returncode)+'). Xem thong bao phia tren de phan biet loi ket noi hay thu vien; khong doi du lieu/tai khoan.')
    if not dependencies_ok(python,requirements,report=True):
        raise SetupError('Thu vien sau khi cai chua vuot qua kiem tra. Chay Sua_moi_truong.bat; neu van loi, gui nguyen thong bao de kiem tra.')
    return python

def diagnose(root):
    emit('Chan doan chi doc, khong cai/sua thu vien, khong doc co so du lieu, khong dump bien moi truong hay doc tep bi mat.')
    run_probe([sys.executable,'-I','-m','ensurepip','--version'],report=True)
    run_probe([sys.executable,'-I','-m','pip','--version'],report=True)
    selected=choose_environment(root)
    candidates=[selected,root/'.venv']
    failed=sorted(root.glob('.venv-failed-*'),key=lambda p:p.name,reverse=True)
    cached_failed=sorted(selected.parent.glob(selected.name+'-failed-*'),key=lambda p:p.name,reverse=True)
    candidates.extend(failed[:1]);candidates.extend(cached_failed[:1])
    seen=set()
    for env in candidates:
        if str(env) in seen:continue
        seen.add(str(env));python=env_python(env)
        if python.is_file():pip_ok(python,report=True)
        else:emit('Chua co Python tai: '+str(python))
    emit('Da ghi chan doan. Gui noi dung emga-khoi-dong.log neu can kiem tra; tep co duong dan may.')

def main(argv=None):
    parser=argparse.ArgumentParser(description='Emga local environment launcher')
    parser.add_argument('--repair',action='store_true',help='Keep the previous environment and create a clean one')
    parser.add_argument('--diagnose',action='store_true',help='Read local Python/pip diagnostics without installing anything')
    parser.add_argument('--check',action='store_true',help='Prepare/check dependencies without starting the server')
    args=parser.parse_args(argv)
    if sys.version_info<MIN_PYTHON:
        emit('Can Python 3.12 tro len tu python.org. Python hien tai: '+sys.version.split()[0])
        return 1
    root=Path(__file__).resolve().parent
    global DIAGNOSTIC_LOG
    DIAGNOSTIC_LOG=root/'emga-khoi-dong.log'
    emit('\n===== '+datetime.now().isoformat(timespec='seconds')+' | '+('chan doan' if args.diagnose else 'sua moi truong' if args.repair else 'khoi dong')+' =====')
    emit('Emga 0.3.3 | Python '+sys.version.split()[0])
    emit('Python goc: '+sys.executable)
    emit('Nhat ky chan doan: '+str(DIAGNOSTIC_LOG))
    try:
        if args.diagnose:
            diagnose(root)
            return 0
        python=prepare(root,args.repair)
        if args.check:
            emit('Moi truong da san sang. Chay Chay_Emga.bat de mo phan mem.')
            return 0
        env=os.environ.copy();env['EMGA_HOST']='0.0.0.0';env.setdefault('EMGA_PORT','8765')
        emit('Mo trinh duyet tren may chu: http://127.0.0.1:'+env['EMGA_PORT'],flush=True)
        emit('May LAN dung IP may chu va cong tren. Ctrl+C de dung.',flush=True)
        return subprocess.call([str(python),str(root/'app.py')],cwd=str(root),env=env)
    except KeyboardInterrupt:return 130
    except (SetupError,OSError,RuntimeError,subprocess.SubprocessError) as e:
        emit('\nKhong khoi dong duoc: '+str(e),file=sys.stderr,flush=True)
        return 1

if __name__=='__main__':sys.exit(main())
