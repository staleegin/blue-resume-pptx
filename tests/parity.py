"""一致性测试：用同一份 JSON 分别跑 JS 版和 Python 版，逐个形状比较。
用法（仓库根目录，需要已 npm install）:  python tests/parity.py examples/*.json"""
import json, os, subprocess, sys, tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JS = os.path.join(ROOT, 'skill/blue-resume-pptx/scripts/build_resume.js')
PY = os.path.join(ROOT, 'resume_pptx.py')
CMP = os.path.join(ROOT, 'tests/compare.py')

fail = 0
with tempfile.TemporaryDirectory() as tmp:
    for f in sys.argv[1:]:
        b = os.path.splitext(os.path.basename(f))[0]
        j, p = os.path.join(tmp, 'j_' + b + '.pptx'), os.path.join(tmp, 'p_' + b + '.pptx')
        rj = json.loads(subprocess.run(['node', JS, f, j], capture_output=True, text=True, check=True).stdout.splitlines()[0])
        rp = json.loads(subprocess.run([sys.executable, PY, f, p], capture_output=True, text=True, check=True).stdout.splitlines()[0])
        same_result = all(rj[k] == rp[k] for k in ('level', 'overflowPt', 'movedToRight'))
        c = subprocess.run([sys.executable, CMP, j, p], capture_output=True, text=True)
        ok = same_result and c.returncode == 0
        fail += not ok
        print('PASS' if ok else 'FAIL', f, '' if ok else c.stdout[-800:])
print('%d 份，%d 份不一致' % (len(sys.argv) - 1, fail))
sys.exit(1 if fail else 0)
