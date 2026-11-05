import json
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch
import status_sync_hook as hook

WORKER = """import sys,time
from pathlib import Path
sys.path.insert(0,sys.argv[1])
import status_sync_hook as h
root=Path(sys.argv[2]);who=sys.argv[3]
if who=='first':
 original=h._load_approvals
 def pause_after_read(root):
  result=original(root);(root/'first-read').touch()
  deadline=time.monotonic()+5
  while not (root/'release').exists() and time.monotonic()<deadline:time.sleep(.01)
  return result
 h._load_approvals=pause_after_read
print(h.record_approval(root,{'session_id':who,'prompt':'批准修改组件'}))
"""

class ApprovalPersistenceTests(unittest.TestCase):
 def test_two_real_hook_processes_preserve_both_approvals(self):
  with tempfile.TemporaryDirectory() as td:
   root=Path(td);scripts=str(Path(hook.__file__).parent)
   first=subprocess.Popen([sys.executable,'-c',WORKER,scripts,td,'first'],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True)
   second=None
   try:
    deadline=time.monotonic()+5
    while not (root/'first-read').exists() and time.monotonic()<deadline:time.sleep(.01)
    self.assertTrue((root/'first-read').exists())
    second=subprocess.Popen([sys.executable,'-c',WORKER,scripts,td,'second'],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True)
    try:b,be=second.communicate(timeout=.3)
    except subprocess.TimeoutExpired:b=be=None
    (root/'release').touch();a,ae=first.communicate(timeout=5)
    if b is None:b,be=second.communicate(timeout=5)
    self.assertEqual((first.returncode,second.returncode),(0,0),(ae,be))
    self.assertIn('已记录用户批准',a);self.assertIn('已记录用户批准',b)
    self.assertEqual(set(json.loads((root/hook.APPROVALS).read_text())),{'first','second'})
   finally:
    (root/'release').touch()
    for process in [first,second]:
     if process is not None and process.poll() is None:process.kill();process.communicate()

 def test_success_requires_actual_readback(self):
  with tempfile.TemporaryDirectory() as td:
   root=Path(td);actual=Path.write_text
   def silently_discard(file,data,*args,**kwargs):
    if file==root/hook.APPROVALS:return len(data)
    return actual(file,data,*args,**kwargs)
   with patch.object(Path,'write_text',silently_discard):
    note=hook.record_approval(root,{'session_id':'one','prompt':'批准修改组件'})
   # A writer that ignores write_text may use an atomic rename; if so the file
   # really exists and normal success is correct. Success must match persistence.
   exists=(root/hook.APPROVALS).exists()
   self.assertEqual('已记录用户批准' in note,exists)

 def test_replace_failure_keeps_original_and_cleans_temporary_file(self):
  with tempfile.TemporaryDirectory() as td:
   root=Path(td)
   hook.record_approval(root,{'session_id':'original','prompt':'批准修改组件'})
   before=(root/hook.APPROVALS).read_bytes()
   with patch.object(hook.os,'replace',side_effect=PermissionError('replace denied')):
    result=hook.record_approval(root,{'session_id':'new','prompt':'批准修改组件'})
   self.assertIn('保存失败',result);self.assertNotIn('已记录用户批准',result)
   self.assertEqual((root/hook.APPROVALS).read_bytes(),before)
   self.assertEqual(list((root/'.runtime').glob('*.tmp')),[])

 def test_readback_mismatch_never_reports_success(self):
  with tempfile.TemporaryDirectory() as td:
   root=Path(td);original=Path.read_bytes
   def stale_read(file):
    return b'{}' if file==root/hook.APPROVALS else original(file)
   with patch.object(Path,'read_bytes',stale_read):
    result=hook.record_approval(root,{'session_id':'one','prompt':'批准修改组件'})
   self.assertIn('读回不一致',result);self.assertNotIn('已记录用户批准',result)
   self.assertEqual(list((root/'.runtime').glob('*.tmp')),[])

 def test_reading_permission_does_not_extend_its_expiry(self):
  with tempfile.TemporaryDirectory() as td:
   root=Path(td);payload={'session_id':'one','prompt':'批准修改组件'}
   hook.record_approval(root,payload);before=(root/hook.APPROVALS).read_bytes()
   self.assertTrue(hook.approved(root,payload))
   self.assertIn('已获用户批准',hook.record_approval(root,{**payload,'prompt':'继续查问题'}))
   self.assertEqual((root/hook.APPROVALS).read_bytes(),before)
   debug=json.loads((root/'.runtime/组件软锁诊断.json').read_text())
   self.assertEqual(debug['approvalPath'],str((root/hook.APPROVALS).resolve()))
   self.assertEqual(debug['session'],'one');self.assertTrue(debug['approved'])
   self.assertEqual(debug['lastApproval']['session'],'one')
   self.assertTrue(debug['lastApproval']['saved'])
   self.assertEqual(debug['lastCheck']['session'],'one')
   self.assertTrue(debug['lastCheck']['approved'])
   self.assertNotIn('prompt',debug['lastApproval'])
   hook.approved(root,{'session_id':'another'})
   latest=json.loads((root/'.runtime/组件软锁诊断.json').read_text())
   self.assertEqual(latest['lastApproval'],debug['lastApproval'])
   self.assertEqual(latest['lastCheck']['session'],'another')
   self.assertFalse(latest['lastCheck']['approved'])


 def test_maintenance_switch_is_reported_as_its_own_authorization_source(self):
  with tempfile.TemporaryDirectory() as td:
   root=Path(td);(root/'产品管理').mkdir()
   (root/hook.SWITCH).write_text('{"开启":true}')
   self.assertIsNone(hook.component_violation(root,{'session_id':'one','hook_event_name':'PreToolUse',
     'tool_name':'Edit','tool_input':{'file_path':str(root/'.agents/skills/campus-job-application/SKILL.md')}}))
   report=json.loads((root/'.runtime/组件软锁诊断.json').read_text())
   self.assertEqual(report['authorizationSource'],'user_maintenance_switch')
   self.assertTrue(report['approved']);self.assertFalse((root/hook.APPROVALS).exists())

if __name__=='__main__':unittest.main()
