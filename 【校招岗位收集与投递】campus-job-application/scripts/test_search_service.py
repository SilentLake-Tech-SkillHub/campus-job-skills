import asyncio,json,os,subprocess,sys,tempfile,unittest
from pathlib import Path
from unittest.mock import patch
from search_paths import configure
from search_service import execute
from search_mcp_client import call
from test_browser_capabilities import fixture_environment

class IndependentServiceTests(unittest.TestCase):
    def test_cli_and_mcp_read_same_persisted_queue(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);(root/'authorization.md').write_text('Synthetic authorization')
            env={**fixture_environment(root),'CAMPUS_RECRUITING_WORKSPACE':temp,'CAMPUS_SEARCH_PRIVATE_DIR':temp}
            with patch.dict(os.environ,env):
                manifest={'batch_id':'standalone','scope':{'region':'synthetic'},'authorization_ref':'authorization.md','workbook_sha256':'a'*64,'companies':[{'company_key':'C00001','name':'Synthetic','source_key':'C00001'}]}
                execute('search_batch_create',{'manifest':manifest})
                script=Path(__file__).with_name('search_cli.py')
                proc=subprocess.run([sys.executable,'-B',str(script),'search_round_check','--workspace',temp],input=json.dumps({'batch_id':'standalone'}),text=True,capture_output=True,check=True)
                independent=json.loads(proc.stdout)
                native=asyncio.run(call({'tool':'search_round_check','arguments':{'batch_id':'standalone'}}))
                self.assertEqual(independent,native)
                self.assertFalse(independent['passed'])
    def test_component_directory_cannot_hold_private_search_data(self):
        with patch.dict(os.environ,{'CAMPUS_SEARCH_PRIVATE_DIR':str(Path(__file__).resolve().parent)}):
            with self.assertRaisesRegex(ValueError,'USER_DATA_OUTSIDE_COMPONENT_REQUIRED'):configure('/tmp/synthetic-workspace')

    def test_mcp_and_cli_report_identical_release_and_stale_schema_is_rejected(self):
        with tempfile.TemporaryDirectory() as temp:
            with patch.dict(os.environ,{'CAMPUS_RECRUITING_WORKSPACE':temp,'CAMPUS_SEARCH_PRIVATE_DIR':temp}):
                core=execute('search_runtime_info',{})
                native=asyncio.run(call({'tool':'search_runtime_info','arguments':{}}))
                self.assertEqual(core,native)
                env={**os.environ,'CAMPUS_SEARCH_EXPECTED_SCHEMA':'old'}
                result=subprocess.run([sys.executable,'-B',str(Path(__file__).with_name('search_cli.py')),'search_runtime_info'],env=env,input='{}',text=True,capture_output=True)
                self.assertNotEqual(result.returncode,0);self.assertIn('SEARCH_INTERFACE_SCHEMA_STALE',result.stderr)
