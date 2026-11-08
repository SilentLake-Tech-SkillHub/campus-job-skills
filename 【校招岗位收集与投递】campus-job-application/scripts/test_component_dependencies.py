import json,os,tempfile,unittest
from pathlib import Path
from unittest.mock import patch
import component_dependencies as deps
class DependencyTests(unittest.TestCase):
    def test_installed_contract_matches_declared_owner(self):
        data=deps.load_contract('toolArguments');self.assertEqual(data['schemaVersion'],1)
    def test_explicit_missing_and_wrong_revision_do_not_fall_back(self):
        with tempfile.TemporaryDirectory() as temp:
            file=Path(temp)/'missing.json'
            with patch.dict(os.environ,{'RECRUITING_TOOL_ARGUMENTS_FILE':str(file)}):
                with self.assertRaisesRegex(ValueError,'COMPONENT_DEPENDENCY_MISSING'):deps.load_contract('toolArguments')
                file.write_text('{"textTools":["arbitrary"]}')
                with self.assertRaisesRegex(ValueError,'COMPONENT_DEPENDENCY_VERSION_MISMATCH'):deps.load_contract('toolArguments')
    def test_checkout_layout_resolves_original_owner_file(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);component=root/'search';component.mkdir()
            spec=deps.dependency('toolArguments');target=root/spec['checkoutPath'];target.parent.mkdir(parents=True);target.write_bytes(deps.contract_path('toolArguments').read_bytes())
            with patch.object(deps,'COMPONENT',component),patch.dict(os.environ,{},clear=True):
                self.assertEqual(deps.contract_path('toolArguments'),target)
                self.assertEqual(deps.load_contract('toolArguments')['schemaVersion'],1)

    def test_windows_line_endings_do_not_change_contract_identity(self):
        original=deps.contract_path('toolArguments').read_text(encoding='utf-8')
        with tempfile.TemporaryDirectory() as temp:
            file=Path(temp)/'contract.json';file.write_bytes(original.replace('\n','\r\n').encode('utf-8'))
            self.assertEqual(deps.load_contract('toolArguments',file),json.loads(original))
