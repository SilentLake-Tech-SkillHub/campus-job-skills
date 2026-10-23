import tempfile
import types
import unittest
from unittest.mock import patch
import json_excel_sync as S


class PortableLockTests(unittest.TestCase):
    def test_windows_lock_and_unlock_use_same_byte_and_nonblocking_mode(self):
        calls = []
        backend = types.SimpleNamespace(LK_UNLCK=0, LK_NBLCK=2,
            locking=lambda fd, mode, size: calls.append((fd, mode, size)))
        with tempfile.TemporaryFile('w+b') as file, patch.object(S, 'fcntl', None), patch.dict('sys.modules', {'msvcrt': backend}):
            file.write(b'\0')
            S._file_lock(file)
            self.assertEqual(file.tell(), 0)
            S._file_lock(file, unlock=True)
            self.assertEqual(calls, [(file.fileno(), 2, 1), (file.fileno(), 0, 1)])


if __name__ == '__main__':
    unittest.main()
