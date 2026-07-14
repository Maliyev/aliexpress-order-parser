import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from run_latest_order import make_backup


class BackupTests(unittest.TestCase):
    def test_backup_is_saved_from_open_workbook(self):
        workbook = Workbook()
        workbook.active["A1"] = "keep this value"

        with tempfile.TemporaryDirectory() as folder:
            folder_path = Path(folder)
            backup_path = make_backup(
                workbook,
                folder_path / "orders.xlsx",
                folder_path / "backups",
            )

            self.assertTrue(backup_path.exists())
            self.assertEqual(load_workbook(backup_path).active["A1"].value, "keep this value")
