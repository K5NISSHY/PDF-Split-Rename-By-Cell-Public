import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import call, patch

import pdf_splitter_by_cell as app
from openpyxl import Workbook
from pypdf import PdfReader, PdfWriter


class FakeSheet:
    def __init__(self, cells):
        self.cells = cells

    def __getitem__(self, address):
        return self.cells.get(address, SimpleNamespace(value=None, data_type="n"))


class FakeWorkbook:
    sheetnames = ["Sheet1"]

    def __init__(self, cells):
        self.sheet = FakeSheet(cells)
        self.closed = False

    def __getitem__(self, sheet_name):
        return self.sheet

    def close(self):
        self.closed = True


class ReadOutputNamesTests(unittest.TestCase):
    def read(self, formula_cells, value_cells, start="A1", page_count=2):
        formula_book = FakeWorkbook(formula_cells)
        value_book = FakeWorkbook(value_cells)
        with patch.object(app, "load_workbook", side_effect=[formula_book, value_book]) as loader:
            names = app.read_output_names(Path("data.xlsx"), "Sheet1", start, page_count)
        self.assertEqual(loader.call_args_list, [
            call(Path("data.xlsx"), read_only=True, data_only=False),
            call(Path("data.xlsx"), read_only=True, data_only=True),
        ])
        self.assertTrue(formula_book.closed)
        self.assertTrue(value_book.closed)
        return names

    def test_reads_downward_from_selected_start_cell_without_sequence(self):
        cells = {
            "B3": SimpleNamespace(value="あいう", data_type="s"),
            "B4": SimpleNamespace(value="かきく", data_type="s"),
        }
        self.assertEqual(self.read(cells, cells, start="B3"), ["あいう.pdf", "かきく.pdf"])

    def test_uses_cached_formula_results(self):
        formulas = {
            "A1": SimpleNamespace(value="=\"あいう\"", data_type="f"),
            "A2": SimpleNamespace(value="=\"かきく\"", data_type="f"),
        }
        values = {
            "A1": SimpleNamespace(value="あいう", data_type="s"),
            "A2": SimpleNamespace(value="かきく", data_type="s"),
        }
        self.assertEqual(self.read(formulas, values), ["あいう.pdf", "かきく.pdf"])

    def test_missing_formula_cache_prompts_excel_recalculation(self):
        formulas = {"A1": SimpleNamespace(value="=B1", data_type="f")}
        values = {"A1": SimpleNamespace(value=None, data_type="n")}
        formula_book = FakeWorkbook(formulas)
        value_book = FakeWorkbook(values)
        with patch.object(app, "load_workbook", side_effect=[formula_book, value_book]):
            with self.assertRaisesRegex(ValueError, "Excelでブックを開き、数式を再計算してから保存"):
                app.read_output_names(Path("data.xlsx"), "Sheet1", "A1", 1)
        self.assertTrue(formula_book.closed)
        self.assertTrue(value_book.closed)

    def test_empty_cell_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "が空です"):
            self.read({}, {}, page_count=1)

    def test_duplicate_names_are_rejected_case_insensitively(self):
        cells = {
            "A1": SimpleNamespace(value="Report", data_type="s"),
            "A2": SimpleNamespace(value="report", data_type="s"),
        }
        with self.assertRaisesRegex(ValueError, "同じファイル名"):
            self.read(cells, cells)

    def test_invalid_windows_filename_is_rejected(self):
        cells = {"A1": SimpleNamespace(value="bad/name", data_type="s")}
        with self.assertRaisesRegex(ValueError, "ファイル名に使えない文字"):
            self.read(cells, cells, page_count=1)

    def test_splits_real_pdf_and_uses_excel_values_as_filenames(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path = root / "input.pdf"
            excel_path = root / "names.xlsx"
            output_dir = root / "output"
            output_dir.mkdir()

            writer = PdfWriter()
            for _ in range(3):
                writer.add_blank_page(width=72, height=72)
            with pdf_path.open("wb") as stream:
                writer.write(stream)

            workbook = Workbook()
            sheet = workbook.active
            sheet["A1"] = "あいう"
            sheet["A2"] = "かきく"
            sheet["A3"] = "さしす"
            workbook.save(excel_path)
            workbook.close()

            results = app.split_pdf_by_cell(pdf_path, excel_path, output_dir, "Sheet", "A1")

            self.assertEqual([path.name for path in results], ["あいう.pdf", "かきく.pdf", "さしす.pdf"])
            self.assertEqual([len(PdfReader(str(path)).pages) for path in results], [1, 1, 1])


if __name__ == "__main__":
    unittest.main()

