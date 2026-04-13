import unittest
from unittest.mock import patch, MagicMock
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_functions import (
    get_cover_cell,
    get_grd_number,
    get_acronym_default_list,
    copy_values,
    reorder_description_cells,
)


# ---------------------------------------------------------------------------
# get_cover_cell
# ---------------------------------------------------------------------------

class TestGetCoverCell(unittest.TestCase):
    def test_rev_0(self):
        self.assertEqual(get_cover_cell(0), [32, 3])

    def test_rev_1(self):
        self.assertEqual(get_cover_cell(1), [32, 5])

    def test_rev_2(self):
        self.assertEqual(get_cover_cell(2), [32, 7])

    def test_rev_3(self):
        self.assertEqual(get_cover_cell(3), [32, 8])

    def test_rev_4(self):
        self.assertEqual(get_cover_cell(4), [32, 11])

    def test_rev_5(self):
        self.assertEqual(get_cover_cell(5), [37, 3])

    def test_rev_6(self):
        self.assertEqual(get_cover_cell(6), [37, 5])

    def test_rev_7(self):
        self.assertEqual(get_cover_cell(7), [37, 7])

    def test_rev_8(self):
        self.assertEqual(get_cover_cell(8), [37, 8])

    def test_rev_9(self):
        self.assertEqual(get_cover_cell(9), [37, 11])

    def test_rev_15(self):
        self.assertEqual(get_cover_cell(15), [37, 11])


# ---------------------------------------------------------------------------
# get_grd_number
# ---------------------------------------------------------------------------

class TestGetGrdNumber(unittest.TestCase):
    @patch('excel_functions.openpyxl.load_workbook')
    def test_duas_abas_existentes_retorna_3(self, mock_load_wb):
        mock_wb = MagicMock()
        mock_wb.sheetnames = ['Capa', 'GRD-XXX', 'GRD-001', 'GRD-002']
        mock_load_wb.return_value = mock_wb
        result = get_grd_number('/fake/path', 'ld_file.xlsx')
        self.assertEqual(result, 3)

    @patch('excel_functions.openpyxl.load_workbook')
    def test_sem_abas_grd_retorna_1(self, mock_load_wb):
        mock_wb = MagicMock()
        mock_wb.sheetnames = ['Capa', 'GRD-XXX']
        mock_load_wb.return_value = mock_wb
        result = get_grd_number('/fake/path', 'ld_file.xlsx')
        self.assertEqual(result, 1)


# ---------------------------------------------------------------------------
# get_acronym_default_list
# ---------------------------------------------------------------------------

class TestGetAcronymDefaultList(unittest.TestCase):
    @patch('excel_functions.openpyxl.load_workbook')
    def test_retorna_siglas_das_celulas_corretas(self, mock_load_wb):
        mock_wb = MagicMock()
        mock_sheet = MagicMock()
        mock_wb.__getitem__ = MagicMock(return_value=mock_sheet)

        # previous_cover_cell = [32, 3]  →  lê linhas 33, 34, 35 / coluna 3
        def mock_cell(row, column):
            values = {(33, 3): 'ABC', (34, 3): 'DEF', (35, 3): 'GHI'}
            cell = MagicMock()
            cell.value = values.get((row, column))
            return cell

        mock_sheet.cell = mock_cell
        mock_load_wb.return_value = mock_wb

        result = get_acronym_default_list('/fake/ld.xlsx', [32, 3])
        self.assertEqual(result, ['ABC', 'DEF', 'GHI'])


# ---------------------------------------------------------------------------
# copy_values
# ---------------------------------------------------------------------------

class TestCopyValues(unittest.TestCase):
    def test_copia_valor_da_origem_para_destino(self):
        source_cell = MagicMock()
        source_cell.value = 'valor_teste'
        dest_cell = MagicMock()

        mock_sheet = MagicMock()
        mock_sheet.cell.side_effect = (
            lambda row, column: source_cell if (row, column) == (1, 1) else dest_cell
        )

        copy_values(mock_sheet, 1, 1, 2, 2)

        self.assertEqual(dest_cell.value, 'valor_teste')


# ---------------------------------------------------------------------------
# reorder_description_cells
# ---------------------------------------------------------------------------

class TestReorderDescriptionCells(unittest.TestCase):
    def test_desloca_13_linhas_em_3_colunas(self):
        call_args = []

        def mock_copy(sheet, from_row, from_col, to_row, to_col):
            call_args.append((from_row, from_col, to_row, to_col))

        mock_sheet = MagicMock()

        with patch('excel_functions.copy_values', side_effect=mock_copy):
            reorder_description_cells(mock_sheet)

        # 13 linhas × 3 colunas = 39 chamadas
        self.assertEqual(len(call_args), 39)
        # Primeira chamada: linha 18 → 17, coluna 1
        self.assertIn((18, 1, 17, 1), call_args)
        # Última chamada: linha 30 → 29, coluna 3
        self.assertIn((30, 3, 29, 3), call_args)


if __name__ == '__main__':
    unittest.main()
