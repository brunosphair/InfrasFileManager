import os
import sys
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from InfrasEmission import Emission
from excel_functions import create_excel_grd

FIXTURE_PATH = os.path.join(os.path.dirname(__file__), 'fixtures', 'IFS-XXXX-XXX-X-LD-XXXX.xlsx')


def make_emission(**kwargs):
    obj = object.__new__(Emission)
    obj.doc_reg_expression = r'^IFS-\d{4}-\d{3}-(?:\w{3}|\w)-\w{2}-\d{4,5}.*(_R\d{1,2})?$'
    obj.rev_reg_expression = r'(?i)_R\d+$'
    obj.file_num_caract = 23
    for k, v in kwargs.items():
        setattr(obj, k, v)
    return obj


# ---------------------------------------------------------------------------
# TestMoveFiles — arquivo é movido para a pasta correta em 3_Emitidos
# ---------------------------------------------------------------------------

class TestMoveFiles(unittest.TestCase):
    def setUp(self):
        self.original_cwd = os.getcwd()

    def tearDown(self):
        os.chdir(self.original_cwd)

    @patch('InfrasEmission.msgbox')
    def test_arquivo_vai_para_pasta_correta(self, mock_msgbox):
        with tempfile.TemporaryDirectory() as tmp:
            doc_folder = os.path.join(tmp, '3_Emitidos', 'IFS-2227-001-A-AR-00001')
            os.makedirs(doc_folder)

            filename = 'IFS-2227-001-A-AR-00001_R2.pdf'
            open(os.path.join(tmp, filename), 'w').close()

            os.chdir(tmp)
            try:
                emis = make_emission(
                    base_path=Path(tmp),
                    emited_path=os.path.join(tmp, '3_Emitidos'),
                    directories={'IFS-2227-001-A-AR-00001': '.'},
                    docs=[{
                        'file_name': filename,
                        'rev': 2,
                        'emit': True,
                        'subdir': '.',
                    }],
                )
                emis.move_files()
                expected = os.path.join(tmp, '3_Emitidos', 'IFS-2227-001-A-AR-00001', filename)
                self.assertTrue(os.path.exists(expected))
                mock_msgbox.assert_called_once()
            finally:
                os.chdir(self.original_cwd)

    @patch('InfrasEmission.msgbox')
    def test_arquivo_com_emit_false_nao_e_movido(self, mock_msgbox):
        with tempfile.TemporaryDirectory() as tmp:
            doc_folder = os.path.join(tmp, '3_Emitidos', 'IFS-2227-001-A-AR-00001')
            os.makedirs(doc_folder)

            filename = 'IFS-2227-001-A-AR-00001_R2.pdf'
            src = os.path.join(tmp, filename)
            open(src, 'w').close()

            os.chdir(tmp)
            try:
                emis = make_emission(
                    emited_path=os.path.join(tmp, '3_Emitidos'),
                    directories={'IFS-2227-001-A-AR-00001': '.'},
                    docs=[{
                        'file_name': filename,
                        'rev': 2,
                        'emit': False,
                        'subdir': '.',
                    }],
                )
                emis.move_files()
                self.assertTrue(os.path.exists(src))
                dest = os.path.join(tmp, '3_Emitidos', 'IFS-2227-001-A-AR-00001', filename)
                self.assertFalse(os.path.exists(dest))
            finally:
                os.chdir(self.original_cwd)


# ---------------------------------------------------------------------------
# TestCreateZip — ZIP é criado com os arquivos corretos
# ---------------------------------------------------------------------------

class TestCreateZip(unittest.TestCase):
    def setUp(self):
        self.original_cwd = os.getcwd()

    def tearDown(self):
        os.chdir(self.original_cwd)

    def test_zip_e_criado_com_arquivos_emit_true(self):
        with tempfile.TemporaryDirectory() as tmp:
            file1 = 'IFS-2227-001-A-AR-00001_R2.pdf'
            file2 = 'IFS-2227-001-A-AR-00002_R0.pdf'
            open(os.path.join(tmp, file1), 'w').close()
            open(os.path.join(tmp, file2), 'w').close()

            os.chdir(tmp)
            try:
                emis = make_emission(
                    base_path=Path(tmp),
                    grd_name='IFS-GRD-2227-001',
                    docs=[
                        {'file_name': file1, 'rev': 2, 'emit': True,  'subdir': '.'},
                        {'file_name': file2, 'rev': 0, 'emit': False, 'subdir': '.'},
                    ],
                )
                emis.create_zip()
                zip_path = os.path.join(tmp, 'IFS-GRD-2227-001.zip')
                self.assertTrue(os.path.exists(zip_path))

                with ZipFile(zip_path) as z:
                    names = z.namelist()
            finally:
                os.chdir(self.original_cwd)

        self.assertIn(file1, names)
        self.assertNotIn(file2, names)

    def test_zip_sem_arquivos_emitidos_esta_vazio(self):
        with tempfile.TemporaryDirectory() as tmp:
            os.chdir(tmp)
            try:
                emis = make_emission(
                    base_path=Path(tmp),
                    grd_name='IFS-GRD-2227-002',
                    docs=[
                        {'file_name': 'doc.pdf', 'rev': 0, 'emit': False, 'subdir': '.'},
                    ],
                )
                emis.create_zip()
                zip_path = os.path.join(tmp, 'IFS-GRD-2227-002.zip')
                self.assertTrue(os.path.exists(zip_path))

                with ZipFile(zip_path) as z:
                    names = z.namelist()
            finally:
                os.chdir(self.original_cwd)

        self.assertEqual(names, [])


# ---------------------------------------------------------------------------
# TestCreateExcelGrd — planilha .xlsm é criada na pasta LD
# Requer: tests/fixtures/template_ld.xlsm (LD real com abas GRD-XXX e Capa)
#         e Excel instalado (xlwings)
# ---------------------------------------------------------------------------

@unittest.skipUnless(
    os.path.exists(FIXTURE_PATH) and sys.platform == 'win32',
    "Requer Windows + Excel instalado + fixture tests/fixtures/IFS-XXXX-XXX-X-LD-XXXX.xlsx",
)
class TestCreateExcelGrd(unittest.TestCase):
    def test_cria_arquivo_xlsx_na_pasta_ld(self):
        ld_output_name = 'IFS-2227-001-GER-LD-0001'
        ld_information = {
            'emission_date': '01/04/26',
            'ld_name': ld_output_name,
            'project_title': 'PROJETO TESTE',
            'ld_title': 'PROJETO TESTE\nLISTA DE DOCUMENTOS',
            'acronym1': 'ABC',
            'acronym2': 'DEF',
            'acronym3': 'GHI',
        }
        grd_items = [('IFS-2227-001-A-AR-00001', 0)]

        with tempfile.TemporaryDirectory() as tmp:
            fixture_name = os.path.basename(FIXTURE_PATH)
            shutil.copy(FIXTURE_PATH, os.path.join(tmp, fixture_name))

            create_excel_grd(
                ld_path=tmp,
                ld_name=fixture_name,
                grd_number=1,
                grd_name='IFS-GRD-2227-001',
                ld_information=ld_information,
                ld_rev=-1,
                grd_items=grd_items,
            )

            expected_file = os.path.join(tmp, ld_output_name + '_R0.xlsx')
            self.assertTrue(
                os.path.exists(expected_file),
                f"Arquivo esperado não encontrado: {expected_file}",
            )

    def test_cria_arquivo_xlsx_disciplina_1_letra(self):
        ld_output_name = 'IFS-2227-001-A-LD-00001'
        ld_information = {
            'emission_date': '01/04/26',
            'ld_name': ld_output_name,
            'project_title': 'PROJETO TESTE',
            'ld_title': 'PROJETO TESTE\nLISTA DE DOCUMENTOS',
            'acronym1': 'ABC',
            'acronym2': 'DEF',
            'acronym3': 'GHI',
        }
        grd_items = [('IFS-2227-001-A-AR-00001', 0)]

        with tempfile.TemporaryDirectory() as tmp:
            fixture_name = os.path.basename(FIXTURE_PATH)
            shutil.copy(FIXTURE_PATH, os.path.join(tmp, fixture_name))

            create_excel_grd(
                ld_path=tmp,
                ld_name=fixture_name,
                grd_number=1,
                grd_name='IFS-GRD-2227-001',
                ld_information=ld_information,
                ld_rev=-1,
                grd_items=grd_items,
            )

            expected_file = os.path.join(tmp, ld_output_name + '_R0.xlsx')
            self.assertTrue(
                os.path.exists(expected_file),
                f"Arquivo esperado não encontrado: {expected_file}",
            )

    def test_cria_arquivo_xlsx_segunda_emissao(self):
        ld_output_name = 'IFS-2227-001-GER-LD-0001'
        grd_items = [('IFS-2227-001-A-AR-00001', 0)]

        with tempfile.TemporaryDirectory() as tmp:
            fixture_name = os.path.basename(FIXTURE_PATH)
            shutil.copy(FIXTURE_PATH, os.path.join(tmp, fixture_name))

            # 1ª emissão — cria R0 a partir da fixture template
            create_excel_grd(
                ld_path=tmp,
                ld_name=fixture_name,
                grd_number=1,
                grd_name='IFS-GRD-2227-001',
                ld_information={
                    'emission_date': '01/04/26',
                    'ld_name': ld_output_name,
                    'project_title': 'PROJETO TESTE',
                    'ld_title': 'PROJETO TESTE\nLISTA DE DOCUMENTOS',
                    'acronym1': 'ABC',
                    'acronym2': 'DEF',
                    'acronym3': 'GHI',
                },
                ld_rev=-1,
                grd_items=grd_items,
            )

            r0_name = ld_output_name + '_R0.xlsx'
            self.assertTrue(
                os.path.exists(os.path.join(tmp, r0_name)),
                f"R0 não foi criado: {r0_name}",
            )

            # 2ª emissão — usa o R0 criado (última planilha) como base
            create_excel_grd(
                ld_path=tmp,
                ld_name=r0_name,
                grd_number=2,
                grd_name='IFS-GRD-2227-002',
                ld_information={
                    'emission_date': '01/05/26',
                    'acronym1': 'ABC',
                    'acronym2': 'DEF',
                    'acronym3': 'GHI',
                },
                ld_rev=0,
                grd_items=grd_items,
            )

            base = Emission.get_folder_name(r0_name, 23)
            expected_r1 = os.path.join(tmp, base + '_R1.xlsx')
            self.assertTrue(
                os.path.exists(expected_r1),
                f"Arquivo R1 esperado não encontrado: {expected_r1}",
            )


if __name__ == '__main__':
    unittest.main()
