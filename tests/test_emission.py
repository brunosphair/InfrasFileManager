import unittest
from unittest.mock import patch, MagicMock
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from InfrasEmission import Emission


def make_emission(**kwargs):
    """Cria uma instância de Emission sem chamar __init__, evitando efeitos colaterais
    de filesystem, GUI e Excel."""
    obj = object.__new__(Emission)
    obj.doc_reg_expression = r'^IFS-\d{4}-\d{3}-(?:\w{3}|\w)-\w{2}-\d{4,5}.*(_R\d{1,2})?$'
    obj.rev_reg_expression = r'(?i)_R\d+$'
    obj.file_num_caract = 23
    for k, v in kwargs.items():
        setattr(obj, k, v)
    return obj


# ---------------------------------------------------------------------------
# verify_pattern
# ---------------------------------------------------------------------------

class TestVerifyPattern(unittest.TestCase):
    def setUp(self):
        self.emis = make_emission()

    def test_valido_sem_revisao_disciplina_1_letra(self):
        self.assertTrue(self.emis.verify_pattern('IFS-2227-001-A-AR-00001.pdf'))

    def test_valido_com_revisao_disciplina_1_letra(self):
        self.assertTrue(self.emis.verify_pattern('IFS-2227-001-A-AR-00001_R1.pdf'))

    def test_valido_disciplina_3_letras(self):
        self.assertTrue(self.emis.verify_pattern('IFS-2227-001-GER-LD-0001.pdf'))

    def test_invalido_disciplina_2_letras(self):
        # (?:\w{3}|\w) aceita apenas 1 ou 3 caracteres — 2 deve ser rejeitado
        self.assertFalse(self.emis.verify_pattern('IFS-2227-001-GE-LD-00001.pdf'))

    def test_invalido_sem_prefixo_ifs(self):
        self.assertFalse(self.emis.verify_pattern('documento.pdf'))

    def test_invalido_numero_projeto_com_poucos_digitos(self):
        self.assertFalse(self.emis.verify_pattern('IFS-22-001-A-AR-00001.pdf'))


# ---------------------------------------------------------------------------
# verify_date_pattern
# ---------------------------------------------------------------------------

class TestVerifyDatePattern(unittest.TestCase):
    def test_data_valida(self):
        self.assertTrue(Emission.verify_date_pattern('01/04/26'))

    def test_invalida_digito_unico(self):
        self.assertFalse(Emission.verify_date_pattern('1/4/26'))

    def test_invalida_string_vazia(self):
        self.assertFalse(Emission.verify_date_pattern(''))

    def test_none_levanta_type_error(self):
        # re.match não aceita None — comportamento atual do código
        with self.assertRaises(TypeError):
            Emission.verify_date_pattern(None)


# ---------------------------------------------------------------------------
# verify_ld_pattern_no_rev
# ---------------------------------------------------------------------------

class TestVerifyLdPatternNoRev(unittest.TestCase):
    def test_valido_sem_revisao(self):
        self.assertTrue(Emission.verify_ld_pattern_no_rev('IFS-2227-001-GER-LD-0001'))

    def test_valido_com_r0(self):
        self.assertTrue(Emission.verify_ld_pattern_no_rev('IFS-2227-001-GER-LD-0001_R0'))

    def test_invalido_disciplina_2_letras(self):
        self.assertFalse(Emission.verify_ld_pattern_no_rev('IFS-2227-001-GE-LD-00001'))

    def test_invalido_nome_generico(self):
        self.assertFalse(Emission.verify_ld_pattern_no_rev('documento'))


# ---------------------------------------------------------------------------
# get_folder_name
# ---------------------------------------------------------------------------

class TestGetFolderName(unittest.TestCase):
    def test_corta_nos_23_primeiros_caracteres(self):
        self.assertEqual(
            Emission.get_folder_name('IFS-2227-001-A-AR-00001_R2.pdf', 23),
            'IFS-2227-001-A-AR-00001'
        )

    def test_sem_revisao_corta_nos_23_primeiros_caracteres(self):
        self.assertEqual(
            Emission.get_folder_name('IFS-2227-001-A-AR-00001.pdf', 23),
            'IFS-2227-001-A-AR-00001'
        )

    def test_corta_nos_24_primeiros_caracteres_disciplina_3_letras(self):
        self.assertEqual(
            Emission.get_folder_name('IFS-2227-001-GER-LD-0001_R0.xlsx', 23),
            'IFS-2227-001-GER-LD-0001'
        )

    def test_sem_revisao_disciplina_3_letras(self):
        self.assertEqual(
            Emission.get_folder_name('IFS-2227-001-GER-LD-0001.xlsx', 23),
            'IFS-2227-001-GER-LD-0001'
        )


# ---------------------------------------------------------------------------
# get_revision
# ---------------------------------------------------------------------------

class TestGetRevision(unittest.TestCase):
    def setUp(self):
        self.emis = make_emission()

    def test_revisao_2(self):
        self.assertEqual(self.emis.get_revision('IFS-2227-001-A-AR-00001_R2.pdf'), 2)

    def test_sem_revisao_retorna_0(self):
        self.assertEqual(self.emis.get_revision('IFS-2227-001-A-AR-00001.pdf'), 0)

    def test_revisao_0(self):
        self.assertEqual(self.emis.get_revision('IFS-2227-001-A-AR-00001_R0.pdf'), 0)

    def test_revisao_15(self):
        self.assertEqual(self.emis.get_revision('IFS-2227-001-A-AR-00001_R15.pdf'), 15)


# ---------------------------------------------------------------------------
# get_file_name
# ---------------------------------------------------------------------------

class TestGetFileName(unittest.TestCase):
    def setUp(self):
        self.emis = make_emission()

    def test_remove_revisao_e_extensao(self):
        self.assertEqual(
            self.emis.get_file_name('IFS-2227-001-A-AR-00001_R2.pdf'),
            'IFS-2227-001-A-AR-00001'
        )

    def test_sem_revisao_remove_apenas_extensao(self):
        self.assertEqual(
            self.emis.get_file_name('IFS-2227-001-A-AR-00001.pdf'),
            'IFS-2227-001-A-AR-00001'
        )

    def test_disciplina_3_letras_retorna_24_chars(self):
        self.assertEqual(
            self.emis.get_file_name('IFS-2227-001-GER-LD-0001_R0.xlsx'),
            'IFS-2227-001-GER-LD-0001'
        )


# ---------------------------------------------------------------------------
# get_ld_revision
# ---------------------------------------------------------------------------

class TestGetLdRevision(unittest.TestCase):
    def test_ld_com_revisao_3(self):
        self.assertEqual(
            Emission.get_ld_revision('IFS-2227-001-GER-LD-0001_R3.xlsx'), 3
        )

    def test_nome_invalido_retorna_menos_1(self):
        self.assertEqual(Emission.get_ld_revision('documento.pdf'), -1)

    def test_placeholder_padrao_retorna_menos_1(self):
        # Nome padrão usado quando nenhuma LD é encontrada
        self.assertEqual(Emission.get_ld_revision('IFS-XXXX-XXX-X-LD-XXXX.xlsx'), -1)


# ---------------------------------------------------------------------------
# get_reg_expressions
# ---------------------------------------------------------------------------

class TestGetRegExpressions(unittest.TestCase):
    def setUp(self):
        self.emis = make_emission()

    @patch('InfrasEmission.load_dotenv')
    @patch('InfrasEmission.os.getenv')
    def test_retorna_valores_do_env(self, mock_getenv, mock_load_dotenv):
        mock_getenv.side_effect = lambda key: {
            'DOC_REG_EXPRESSION': r'^CUSTOM.*$',
            'REV_REG_EXPRESSION': r'_REV\d+$',
        }.get(key)
        doc_re, rev_re = self.emis.get_reg_expressions()
        self.assertEqual(doc_re, r'^CUSTOM.*$')
        self.assertEqual(rev_re, r'_REV\d+$')

    @patch('InfrasEmission.load_dotenv')
    @patch('InfrasEmission.os.getenv', return_value=None)
    def test_retorna_defaults_sem_env(self, mock_getenv, mock_load_dotenv):
        doc_re, rev_re = self.emis.get_reg_expressions()
        self.assertEqual(doc_re, r'^IFS-\d{4}-\d{3}-(?:\w{3}|\w)-\w{2}-\d{4,5}.*(_R\d{1,2})?$')
        self.assertEqual(rev_re, r'(?i)_R\d+$')


# ---------------------------------------------------------------------------
# get_file_num_caract
# ---------------------------------------------------------------------------

class TestGetFileNumCaract(unittest.TestCase):
    def setUp(self):
        self.emis = make_emission()

    @patch('InfrasEmission.os.getenv', return_value='30')
    def test_retorna_inteiro_do_env(self, mock_getenv):
        result = self.emis.get_file_num_caract()
        self.assertEqual(result, 30)
        self.assertIsInstance(result, int)

    @patch('InfrasEmission.os.getenv', return_value=None)
    def test_retorna_default_23_sem_env(self, mock_getenv):
        result = self.emis.get_file_num_caract()
        self.assertEqual(result, 23)


# ---------------------------------------------------------------------------
# get_emited_path
# ---------------------------------------------------------------------------

class TestGetEmitedPath(unittest.TestCase):
    def setUp(self):
        self.emis = make_emission()
        self.original_cwd = os.getcwd()

    def tearDown(self):
        os.chdir(self.original_cwd)

    def test_encontra_3_emitidos(self):
        with tempfile.TemporaryDirectory() as tmp:
            # Estrutura: tmp/projeto/engenharia/emissao  (cwd)
            #            tmp/projeto/3_Emitidos          (destino)
            cwd_path = os.path.join(tmp, 'projeto', 'engenharia', 'emissao')
            emited = os.path.join(tmp, 'projeto', '3_Emitidos')
            os.makedirs(cwd_path)
            os.makedirs(emited)
            os.chdir(cwd_path)
            try:
                result = self.emis.get_emited_path()
            finally:
                # Restaura antes do with fechar — Windows não permite deletar
                # o tempdir enquanto o cwd aponta para dentro dele
                os.chdir(self.original_cwd)
        self.assertEqual(os.path.normpath(result), os.path.normpath(emited))

    def test_levanta_error_sem_3_emitidos(self):
        with tempfile.TemporaryDirectory() as tmp:
            cwd_path = os.path.join(tmp, 'projeto', 'engenharia', 'emissao')
            os.makedirs(cwd_path)
            os.chdir(cwd_path)
            try:
                with self.assertRaises(FileNotFoundError):
                    self.emis.get_emited_path()
            finally:
                os.chdir(self.original_cwd)


# ---------------------------------------------------------------------------
# get_ld_path
# ---------------------------------------------------------------------------

class TestGetLdPath(unittest.TestCase):
    def setUp(self):
        self.emis = make_emission()
        self.original_cwd = os.getcwd()

    def tearDown(self):
        os.chdir(self.original_cwd)

    def test_encontra_lds_dentro_de_emited_path(self):
        with tempfile.TemporaryDirectory() as tmp:
            emited_path = os.path.join(tmp, '3_Emitidos')
            ld_path = os.path.join(emited_path, '_LDs')
            os.makedirs(ld_path)
            self.emis.emited_path = emited_path
            result = self.emis.get_ld_path()
            self.assertEqual(os.path.normpath(result), os.path.normpath(ld_path))

    def test_encontra_00_lds_no_cwd(self):
        with tempfile.TemporaryDirectory() as tmp:
            emited_path = os.path.join(tmp, '3_Emitidos')
            os.makedirs(emited_path)
            cwd_ld = os.path.join(tmp, '00_LDs')
            os.makedirs(cwd_ld)
            os.chdir(tmp)
            self.emis.emited_path = emited_path
            try:
                result = self.emis.get_ld_path()
            finally:
                os.chdir(self.original_cwd)
        self.assertEqual(os.path.normpath(result), os.path.normpath(cwd_ld))

    def test_levanta_error_sem_pasta_ld(self):
        with tempfile.TemporaryDirectory() as tmp:
            emited_path = os.path.join(tmp, '3_Emitidos')
            os.makedirs(emited_path)
            os.chdir(tmp)
            self.emis.emited_path = emited_path
            try:
                with self.assertRaises(FileNotFoundError):
                    self.emis.get_ld_path()
            finally:
                os.chdir(self.original_cwd)


# ---------------------------------------------------------------------------
# get_files
# ---------------------------------------------------------------------------

class TestGetFiles(unittest.TestCase):
    def setUp(self):
        self.emis = make_emission()

    @patch('InfrasEmission.os.walk')
    def test_retorna_lista_correta_de_docs(self, mock_walk):
        mock_walk.return_value = [
            ('.', [], ['IFS-2227-001-A-AR-00001_R1.pdf', 'IFS-2227-001-A-AR-00002_R1.pdf'])
        ]
        docs = self.emis.get_files()
        self.assertEqual(len(docs), 2)
        self.assertEqual(docs[0]['file_name'], 'IFS-2227-001-A-AR-00001_R1.pdf')
        self.assertEqual(docs[0]['rev'], 1)
        self.assertTrue(docs[0]['emit'])

    @patch('InfrasEmission.os.walk')
    @patch('InfrasEmission.msgbox')
    def test_arquivo_duplicado_encerra_programa(self, mock_msgbox, mock_walk):
        mock_walk.return_value = [
            ('.', [], ['IFS-2227-001-A-AR-00001_R1.pdf']),
            ('./subdir', [], ['IFS-2227-001-A-AR-00001_R1.pdf']),
        ]
        with self.assertRaises(SystemExit):
            self.emis.get_files()
        mock_msgbox.assert_called_once()

    @patch('InfrasEmission.os.walk')
    def test_arquivos_ocultos_sao_ignorados(self, mock_walk):
        mock_walk.return_value = [
            ('.', [], ['.arquivo_oculto', 'IFS-2227-001-A-AR-00001_R1.pdf'])
        ]
        docs = self.emis.get_files()
        self.assertEqual(len(docs), 1)
        self.assertEqual(docs[0]['file_name'], 'IFS-2227-001-A-AR-00001_R1.pdf')


# ---------------------------------------------------------------------------
# check_filename_pattern
# ---------------------------------------------------------------------------

class TestCheckFilenamePattern(unittest.TestCase):
    def setUp(self):
        self.emis = make_emission()

    def test_nomes_validos_permanecem_emitaveis(self):
        self.emis.docs = [
            {'file_name': 'IFS-2227-001-A-AR-00001.pdf', 'rev': 0, 'emit': True, 'subdir': '.'},
            {'file_name': 'IFS-2227-001-A-AR-00002_R1.pdf', 'rev': 1, 'emit': True, 'subdir': '.'},
        ]
        with patch.object(self.emis, 'text_box') as mock_text_box:
            self.emis.check_filename_pattern()
        for doc in self.emis.docs:
            self.assertTrue(doc['emit'])
        mock_text_box.assert_not_called()

    def test_nome_invalido_marca_emit_false_e_exibe_aviso(self):
        self.emis.docs = [
            {'file_name': 'IFS-2227-001-A-AR-00001.pdf', 'rev': 0, 'emit': True, 'subdir': '.'},
            {'file_name': 'documento_invalido.pdf', 'rev': 0, 'emit': True, 'subdir': '.'},
        ]
        with patch.object(self.emis, 'text_box') as mock_text_box:
            self.emis.check_filename_pattern()
        self.assertTrue(self.emis.docs[0]['emit'])
        self.assertFalse(self.emis.docs[1]['emit'])
        mock_text_box.assert_called_once()


# ---------------------------------------------------------------------------
# issued_directories
# ---------------------------------------------------------------------------

class TestIssuedDirectories(unittest.TestCase):
    def setUp(self):
        self.emis = make_emission()
        self.emis.emited_path = '/fake/3_Emitidos'
        self.emis.directories = {}

    def test_novo_doc_adicionado_a_dirs_to_create(self):
        self.emis.docs = [
            {'file_name': 'IFS-2227-001-A-AR-00001_R1.pdf', 'rev': 1, 'emit': True, 'subdir': '.'},
        ]
        dirs = self.emis.issued_directories()
        self.assertIn('IFS-2227-001-A-AR-00001', dirs)

    def test_pasta_existente_chama_check_file(self):
        self.emis.directories = {'IFS-2227-001-A-AR-00001': '.'}
        self.emis.docs = [
            {'file_name': 'IFS-2227-001-A-AR-00001_R2.pdf', 'rev': 2, 'emit': True, 'subdir': '.'},
        ]
        with patch.object(self.emis, 'check_file') as mock_check:
            self.emis.issued_directories()
        mock_check.assert_called_once()


# ---------------------------------------------------------------------------
# confirm_files
# ---------------------------------------------------------------------------

class TestConfirmFiles(unittest.TestCase):
    def setUp(self):
        self.emis = make_emission()
        self.emis.grd_name = 'IFS-GRD-2227-001'

    @patch('InfrasEmission.multchoicebox')
    def test_arquivo_desmarcado_recebe_emit_false(self, mock_multchoice):
        self.emis.docs = [
            {'file_name': 'IFS-2227-001-A-AR-00001.pdf', 'rev': 0, 'emit': True, 'subdir': '.'},
            {'file_name': 'IFS-2227-001-A-AR-00002.pdf', 'rev': 0, 'emit': True, 'subdir': '.'},
        ]
        mock_multchoice.return_value = ['IFS-2227-001-A-AR-00001.pdf']
        self.emis.confirm_files({})
        self.assertTrue(self.emis.docs[0]['emit'])
        self.assertFalse(self.emis.docs[1]['emit'])

    @patch('InfrasEmission.ccbox')
    def test_unico_arquivo_exibe_ccbox(self, mock_ccbox):
        self.emis.docs = [
            {'file_name': 'IFS-2227-001-A-AR-00001.pdf', 'rev': 0, 'emit': True, 'subdir': '.'},
        ]
        self.emis.confirm_files({})
        mock_ccbox.assert_called_once()

    @patch('InfrasEmission.msgbox')
    def test_sem_arquivos_emitaveis_encerra_programa(self, mock_msgbox):
        self.emis.docs = []
        with self.assertRaises(SystemExit):
            self.emis.confirm_files({})
        mock_msgbox.assert_called_once()


# ---------------------------------------------------------------------------
# duplicated_file
# ---------------------------------------------------------------------------

class TestDuplicatedFile(unittest.TestCase):
    def _make_doc(self):
        return {'file_name': 'IFS-2227-001-A-AR-00001_R1.pdf', 'rev': 1, 'emit': True}

    @patch('InfrasEmission.buttonbox', return_value='Não emitir esse arquivo')
    def test_nao_emitir_define_emit_false(self, mock_btn):
        doc = self._make_doc()
        Emission.duplicated_file(
            'IFS-2227-001-A-AR-00001', doc, '/fake/dir', 'IFS-2227-001-A-AR-00001_R1.pdf'
        )
        self.assertFalse(doc['emit'])

    @patch('InfrasEmission.buttonbox', return_value='Cancelar')
    def test_cancelar_encerra_programa(self, mock_btn):
        doc = self._make_doc()
        with self.assertRaises(SystemExit):
            Emission.duplicated_file(
                'IFS-2227-001-A-AR-00001', doc, '/fake/dir', 'IFS-2227-001-A-AR-00001_R1.pdf'
            )

    @patch('InfrasEmission.buttonbox', return_value='Emitir mesmo assim')
    @patch('InfrasEmission.os.replace')
    @patch('InfrasEmission.os.mkdir')
    @patch('InfrasEmission.os.path.isfile', return_value=False)
    @patch('InfrasEmission.os.path.isdir', return_value=False)
    def test_emitir_mesmo_assim_cria_obsoleto_e_move(
            self, mock_isdir, mock_isfile, mock_mkdir, mock_replace, mock_btn):
        doc = self._make_doc()
        Emission.duplicated_file(
            'IFS-2227-001-A-AR-00001', doc, '/fake/dir', 'IFS-2227-001-A-AR-00001_R1.pdf'
        )
        self.assertTrue(doc['emit'])
        mock_mkdir.assert_called_once()
        mock_replace.assert_called_once()


if __name__ == '__main__':
    unittest.main()
