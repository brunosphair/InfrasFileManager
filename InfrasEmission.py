import os
import re
import sys
import stat
from pathlib import Path
from easygui import buttonbox, ccbox, multchoicebox, enterbox, msgbox, \
                    multenterbox
import tkinter as tk
from tkinter import filedialog
from zipfile import ZipFile
import datetime
from dotenv import load_dotenv
import fitz

from excel_functions import get_grd_number, create_excel_grd, \
                            get_acronym_default_list, get_cover_cell



class Emission:
    def __init__(self, folder_path=None):
        load_dotenv()
        self.base_path = Path(folder_path).absolute() if folder_path else Path(os.getcwd()).absolute()
        self.doc_reg_expression, self.rev_reg_expression = \
                                                    self.get_reg_expressions()
        self.file_num_caract = self.get_file_num_caract()
        self.emited_path = self.get_emited_path()
        self.ld_path = self.get_ld_path()
        self.docs = self.get_files()
        self.directories = self.get_emited_directories()
        self.ld_rev = self.get_ld_rev()
        self.project_number = self.get_project_number()
        self.grd_number = get_grd_number(self.ld_path, self.ld_name)
        self.grd_name = 'IFS-GRD-' + \
                        str(self.project_number) + \
                        "-" + str(self.grd_number).zfill(3)
        self.ld_information = {}


    def get_files(self):
        '''
        Returns a list of dictionarys, where which dict has the filename, the
        revision number and declares the 'emit' key as True.
        '''
        docs = []
        file_names = []
        for path, subdir, files in os.walk(self.base_path):
            subdir[:] = [d for d in subdir if d != '00_LDs']
            # subdir.clear()
            for file in files:

                full_path = os.path.join(path, file)

                # Ignorar arquivos ocultos
                if file.startswith('.') or self.is_hidden(full_path):
                    continue

                if file not in file_names:
                    file_names.append(file)
                    dict_item = {}
                    rev = self.get_revision(file)
                    dict_item['file_name'] = file
                    dict_item['rev'] = rev
                    dict_item['emit'] = True
                    dict_item['subdir'] = os.path.relpath(path, self.base_path)
                    docs.append(dict_item)
                else:
                    msg = f"Há dois arquivos com o nome {file} dentro da emissão"
                    title = "ERRO"
                    msgbox(msg, title)
                    sys.exit(0)
        return docs

    def get_emited_path(self):
        '''
        Returns the path to the folder 3_Emitidos
        '''
        path = self.base_path.parent
        parent_path = path.parent
        issued_path = os.path.join(parent_path, '3_Emitidos')
        if not os.path.isdir(issued_path):
            raise FileNotFoundError("A pasta 3_Emitidos não foi encontrada")
        return issued_path
    
    def get_ld_path(self):
        '''
        Returns the path to the folder containing the LDs.
        '''
        ld_path = os.path.join(self.emited_path, '_LDs')
        if not os.path.isdir(ld_path):
            ld_path = os.path.join(str(self.base_path), "00_LDs")
        if not os.path.isdir(ld_path):
            raise FileNotFoundError("A pasta de LDs não foi encontrada")
        return ld_path

    def get_emited_directories(self):
        '''
        Return all the directories in the 3_Emitidos path. Therefore, returns
        the name of the files which was alredy emited.
        '''

        directories = {}
        for path, subdirs, files in os.walk(self.emited_path):
            for subdir in subdirs:
                                      
                directories[subdir] = os.path.relpath(path, self.emited_path)

        return directories

    def get_ld_rev(self):
        '''
        Returns the revision of the last LD emited. If this is the first LD,
        then the function returns -1.
        '''
        
        lds = os.listdir(self.ld_path)
        self.ld_name = 'IFS-XXXX-XXX-X-LD-XXXX.xlsx'
        last_revision = -1
        for item in lds:
            ld_revision = self.get_ld_revision(item)
            if last_revision < ld_revision:
                last_revision = ld_revision
                self.ld_name = item

        return last_revision

    def get_reg_expressions(self):

        load_dotenv()
        doc_reg_expression = os.getenv("DOC_REG_EXPRESSION")
        rev_reg_expression = os.getenv("REV_REG_EXPRESSION")

        if doc_reg_expression is None:
            doc_reg_expression = \
                    r'^IFS-\d{4}-\d{3}-(\w{3}-\w{2}-\d{4}[^_\s]*|\w-\w{2}-\d{5}[^_\s]*)(_R\d+)?$'
        if rev_reg_expression is None:
            rev_reg_expression = r'(?i)_R\d+$'

        return doc_reg_expression, rev_reg_expression

    def get_file_num_caract(self):
        file_num_caract = os.getenv("FILE_NUM_CARACT")
        if file_num_caract is None:
            file_num_caract = 23
        else:
            file_num_caract = int(file_num_caract)
        
        return file_num_caract

    def get_project_number(self):
        '''
        Retuns the number of the project.
        '''
        path = self.base_path.parent
        project_path = path.parent
        dir_name = os.path.basename(project_path)
        project_number = dir_name[:4]
        if not project_number.isnumeric():
            raise ValueError("O arquivo executado não está na pasta correta")
        return project_number

    def check_filename_pattern(self):
        '''
        Checks if the name of the files corresponds to the specified pattern.
        If the filename dont correspond to the pattern, the doc['emit'] is
        declared as False, so the file is not going to be emited anymore
        '''
        ignored_files = []
        for doc in self.docs:
            if not self.verify_pattern(doc['file_name']):
                doc['emit'] = False
                if not (doc['file_name'].startswith('InfrasEmission') or
                        doc['file_name'] == '.env'):
                    ignored_files.append(doc['file_name'])

        if len(ignored_files):
            msg = "Os seguintes arquivos não serão emitidos, pois não "\
                "correspondem ao padrão de nomenclatura de arquivos:\n\n" \
                + '\n'.join(ignored_files) + "\n\nO que deseja fazer?"
            title = "Inconsistência na nomenclatura dos arquivos"
            self.text_box(msg, title)

    def check_file(self, doc, folder_name):
        '''
        Checks if the file being issued is already on the issued path. The user
        can choice between cancel the operation, dont issue the doc or issue
        aniway. If the choice was issue anyway, a new folder is created inside
        the doc folder with the name "Obsoleto", and the old file is moved
        inside this folder
        '''
        doc_name = self.get_file_name(doc['file_name'])
        doc_directory = os.path.join(self.emited_path,
                                     self.directories[folder_name],
                                     folder_name)
        if os.path.isdir(doc_directory):
            for file in os.listdir(doc_directory):
                file_name = self.get_file_name(file)
                if self.get_revision(file
                                        ) == doc['rev'] and file_name == doc_name:
                    self.duplicated_file(doc_name, doc, doc_directory, file)

    @staticmethod
    def duplicated_file(doc_name, doc, doc_directory, file):
        msg = 'O arquivo ' + doc_name + ' com a revisão '\
            + str(doc['rev']) \
            + ' já existe. O que deseja fazer?'
        choices = [
                    "Não emitir esse arquivo",
                    "Emitir mesmo assim",
                    "Cancelar"
                    ]
        title = "Arquivo duplicado"
        choice = buttonbox(msg, title, choices)
        if choice == "Não emitir esse arquivo":
            # print("arquivo ignorado")
            doc['emit'] = False
        elif choice == "Emitir mesmo assim":
            obsolete_path = os.path.join(doc_directory, "Obsoleto")
            if not os.path.isdir(obsolete_path):
                os.mkdir(obsolete_path)
            i = 1
            file_aux = file
            while os.path.isfile(os.path.join(obsolete_path, file_aux)):
                file_aux = os.path.splitext(file)[0]\
                    + "(" + str(i) + ")"\
                    + os.path.splitext(file)[0]
                i += 1
            file_source_path = os.path.join(doc_directory, file)
            file_destiny_path = os.path.join(obsolete_path, file_aux)
            os.replace(file_source_path, file_destiny_path)
            doc['emit'] = True
        elif choice == "Cancelar":
            sys.exit(0)

    def confirm_files(self, dirs_to_create):
        list_of_options = []
        for doc in self.docs:
            if doc['emit']:
                list_of_options.append(doc['file_name'])
        if len(list_of_options) == 0:
            msg = "Não há arquivos para serem emitidos."
            title = "Erro"
            msgbox(msg, title)
            sys.exit(0)
        elif len(list_of_options) == 1:
            ccbox("O seguinte arquivo será emitido:\n\n" + list_of_options[0])
        else:
            msg = "Os seguintes arquivos serão emitidos na "\
                + self.grd_name\
                + ". Desmarque caso não queira emitir algum."
            title = 'Deseja continuar?'
            choices = multchoicebox(msg,
                                    title,
                                    list_of_options,
                                    preselect=[*range(len(list_of_options))])
            # print(choices)
            for doc in self.docs:
                if not doc['file_name'] in choices:
                    doc['emit'] = False
                    folder_name = self.get_folder_name(doc['file_name'], self.file_num_caract)
                    if folder_name in dirs_to_create:
                        del dirs_to_create[folder_name]

    def create_zip(self):
        zip_path = os.path.join(str(self.base_path), self.grd_name + '.zip')
        zipObj = ZipFile(zip_path, 'w')
        for doc in self.docs:
            if doc['emit']:
                file_path = os.path.join(str(self.base_path), doc['subdir'], doc['file_name'])
                arcname = os.path.join(doc['subdir'], doc['file_name'])
                zipObj.write(file_path, arcname)
        zipObj.close()

    def create_ld(self):
        no_docs = []
        grd_items = []
        for doc in self.docs:
            doc_name = self.get_file_name(doc['file_name'])
            if doc['emit'] and doc_name not in no_docs:
                no_docs.append(doc_name)
                grd_items.append([doc_name, doc['rev']])

        doc_items = self.get_docx_items()

        create_excel_grd(self.ld_path, self.ld_name, self.grd_number,
                         self.grd_name, self.ld_information, self.ld_rev,
                         grd_items, doc_items)

    # Linha de escala do carimbo (ex.: "1:2000", "1:40.000").
    _ESCALA = re.compile(r'^\d{1,4}:\d[\d.,]*')

    # Palavras que NÃO são título: empreendimento, obra, assinaturas, rótulos.
    _TITULO_RUIDO = ["PONTE", "BRIDGE", "EMPREEND", "GOVERNO", "CONCESS",
                     "ESTUDO DE VIABILIDADE", "EVTEA", "SALVADOR", "ITAPARICA",
                     "ASSINATURA", "ELABORAD", "VERIFIC", "APROVA",
                     "DESCRI", "CLIENTE", "PLANTA-CHAVE", "NOTAS", "LEGENDA"]

    @staticmethod
    def _tenta_layout_A(txt):
        """Layout A = folha A4 (relatório): título na capa ou no carimbo."""
        return Emission._relatorio_por_empreendimento(txt) \
            or Emission._relatorio_por_rev(txt)

    @staticmethod
    def _relatorio_por_empreendimento(txt):
        """Capa bilíngue: a obra aparece 2x; o título vem após o 2º bloco dela."""
        m = re.search(r'EMPREENDIMENTO\b', txt, re.I)
        if not m:
            return None
        lines = [l.strip() for l in txt[m.end():].splitlines()]

        def norm(s):
            return re.sub(r'\s+', ' ',
                          s.replace('–', '-').replace('—', '-')).upper().strip()

        obra_idx = next((i for i, l in enumerate(lines) if len(l) >= 10), None)
        if obra_idx is None:
            return None
        obra = norm(lines[obra_idx])
        second = next((i for i in range(obra_idx + 1, len(lines))
                       if norm(lines[i]) == obra), None)
        if second is None:
            return None
        for l in lines[second + 1:]:
            if not l:
                continue
            if any(x in l.upper() for x in Emission._TITULO_RUIDO):
                continue
            return l
        return None

    @staticmethod
    def _relatorio_por_rev(txt):
        """Carimbo sem EMPREENDIMENTO: título é a linha logo acima de 'REV:'."""
        lines = [l.strip() for l in txt.splitlines()]
        for i, l in enumerate(lines):
            if re.match(r'^REV\.?:?$', l, re.I):
                for j in range(i - 1, -1, -1):
                    if not lines[j]:
                        continue
                    if any(x in lines[j].upper() for x in Emission._TITULO_RUIDO):
                        return None
                    return lines[j]
                return None
        return None

    @staticmethod
    def _tenta_layout_B(txt):
        """Layout B = folha A0/A1 (desenho): título no carimbo."""
        return Emission._desenho_projeto_executivo(txt) \
            or Emission._desenho_por_escala(txt)

    @staticmethod
    def _desenho_projeto_executivo(txt):
        """Carimbo com fase 'PROJETO EXECUTIVO': título nas linhas acima dela."""
        lines = [l.strip() for l in txt.splitlines()]
        pe_indices = [i for i, l in enumerate(lines)
                      if re.match(r'^PROJETO EXECUTIVO$', l, re.I)]
        if not pe_indices:
            return None
        segment = []
        for i in range(pe_indices[-1] - 1, -1, -1):
            line = lines[i]
            if not line:
                continue
            if line == '-':
                break
            segment.insert(0, line)
        if not segment:
            return None
        title_parts = segment[:-1] if len(segment) > 1 else segment
        return ' '.join(title_parts) or None

    @staticmethod
    def _desenho_por_escala(txt):
        """Carimbo sem rótulo de título: título nas linhas acima da escala."""
        lines = [l.strip() for l in txt.splitlines()]
        escala_idx = None
        for i, l in enumerate(lines):
            if Emission._ESCALA.match(l):
                escala_idx = i
        if escala_idx is None:
            return None
        titulo = []
        for i in range(escala_idx - 1, -1, -1):
            l = lines[i]
            if not l:
                continue
            up = l.upper()
            if any(x in up for x in Emission._TITULO_RUIDO):
                break
            if l == '-' or re.fullmatch(r'[\d.,\s/]+', l):
                break
            titulo.insert(0, l)
            if len(titulo) >= 3:
                break
        if not titulo:
            return None
        return ' '.join(titulo).strip(' -\t') or None

    @staticmethod
    def _extract_doc_info(pdf_path):
        doc = fitz.open(pdf_path)
        page = doc[0]
        txt = page.get_text()
        long_edge = max(page.rect.width, page.rect.height)
        doc.close()
        # Roteia pelo tamanho da folha (a nomenclatura do arquivo não importa):
        # A4 (lado maior <= 1000 pt) = relatório; maior que isso = desenho A0/A1.
        if long_edge <= 1000:
            return Emission._tenta_layout_A(txt)
        return Emission._tenta_layout_B(txt)
    
    def get_docx_items(self):
        items = []
        for doc in self.docs:
            if not doc['emit'] or doc['rev'] != 0:
                continue
            name = doc['file_name']
            if not name.lower().endswith('.pdf'):
                continue
            path = os.path.join(str(self.base_path), doc['subdir'], name)
            code = self.get_file_name(name)
            try:
                title = self._extract_doc_info(path)
            except Exception:
                title = None
            items.append((code, title))
        return items

    def get_client_img(self):
        result = [None]

        existing_root = tk._default_root
        if existing_root is not None:
            win = tk.Toplevel(existing_root)
            win.grab_set()
        else:
            win = tk.Tk()

        win.title("Logotipo do cliente")
        win.resizable(False, False)

        tk.Label(win, text="Logotipo do cliente (opcional)",
                 font=("Arial", 11, "bold")).pack(padx=30, pady=(20, 5))

        status_label = tk.Label(win, text="Nenhuma imagem selecionada",
                                font=("Arial", 10), fg="gray")
        status_label.pack(padx=30, pady=(0, 15))

        def importar():
            path = filedialog.askopenfilename(
                parent=win,
                title="Selecione o logotipo do cliente",
                filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp"),
                           ("Todos os arquivos", "*.*")]
            )
            if path:
                result[0] = path
                status_label.config(text=os.path.basename(path), fg="black")

        def continuar():
            win.destroy()

        btn_frame = tk.Frame(win)
        btn_frame.pack(padx=30, pady=(0, 20))
        tk.Button(btn_frame, text="Importar imagem", font=("Arial", 10),
                  width=15, command=importar).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Continuar", font=("Arial", 10),
                  width=15, command=continuar).pack(side=tk.LEFT, padx=5)

        if existing_root is not None:
            existing_root.wait_window(win)
        else:
            win.mainloop()

        return result[0]

    def get_ld_information(self):
        date_defined = False
        while not date_defined:
            title = "Data de emissão"
            text = "Digite a data de emissão da GRD no formato DD/MM/YY:"
            today_date = datetime.datetime.now().strftime("%d/%m/%y")
            emission_date = enterbox(text, title, today_date)
            if self.verify_date_pattern(emission_date):
                date_defined = True
            else:
                msgbox("Formato de data inválido")

        ld_information = {}
        ld_information["emission_date"] = emission_date

        if self.ld_rev == -1:
            text = "Como essa é a primeira emissão desse projeto, digite um "\
                "nome para a LD no padrão IFS-NNNN-NNN-X-LD-NNNNN (disciplina "\
                "com 1 letra e Número do doc com 5 números) ou IFS-NNNN-NNN-XXX-LD-NNNN (disciplina com 3 "\
                "letras e Número do doc com 4 números), onde X são letras e N são números"
            title = "Nomeie a LD"
            probably_name = self.get_probably_name()
            d_text = "IFS-"\
                     + str(self.project_number)\
                     + "-" + probably_name + "-GER-LD-0001"
            defined_name = False
            while not defined_name:
                ld_name = enterbox(text, title, d_text)
                if self.verify_ld_pattern_no_rev(ld_name):
                    defined_name = True
                else:
                    msgbox("O nome que você digitou não atende aos requisitos"
                           " de IFS-NNNN-NNN-XXX-LD-NNNNN, digite novamente",
                           "Nome inválido!")
            ld_information["ld_name"] = ld_name

            text = "Defina os títulos da LD:"
            title = "Definir títulos"
            input_list = ["1ª LINHA - TIPO DE PROJETO ",
                          "2ª LINHA - TÍTULO DO PROJETO",
                          "3ª LINHA - SUBTÍTULO DO PROJETO"]
            default_list = ["PROJETO CONCEITUAL/BÁSICO/EXECUTIVO",
                            "EXEMPLO (NOME DO PORTO)",
                            "EXEMPLO (PROJETO DE DRAGAGEM)"]
            output = multenterbox(text, title, input_list, default_list)
            ld_information["project_title"] = output[1]
            ld_information["ld_title"] = "\n".join(output)\
                                         + "\nLISTA DE DOCUMENTOS"

        text = "Defina as iniciais dos responsáveis (formato XXX)"
        title = "Defina as iniciais"
        input_list = ["EXECUÇÃO", "VERIFICAÇÃO", "APROVAÇÃO"]
        if self.ld_rev == -1:
            default_list = ["XXX", "XXX", "XXX"]
        else:
            revision = self.ld_rev + 1
            previous_cover_cell = get_cover_cell(revision - 1)
            # book_path = os.path.join(self.emited_path, '_LDs', self.ld_name + '.xlsx')
            book_path = os.path.join(self.ld_path, self.ld_name)
            default_list = get_acronym_default_list(book_path,
                                                    previous_cover_cell)
        output = multenterbox(text, title, input_list, default_list)
        ld_information["acronym1"] = output[0]
        ld_information["acronym2"] = output[1]
        ld_information["acronym3"] = output[2]
        if self.ld_rev == -1:
            ld_information["client_img"] = self.get_client_img()
        else:
            ld_information["client_img"] = None

        return ld_information

    def check_open_files(self):
        '''
        Checks if a file is open
        '''
        file_open = True
        while file_open:
            try:
                for doc in self.docs:
                    if doc['emit']:
                        src = Path(os.path.join(str(self.base_path), doc['subdir'], doc['file_name']))
                        os.replace(src, src)
                file_open = False
            except OSError:
                file_open = True
                text = "O arquivo " + doc['file_name'] + " está aberto. Feche-o e clique em repetir para continuar a operação."
                title = "Todos os arquivos devem estar fechados"
                button_list = ["Repetir", "Cancelar"]
                output = buttonbox(text, title, button_list)
                if output == "Repetir":
                    pass
                elif output == "Cancelar":
                    sys.exit(0)

    def issued_directories(self):
        # Deletes the revision suffix from the filename
        # filenames = []
        dirs_to_create = {}
        for doc in self.docs:
            if doc['emit']:
                folder_name = self.get_folder_name(doc['file_name'], self.file_num_caract)
                if folder_name not in self.directories:
                    dir_to_create = os.path.join(self.emited_path,
                                                 doc['subdir'], folder_name)
                    dirs_to_create[folder_name] = dir_to_create
                    self.directories[folder_name] = doc['subdir']
                else:
                    self.check_file(doc, folder_name)

        return dirs_to_create

    @staticmethod
    def create_dirs(dirs_to_create):
        for dir in dirs_to_create.values():
            Path(dir).mkdir(parents=True, exist_ok=True)

    @staticmethod
    def detect_num_caract(filename):
        if len(filename) > 14 and filename[14] == '-':
            return 23
        if len(filename) > 16 and filename[16] == '-':
            return 24
        return None

    @staticmethod
    def get_folder_name(filename, num_caract):
        detected = Emission.detect_num_caract(filename)
        return filename[:detected] if detected is not None else filename[:num_caract]


    def move_files(self):
        for directory in self.directories.keys():
            for doc in self.docs:
                if doc['emit'] and doc['file_name'].startswith(directory):
                    src = Path(os.path.join(str(self.base_path), doc['subdir'], doc['file_name']))
                    dest = Path(os.path.join(os.path.join(self.emited_path,
                                                          self.directories[directory],
                                                          directory),
                                             doc['file_name']))
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    os.replace(src, dest)
        msg = "A emissão foi realizada com sucesso."
        title = "Documentos emitidos"
        msgbox(msg, title)

    def get_file_name(self, doc):
        detected = self.detect_num_caract(doc)
        return doc[:detected] if detected is not None else doc[:self.file_num_caract]

    def get_probably_name(self):
        doc_name = self.docs[0]
        probably_name = doc_name['file_name'][9:12]

        return probably_name

    def get_revision(self, doc):
        filename = os.path.splitext(doc)[0]
        pattern = self.rev_reg_expression
        if re.search(pattern, filename) is not None:
            rev = re.search(pattern, filename).group()
            rev = int(''.join(filter(str.isdigit, rev)))
        else:
            rev = 0
        return rev

    def verify_pattern(self, doc_name):
        doc_name_no_extension = os.path.splitext(doc_name)[0]
        pattern = self.doc_reg_expression
        if re.match(pattern, doc_name_no_extension):
            return True
        else:
            return False

    @staticmethod
    def verify_ld_pattern_no_rev(doc_name):
        pattern = r'^IFS-\d{4}-\d{3}-(\w{3}-\w{2}-\d{4,}|\w-\w{2}-\d{5,})(_R\d+)?$'
        if re.match(pattern, doc_name):
            return True
        else:
            return False
        
    @staticmethod
    def is_hidden(filepath):
        name = os.path.basename(filepath)
        if name.startswith('.'):
            return True
        try:
            # Windows: FILE_ATTRIBUTE_HIDDEN = 0x02
            return bool(os.stat(filepath).st_file_attributes & stat.FILE_ATTRIBUTE_HIDDEN)
        except Exception:
            return False

    @staticmethod
    def get_ld_revision(doc_name):
        doc_name_no_extension = os.path.splitext(doc_name)[0]
        pattern = r'^IFS-\d{4}-\d{3}-(\w{3}-\w{2}-\d{4}|\w-\w{2}-\d{5})(_R\d+)?$'
        if not re.match(pattern, doc_name_no_extension):
            return -1
        else:
            revision = re.search(r'(?i)_R\d+$',
                                 os.path.splitext(doc_name_no_extension)[0]
                                 ).group()
            revision = int(''.join(filter(str.isdigit, revision)))
            return revision

    @staticmethod
    def verify_date_pattern(date):
        pattern = r"\d{2}/\d{2}/\d{2}"
        if re.match(pattern, date):
            return True
        else:
            return False

#   TODO : Verificar arquivos que terminam com Rev

    @staticmethod
    def text_box(msg, title):
        if ccbox(msg, title):
            pass
        else:
            sys.exit(0)


if __name__ == '__main__':
    # os.chdir(r'C:\Users\bruno\OneDrive\Documentos\LD\2227 Exemplo\5_Engenharia\_PARA EMISSAO')
    emis = Emission()
    emis.check_filename_pattern()
    dirs_to_create = emis.issued_directories()
    emis.confirm_files(dirs_to_create)
    emis.create_dirs(dirs_to_create)
    emis.ld_information = emis.get_ld_information()
    emis.check_open_files()
    emis.create_zip()
    emis.create_ld()
    emis.move_files()
