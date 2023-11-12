import os
import re
import sys
from pathlib import Path
from easygui import buttonbox, ccbox, multchoicebox, enterbox, msgbox, \
                    multenterbox
from zipfile import ZipFile
import datetime
from dotenv import dotenv_values

from excel_functions import get_grd_number, create_excel_grd, \
                            get_acronym_default_list, get_cover_cell


class Emission:
    def __init__(self):
        self.doc_reg_expression, self.rev_reg_expression = \
                                                    self.get_reg_expressions()
        self.emited_path = self.get_emited_path()
        self.docs = self.get_files()
        self.directories = self.get_emited_directories()
        self.ld_rev = self.get_ld_rev()
        self.project_number = self.get_project_number()
        self.grd_number = get_grd_number(self.emited_path, self.ld_name)
        self.grd_name = 'IFS-GRD-' + \
                        str(self.project_number) + \
                        "-" + str(self.grd_number).zfill(3)
        self.file_num_caract = 23
        self.ld_information = {}

    def get_files(self):
        '''
        Returns a list of dictionarys, where which dict has the filename, the
        revision number and declares the 'emit' key as True.
        '''
        docs = []
        file_names = []
        for path, subdir, files in os.walk('.'):
            for file in files:
                if file not in file_names:
                    file_names.append(file)
                    dict = {}
                    rev = self.get_revision(file)
                    dict['file_name'] = file
                    dict['rev'] = rev
                    dict['emit'] = True
                    dict['subdir'] = os.path.relpath(path)
                    docs.append(dict)
                else:
                    msg = "Há dois arquivos com o nome " + file + " dentro da emissão"
                    title = "ERRO"
                    msgbox(msg, title)
                    sys.exit(0)
        return docs

    def get_emited_path(self):
        '''
        Returns the path to the folder 3_Emitidos
        '''
        path = Path(os.getcwd()).parent.absolute()
        parent_path = path.parent.absolute()
        issued_path = os.path.join(parent_path, '3_Emitidos')
        if not os.path.isdir(issued_path):
            raise FileNotFoundError("A pasta 3_Emitidos não foi encontrada")
        return issued_path

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
        lds_directory = os.path.join(self.emited_path, '_LDs')
        if not os.path.isdir(lds_directory):
            raise FileNotFoundError("A pasta _LDs não foi encontrada")
        lds = os.listdir(lds_directory)
        self.ld_name = 'IFS-XXXX-XXX-X-LD-XXXX.xlsx'
        last_revision = -1
        for item in lds:
            ld_revision = self.get_ld_revision(item)
            if last_revision < ld_revision:
                last_revision = ld_revision
                self.ld_name = item

        return last_revision

    def get_reg_expressions(self):
        try:
            config = dotenv_values('config.env')
            print(config)
            doc_reg_expression = config['DOC_REGULAR_EXPRESSION']
            rev_reg_expression = config['REV_REGULAR_EXPRESSION']
        except KeyError:
            doc_reg_expression = \
                        r'^IFS-\d{4}-\d{3}-\w{1}-\w{2}-\d{5}.*(_R\d{1,2})?$'
            rev_reg_expression = r'(?i)_R\d+$'

        return doc_reg_expression, rev_reg_expression

    def get_project_number(self):
        '''
        Retuns the number of the project.
        '''
        path = Path(os.getcwd()).parent.absolute()
        project_path = path.parent.absolute()
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
                if not doc['file_name'].startswith('InfrasEmission'):
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
                    folder_name = self.get_folder_name(doc['file_name'],
                                                       self.file_num_caract)
                    if folder_name in dirs_to_create:
                        del dirs_to_create[folder_name]

    def create_zip(self):
        zipObj = ZipFile(self.grd_name + '.zip', 'w')
        for doc in self.docs:
            if doc['emit']:
                zipObj.write(os.path.join(doc['subdir'], doc['file_name']))
        zipObj.close()

    def create_ld(self):
        no_docs = []
        grd_items = []
        for doc in self.docs:
            doc_name = self.get_file_name(doc['file_name'])
            if doc['emit'] and doc_name not in no_docs:
                no_docs.append(doc_name)
                grd_items.append([doc_name, doc['rev']])
        create_excel_grd(self.emited_path, self.ld_name, self.grd_number,
                         self.grd_name, self.ld_information, self.ld_rev,
                         self.file_num_caract, grd_items)

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
                "nome para a LD no padrão IFS-NNNN-NNN-X-LD-NNNNN onde X são "\
                "letras e N são números"
            title = "Nomeie a LD"
            probably_name = self.get_probably_name()
            d_text = "IFS-"\
                     + str(self.project_number)\
                     + "-" + probably_name + "-G-LD-00001"
            defined_name = False
            while not defined_name:
                ld_name = enterbox(text, title, d_text)
                if self.verify_ld_pattern_no_rev(ld_name):
                    defined_name = True
                else:
                    msgbox("O nome que você digitou não atende aos requisitos"
                           " de IFS-NNNN-NNN-X-LD-NNNNN, digite novamente",
                           "Nome inválido!")
            ld_information["ld_name"] = ld_name

            text = "Defina os títulos da LD:"
            title = "Definir títulos"
            input_list = ["1ª LINHA - TIPO DE PROJETO ",
                          "2ª LINHA - TÍTULO DO PROJETO",
                          "3ª LINHA - TÍTULO DO DOCUMENTO"]
            default_list = ["PROJETO CONCEITUAL/BÁSICO/EXECUTIVO",
                            "TÍTULO DO PROJETO",
                            "TÍTULO DO DOCUMENTO"]
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
            book_path = os.path.join(self.emited_path, '_LDs', self.ld_name)
            default_list = get_acronym_default_list(book_path,
                                                    previous_cover_cell)
        output = multenterbox(text, title, input_list, default_list)
        ld_information["acronym1"] = output[0]
        ld_information["acronym2"] = output[1]
        ld_information["acronym3"] = output[2]

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
                        src = Path(os.path.join(doc['subdir'], doc['file_name']))
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
                folder_name = self.get_folder_name(doc['file_name'],
                                                   self.file_num_caract)
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
    def get_folder_name(filename, num_caract):
        folder_name = filename[:num_caract]
        return folder_name

    def move_files(self):
        for directory in self.directories.keys():
            for doc in self.docs:
                if doc['emit'] and doc['file_name'].startswith(directory):
                    src = Path(os.path.join(doc['subdir'], doc['file_name']))
                    dest = Path(os.path.join(os.path.join(self.emited_path,
                                                          self.directories[directory],
                                                          directory),
                                             doc['file_name']))
                    os.replace(src, dest)
        msg = "A emissão foi realizada com sucesso."
        title = "Documentos emitidos"
        msgbox(msg, title)

    def get_file_name(self, doc):
        filename = doc[:self.file_num_caract]
        return filename

    def get_probably_name(self):
        doc_name = self.docs[0]
        probably_name = doc_name['file_name'][9:12]

        return probably_name

    def get_revision(self, doc):
        filename = os.path.splitext(doc)[0]
        pattern = self.rev_reg_expression
        if re.search(pattern, os.path.splitext(filename)[0]) is not None:
            rev = re.search(pattern, os.path.splitext(filename)[0]).group()
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
        pattern = r'^IFS-\d{4}-\d{3}-\w{1}-LD-\d{5}$'
        if re.match(pattern, doc_name):
            return True
        else:
            return False

    @staticmethod
    def get_ld_revision(doc_name):
        doc_name_no_extension = os.path.splitext(doc_name)[0]
        pattern = r'^IFS-\d{4}-\d{3}-\w{1}-LD-\d{5}.*(_R\d{1,2})?$'
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
    # os.chdir(r'C:\Users\bruno\OneDrive\Documentos\LD\2301 EMAP_Executivo\5_Engenharia\_PARA EMISSAO')
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
