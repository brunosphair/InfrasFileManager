import os
import re
import sys
from pathlib import Path
from easygui import *
from zipfile import ZipFile
import openpyxl

class Emission:
    def __init__(self):
        self.issued_path = self.get_issued_path()
        self.docs = self.get_files()
        self.directories = self.get_directories()
        self.grd_number = self.get_grd_number()
        self.project_number = self.get_project_number()
        self.grd_name = 'IFS-GRD-' + self.project_number + '-' + str(self.grd_number).zfill(3)
        self.file_num_caract = 23

    
    def get_files(self):
        docs = []
        for file in os.listdir('.'):
            if os.path.isfile(file):
                rev = self.get_revision(file)
                docs.append([file, rev, True])
        return docs

    def get_issued_path(self):
        path = Path(os.getcwd()).parent.absolute()
        parent_path = path.parent.absolute()
        issued_path = os.path.join(parent_path, '3_Emitidos')
        if not os.path.isdir(issued_path):
            raise FileNotFoundError("A pasta 3_Emitidos não foi encontrada")
        return issued_path

    def get_directories(self):
        directories = []
        for file in os.listdir(self.issued_path):
            if os.path.isdir(os.path.join(self.issued_path, file)):
                directories.append(file)
        return directories

    def get_grd_number(self):
        grd_directory = os.path.join(self.issued_path, '_GRDs')
        if not os.path.isdir(grd_directory):
            raise FileNotFoundError("A pasta _GRDs não foi encontrada")
        grds = os.listdir(grd_directory)
        print(grds)
        grd_number = 1
        for item in grds:
            if os.path.isdir(os.path.join(grd_directory, item)):
                if item.startswith('GRD-') and int(item[-3:])>= grd_number:
                    grd_number = int(item[-3:]) + 1
        return grd_number

    def get_project_number(self):
        path = Path(os.getcwd()).parent.absolute()
        project_path = path.parent.absolute()
        dir_name = os.path.basename(project_path)
        project_number = dir_name[:4]
        if not project_number.isnumeric():
            raise ValueError("O arquivo executado não está na pasta correta")
        return project_number

    def check_pattern(self):
        ignored_files = []
        for doc in self.docs:
            if not self.verify_pattern(doc[0]):
                ignored_files.append(doc[0])
                doc[2] = False
        if len(ignored_files):
            msg = "Os seguintes arquivos não serão emitidos, pois não correspondem ao padrão de nomenclatura de arquivos:\n\n" + '\n'.join(ignored_files) + "\n\nO que deseja fazer?"
            title = "Inconsistência na nomenclatura dos arquivos"
            if ccbox(msg,title):
                pass
            else:
                sys.exit(0)

    def check_files(self):
        for doc in self.docs:
            if doc[2]:
                path_name = doc[0][:self.file_num_caract]
                doc_name = self.get_file_name(doc[0])
                doc_directory = os.path.join(self.issued_path, path_name)
                if os.path.isdir(doc_directory):
                    for file in os.listdir(doc_directory):
                        try:
                            file_name = self.get_file_name(file)
                            if self.get_revision(file) == doc[1] and file_name == doc_name:
                                raise NameError
                        except NameError:
                            msg = 'O arquivo ' + doc_name + ' com a revisão ' + str(doc[1]) + ' já existe. O que deseja fazer?'
                            choices = ["Ignorar arquivo", "Substituir"]
                            title = "Arquivo duplicado"
                            choice = buttonbox(msg, title, choices)
                            if choice == "Ignorar arquivo":
                                print("arquivo ignorado")
                                doc[2] = False
                            elif choice == "Substituir":
                                print("arquivo substituído")
                                doc [2] = True
        msg = 'Os seguintes arquivos serão emitidos na GRD 0' + str(self.grd_number) + '. Desmarque caso não deseje emitir algum:'
        list_of_options = []
        for doc in self.docs:
            if doc[2]:
                list_of_options.append(doc[0])
        title = 'Deseja continuar?'
        choices = multchoicebox(msg, title, list_of_options, preselect=[*range(len(list_of_options))])
        print(choices)
        for doc in self.docs:
            if not doc[0] in choices:
                doc[2] = False

    def create_zip(self):
        zipObj = ZipFile(self.grd_name + '.zip', 'w')
        for doc in self.docs:
            if doc[2]:
                zipObj.write(doc[0])
        zipObj.close()

    def create_grd(self):
        grd_directory = os.path.join(self.issued_path, '_GRDs')
        os.mkdir(os.path.join(grd_directory, self.grd_name))
        no_docs = []
        grd_items = []
        for doc in self.docs:
            doc_name = self.get_file_name(doc[0])
            if doc[2] and not doc_name in no_docs:
                no_docs.append(doc_name)
                grd_items.append([doc_name, doc[1]])
        self.create_excel_grd(grd_items)


    def create_excel_grd(self, grd_items):
        book_path = os.path.join(self.issued_path, '_GRDs', 'IFS-GRD-XXXX-XXX.xlsx')
        book = openpyxl.load_workbook(book_path)
        sheet = book.active
        i = 1
        for item in grd_items:
            sheet.cell(row = 25 + i, column = 1).value = str(i)
            sheet.cell(row = 25 + i, column = 2).value = item[0]
            sheet.cell(row = 25 + i, column = 16).value = str(item[1])
            i += 1
        sheet.cell(row = 10, column = 12).value = self.grd_name
        grd_final_path = os.path.join(self.issued_path, '_GRDs', self.grd_name,  self.grd_name + '.xlsx')
        book.save(filename = grd_final_path)
        book.close()

    def move_files(self):
        # Deletes the revision suffix from the filename
        filenames = []
        for doc in self.docs:
            if doc[2]:
                filenames.append(doc[0][:self.file_num_caract])
        set(filenames)
        # If the file directory doesn't exists, the code creates it
        for item in filenames:
            if not item in self.directories:
                dir_to_create = os.path.join(self.issued_path, item)
                os.mkdir(dir_to_create)
                self.directories.append(item)
        for directory in self.directories:
            for doc in self.docs:
                if doc[2] and doc[0].startswith(directory):
                    src = Path(doc[0])
                    dest = Path(os.path.join(os.path.join(self.issued_path, directory), doc[0]))
                    os.replace(src, dest)


    @staticmethod
    def get_file_name(doc):
        doc = os.path.splitext(doc)[0]
        if re.search(r'(?i)_R\d+$', doc):
            filename = doc.rsplit("_", 1)[0]
        else:
            filename = doc
        return filename

    @staticmethod
    def get_revision(doc):
        filename = os.path.splitext(doc)[0]
        if re.search(r'(?i)_rev\d+$',os.path.splitext(filename)[0]) != None:
            rev = re.search(r'(?i)_rev\d+$',os.path.splitext(filename)[0]).group()
            rev = int(''.join(filter(str.isdigit, rev)))
        elif re.search(r'(?i)_R\d+$',os.path.splitext(filename)[0]) != None:
            rev = re.search(r'(?i)_R\d+$',os.path.splitext(filename)[0]).group()
            rev = int(''.join(filter(str.isdigit, rev)))
        else:
            rev = 0
        return rev
    
    @staticmethod
    def verify_pattern(doc_name):
        doc_name_no_extension = os.path.splitext(doc_name)[0]
        pattern = r'^IFS-\d{4}-\d{3}-\w{1}-\w{2}-\d{5}.*(_R\d{1,2})?$'
        if re.match(pattern, doc_name_no_extension):
            return True
        else:
            return False
        
if __name__ == '__main__':
    os.chdir(r'C:\Users\Bruno\OneDrive\Documentos\LD\2227 Exemplo\5_Engenharia\_PARA EMISSAO')
    emis = Emission()
    emis.check_pattern()
    emis.check_files()
    emis.create_zip()
    emis.create_grd()
    emis.move_files()
    # print(emis.get_project_number())
    # print(emis.verify_pattern("IFS-2122-110-B-CP-00001_R0"))

