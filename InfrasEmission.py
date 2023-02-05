import os
import re
import sys
from pathlib import Path
from easygui import *
from zipfile import ZipFile

class Emission:
    def __init__(self):
        self.issued_path = self.get_issued_path()
        self.docs = self.get_files()
        self.directories = self.get_directories()
        self.grd_number = self.get_grd_number()

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
        return issued_path

    def get_directories(self):
        directories = []
        for file in os.listdir(self.issued_path):
            if os.path.isdir(os.path.join(self.issued_path, file)):
                directories.append(file)
        return directories

    def get_grd_number(self):
        grd_directory = os.path.join(self.issued_path, '_GRDs')
        grds = os.listdir(grd_directory)
        print(grds)
        grd_number = 1
        for item in grds:
            if os.path.isdir(os.path.join(grd_directory, item)):
                if item.startswith('GRD-') and int(item[-4:])>= grd_number:
                    grd_number = int(item[-4:]) + 1
        return grd_number

    def check_files(self):
        for doc in self.docs:
            print(doc)
            doc_name = self.get_file_name(doc[0])
            doc_directory = os.path.join(self.issued_path, doc_name)
            if os.path.isdir(doc_directory):
                for file in os.listdir(doc_directory):
                    try:
                        if self.get_revision(file) == doc[1]:
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
        zipObj = ZipFile('GRD-' + str(self.grd_number).zfill(4) + '.zip', 'w')
        for doc in self.docs:
            if doc[2]:
                zipObj.write(doc[0])
        zipObj.close()

    def create_grd_directory(self):
        grd_directory = os.path.join(self.issued_path, '_GRDs')
        grd_name = 'GRD-' + str(self.grd_number).zfill(4)
        os.mkdir(os.path.join(grd_directory, grd_name))


        # if confirm:

        # else:
        #     sys.exit()

    def create_directories(self):
        # Deletes the revision suffix from the filename
        filenames = []
        for doc in self.docs:
            if doc[2]:
                filenames.append(self.get_file_name(doc[0]))
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
                    # dest.write_bytes(src.read_bytes())


    @staticmethod
    def get_file_name(doc):
        doc = os.path.splitext(doc)[0]
        if re.search(r'(?i)_rev\d+$', doc) or re.search(r'(?i)_R\d+$', doc):
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

if __name__ == '__main__':
    os.chdir(r'C:\Users\Bruno\OneDrive\Documentos\LD\5_Engenharia\_PARA EMISSAO')
    emis = Emission()
    emis.check_files()
    emis.create_zip()
    emis.create_grd_directory()
    emis.create_directories()
