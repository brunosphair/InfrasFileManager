import os
from easygui import *
import pyperclip
from zipfile import ZipFile


def rename_files():
    text = 'Entre com os dados abaixo (para apenas apagar o texto desejado, deixar o "Substituir por:" vazio)'
    title = 'Substituir nome de arquivo'
    input_list = ['Localizar', 'Substituir por:']
    default_list = ['Digite Aqui...', '']

    output = multenterbox(text, title, input_list, default_list)

    print(output)
    text = output[0]
    new_text = output[1]

    for filename in os.listdir('.'):
        if text in filename:
            original_path = filename
            new_path = original_path.replace(text, new_text)
            os.rename(original_path, new_path)


def copy_file_names():
    msg = "Deseja copiar as extensões dos arquivos?"
    choices = ["Sim", "Não"]
    title = "Infras File Manager v1.0"
    choice = buttonbox(msg, title, choices)
    clipboard = ''
    if choice == "Sim":
        for filename in os.listdir('.'):
            clipboard += filename + '\n'
        pyperclip.copy(clipboard)
    else:
        for filename in os.listdir('.'):
            filename = os.path.splitext(filename)[0]
            clipboard += filename + '\n'
        pyperclip.copy(clipboard)


def create_zips():
    filenames = sorted(os.listdir('.'))
    n = 0
    while n < len(filenames) - 1:
        filename_no_ext = os.path.splitext(filenames[n])[0]
        if os.path.splitext(filenames[n+1])[0] == filename_no_ext:
            zip_obj = ZipFile(filename_no_ext + '.zip', 'w')
            zip_obj.write(filenames[n])
            n += 1
            while os.path.splitext(filenames[n])[0] == filename_no_ext:
                zip_obj.write(filenames[n])
                n += 1
            zip_obj.close()
        else:
            n += 1

if __name__ == '__main__':
#     msg = "Selecione que função deseja utilizar:"
#     title = "Infras File Manager v1.0"
#     choices = ["Renomear arquivos da pasta", "Copiar nome dos arquivos da pasta",
#                "Criar zips com arquivos da pasta que têm o mesmo nome"]
#     choice = choicebox(msg, title, choices)
#     if choice == "Renomear arquivos da pasta":
#         rename_files()
#     elif choice == "Copiar nome dos arquivos da pasta":
#         copy_file_names()
#     elif choice == "Criar zips com arquivos da pasta que têm o mesmo nome":
#         create_zips()
    doc_list(r'C:\Users\Bruno\PycharmProjects\infras_file_manager\InfrasFileManager\example\3_Emitidos\_GRDs')