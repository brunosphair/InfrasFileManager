import os
import re
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
from zipfile import ZipFile
import openpyxl


class DesfazerEmissao:
    def __init__(self, base_path):
        self.base = Path(base_path)
        self.emited_path = self.get_emited_path()
        self.ld_path = self._get_ld_path()
        self.last_zip = self._find_last_zip()

        self.padrao_re = re.compile(r'^IFS-\d{4}-\d{3}-(\w{3}-\w{2}-\d{4}|\w-\w{2}-\d{5})(_R\d+)?$')


    def _get_ld_path(self) -> Path:
        ld = self.emited_path / '_LDs'
        if ld.is_dir():
            return ld
        ld = self.base / '00_LDs'
        if ld.is_dir():
            return ld
        return None

    def get_emited_path(self) -> Path:
        emited_path = self.base.parent.parent / '3_Emitidos'
        if not emited_path.is_dir():
            raise ValueError("Pasta de emitidos nao encontrada!")
        return emited_path

    @staticmethod
    def extrair_revisao(nome_base):
        match = re.search(r'_R(\d+)$', nome_base)
        return int(match.group(1)) if match else 0
    def confirm_zip(self):
        candidatos = []

        for nome in os.listdir(self.ld_path):
            raiz, ext = os.path.splitext(nome)
            if ext.lower() == ".xlsx" and self.padrao_re.match(raiz):
                candidatos.append((nome, self.extrair_revisao(raiz)))
        if not candidatos:
            messagebox.showwarning('Resultado', 'Nenhuma planilha encontrada com o padrão IFS.')
            return
        maior_rev = max(r for _, r in candidatos)
        planilha_maior_rev = [nome for nome, r in candidatos if r == maior_rev][0]

        wb = openpyxl.load_workbook(os.path.join(self.ld_path, planilha_maior_rev), read_only=True)

        num_ultima_grd = wb.sheetnames[-1].split("-")[1]
        num_ultimo_zip = os.path.splitext(str(self.last_zip).split("/")[-1])[0].split("-")[-1]
        wb.close()
        if num_ultima_grd == num_ultimo_zip:
            return True
        return False

    def _find_last_zip(self):
        pattern = re.compile(r'IFS-GRD-\d{4}-(\d{3})\.zip$', re.IGNORECASE)
        candidates = []
        for f in self.base.iterdir():
            m = pattern.search(f.name)
            if m:
                candidates.append((int(m.group(1)), f))
        print(f'candidatos antes do sort: {candidates}')
        if not candidates:
            return None
        candidates.sort(key=lambda x: x[0])
        print(f'Cantidatos apos o sort: {candidates}')
        print(candidates[-1][1])
        return candidates[-1][1]

    def pode_desfazer(self):
        if not self.ld_path or not self.ld_path.is_dir():
            return False
        matches = sum(
            1 for nome in os.listdir(self.ld_path)
            if (self.ld_path / nome).is_file()
            and os.path.splitext(nome)[1].lower() == '.xlsx'
            and self.padrao_re.match(os.path.splitext(nome)[0])
        )
        return matches >= 1

    def nome_ultima_emissao(self):
        return self.last_zip.stem if self.last_zip else None

    def listar_arquivos(self):
        if not self.last_zip:
            return []
        with ZipFile(str(self.last_zip), 'r') as z:
            entries = z.namelist()
        return [
            entry.replace('\\', '/').split('/')[-1]
            for entry in entries
            if entry.replace('\\', '/').split('/')[-1]
        ]

    def janela_confirmacao(self, parent):
        arquivos = self.listar_arquivos()
        confirmado = tk.BooleanVar(value=False)

        janela = tk.Toplevel(parent)
        janela.title("Confirmar desfazer emissão")
        janela.resizable(False, False)
        janela.grab_set()

        tk.Label(janela, text=f"Emissão: {self.nome_ultima_emissao()}", font=("Arial", 10, "bold")).pack(pady=(15, 5), padx=20)
        tk.Label(janela, text="Arquivos que serão restaurados para Para_Emissao:", font=("Arial", 10)).pack(anchor="w", padx=20)

        lista_frame = tk.Frame(janela)
        lista_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(5, 0))

        scrollbar = tk.Scrollbar(lista_frame, orient=tk.VERTICAL)
        listbox = tk.Listbox(lista_frame, yscrollcommand=scrollbar.set, font=("Arial", 9), height=12, width=60)
        scrollbar.config(command=listbox.yview)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        for arquivo in arquivos:
            listbox.insert(tk.END, arquivo)

        tk.Label(janela, text="O arquivo LD criado também será deletado.\nEsta ação não pode ser desfeita.",
                 font=("Arial", 9), fg="gray").pack(pady=(10, 5))

        btn_frame = tk.Frame(janela)
        btn_frame.pack(pady=(0, 15))

        tk.Button(btn_frame, text="Confirmar", font=("Arial", 10), width=15, fg="darkred",
                  command=lambda: [confirmado.set(True), janela.destroy()]).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Cancelar", font=("Arial", 10), width=15,
                  command=janela.destroy).pack(side=tk.LEFT, padx=10)

        parent.wait_window(janela)
        return confirmado.get()

    def executar(self):
        '''
        Desfaz a última emissão:
          - Move os arquivos de 3_Emitidos de volta para Para_Emissao
          - Deleta o ZIP da emissão
          - Deleta o LD _RN.xlsx mais recente
          - Remove pastas vazias criadas em 3_Emitidos
        Retorna (True, []) em caso de sucesso ou (False, lista_de_erros).
        '''
        if not self.last_zip:
            return False, ["Nenhuma emissão encontrada para desfazer."]


        errors = []
        moved_dirs = set()

        with ZipFile(str(self.last_zip), 'r') as z:
            entries = z.namelist()

        for entry in entries:
            parts = [p for p in entry.replace('\\', '/').strip('/').split('/') if p and p != '.']
            if not parts:
                continue

            filename = parts[-1]
            subdir_parts = parts[:-1]
            folder_name = self._get_folder_name(filename)

            if subdir_parts:
                current = self.emited_path.joinpath(*subdir_parts) / folder_name / filename
                dest = self.base.joinpath(*subdir_parts) / filename
            else:
                current = self._find_in_emited(folder_name, filename)
                dest = self.base / filename

            moved_dirs.add(current.parent)

            try:
                dest.parent.mkdir(parents=True, exist_ok=True)
                os.replace(str(current), str(dest))
            except Exception as e:
                errors.append(f"Erro ao restaurar '{filename}': {e}")

        self._deletar_zip(errors)
        self._deletar_ld(errors)
        self._remover_pastas_vazias(moved_dirs)

        return len(errors) == 0, errors

    def _deletar_zip(self, errors):
        try:
            os.remove(str(self.last_zip))
        except Exception as e:
            errors.append(f"Erro ao remover ZIP: {e}")

    def _deletar_ld(self, errors):
        if not self.ld_path or not self.ld_path.is_dir():
            return
        ld_files = []
        for f in self.ld_path.iterdir():
            m = re.search(r'_R(\d+)', f.stem, re.IGNORECASE)
            if f.suffix == '.xlsx' and m:
                ld_files.append((int(m.group(1)), f))
        if not ld_files:
            return
        ld_files.sort(key=lambda x: x[0])
        try:
            os.remove(str(ld_files[-1][1]))
        except Exception as e:
            errors.append(f"Erro ao remover LD: {e}")

    def _remover_pastas_vazias(self, dirs):
        for dir_path in sorted(dirs, key=lambda p: len(str(p)), reverse=True):
            try:
                if dir_path.is_dir() and not any(dir_path.iterdir()):
                    dir_path.rmdir()
                    parent = dir_path.parent
                    if parent.is_dir() and not any(parent.iterdir()) and parent != self.emited_path:
                        parent.rmdir()
            except Exception:
                pass

    def _find_in_emited(self, folder_name, filename):
        direct = self.emited_path / folder_name / filename
        if direct.exists():
            return direct
        for subdir in self.emited_path.iterdir():
            if subdir.is_dir():
                candidate = subdir / folder_name / filename
                if candidate.exists():
                    return candidate
        return direct

    @staticmethod
    def _get_folder_name(filename):
        if len(filename) > 14 and filename[14] == '-':
            return filename[:23]
        if len(filename) > 16 and filename[16] == '-':
            return filename[:24]
        return filename[:23]
