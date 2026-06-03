import tkinter as tk
from tkinter import messagebox
import os
import re
import sys
from InfrasEmission import Emission
from DesfazerEmissao import DesfazerEmissao
from Exportacao import Exportar

class GerenciadorProjeto():
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Seletor de Pastas")
        self.root.geometry("500x400")
        self.root.resizable(False, False)

        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        self.base_dir = os.path.abspath(os.path.join(base_dir))
        self.pastas = self._listar_pastas()

        self.frame_atual = None

        self._mostrar_frame(self.fazer_login)
        self.root.mainloop()

    def _mostrar_frame(self, builder, *args):
        '''
            Limpa o frame atual e constrói um novo usando a função builder fornecida.
        '''
        if self.frame_atual:
            self.frame_atual.destroy()
        self.frame_atual = tk.Frame(self.root)
        self.frame_atual.pack(fill=tk.BOTH, expand=True, padx=15)
        # Chama a classe enviada como parametro no _mostrar_frame
        builder(self.frame_atual, *args)

    def fazer_login(self, frame):

        tk.Label(frame, text="Insira a senha:", font=("Arial", 10)).pack(pady=(80, 5))
        senha_entry = tk.Entry(frame, font=("Arial", 10), show="*")
        senha_entry.pack(pady=(0, 10))
        senha_entry.focus()

        def validar_senha():
            senha = senha_entry.get()
            if senha == "IFS":
                self._mostrar_frame(self._build_ui_find_project)
            else:
                messagebox.showerror("Erro", "Senha inválida")
                senha_entry.delete(0, tk.END)

        senha_entry.bind("<Return>", lambda _: validar_senha())
        tk.Button(frame, text="Entrar", font=("Arial", 10), width=15, command=validar_senha).pack()
        assets_dir = getattr(sys, '_MEIPASS', self.base_dir)
        img_path = os.path.join(assets_dir, "assets", "Logo_infras.png")
        if os.path.exists(img_path):
            photo = tk.PhotoImage(file=img_path.replace("\\", "/")).subsample(12, 12)
            label_img = tk.Label(frame, image=photo, bg="white")
            label_img.image = photo
            label_img.pack(pady=(20, 5))

    def _listar_pastas(self):
        '''
            Lista as pastas do projeto
        '''
        padrao = re.compile(r'^\d{4}')
        try:
            return [
                p for p in os.listdir(self.base_dir)
                if os.path.isdir(os.path.join(self.base_dir, p)) and padrao.match(p)
            ]
        except (PermissionError, FileNotFoundError):
            return []

    def _build_ui_find_project(self, frame):
        '''
            Constrói a interface para seleção de pastas
        '''
        tk.Label(frame, text="Pastas disponíveis:", font=("Arial", 11)).pack(pady=(15, 5))

        search_var = tk.StringVar()
        search_entry = tk.Entry(frame, textvariable=search_var, font=("Arial", 10))
        search_entry.pack(fill=tk.X, padx=20, pady=(0, 5))

        lista_frame = tk.Frame(frame)
        lista_frame.pack(fill=tk.BOTH, expand=True, padx=20)

        scrollbar = tk.Scrollbar(lista_frame, orient=tk.VERTICAL)
        self.listbox = tk.Listbox(lista_frame, yscrollcommand=scrollbar.set, font=("Arial", 10), height=15)
        scrollbar.config(command=self.listbox.yview)

        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.listbox.bind("<Double-Button-1>", lambda _: self._on_selecionar())
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        for pasta in sorted(self.pastas):
            self.listbox.insert(tk.END, pasta)

        def on_search(*args):
            termo = search_var.get().lower()
            self.listbox.delete(0, tk.END)
            for pasta in sorted(self.pastas):
                if termo in pasta.lower():
                    self.listbox.insert(tk.END, pasta)

        search_var.trace_add("write", on_search)

        tk.Button(
            frame,
            text="Selecionar",
            font=("Arial", 10),
            width=20,
            command=self._on_selecionar,
        ).pack(pady=15)

    def _get_caminhos(self, nome):
        caminho = os.path.join(self.base_dir, nome, "2_Producao", "06_Para_Emissao")
        caminho_ld = os.path.join(caminho, "00_LDs")
        if not os.path.isdir(caminho):
            caminho = os.path.join(self.base_dir, nome, "5_Engenharia", "_PARA EMISSAO")
            caminho_ld = os.path.join(self.base_dir, nome, "3_Emitidos", "_LDs")
        if not os.path.isdir(caminho):
            raise FileNotFoundError(f"Pasta de emissão não encontrada para o projeto '{nome}'")
        if not os.path.isdir(caminho_ld):
            raise FileNotFoundError(f"Pasta de LDs não encontrada para o projeto '{nome}'")
        return caminho, caminho_ld

    def _get_arquivos_ld(self, caminho_ld):
        return [a for a in os.listdir(caminho_ld) if os.path.splitext(a)[1] == ".xlsx"]

    def _on_selecionar(self):
        selecionado = self.listbox.curselection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma pasta primeiro.")
            return
        nome = self.listbox.get(selecionado[0])
        try:
            caminho, caminho_ld = self._get_caminhos(nome)
        except FileNotFoundError as e:
            messagebox.showerror("Erro", str(e))
            return
        arquivos_ld = self._get_arquivos_ld(caminho_ld)
        self._mostrar_frame(self._build_ui_project_decision, nome, caminho, arquivos_ld)

    @staticmethod
    def extrair_revisao(nome_base):
        nome_sem_ext = os.path.splitext(nome_base)[0]
        match = re.search(r'_R(\d+)$', nome_sem_ext)
        return int(match.group(1)) if match else -1

    def _build_ui_project_decision(self, frame, nome, caminho, arquivos_ld):
        '''
            Constrói a interface para exibir os arquivos ZIP encontrados
            e os botoes para a proxima decisao na pasta projeto
        '''
        revisoes_ld = [self.extrair_revisao(planilha) for planilha in arquivos_ld if self.extrair_revisao(planilha) >=0]
        if len(revisoes_ld) > 0:
            maior_rev = max(revisoes_ld)
        else:
            maior_rev = -1

        tk.Label(frame, text=f"Projeto: {nome}", font=("Arial", 11, "bold")).pack(pady=(15, 5))

        tk.Label(frame, text="Emissões realizadas:", font=("Arial", 10)).pack(anchor="w", padx=30, pady=(0, 7))

        lista_frame = tk.Frame(frame)
        lista_frame.pack(fill=tk.BOTH, side=tk.LEFT, expand=True, padx=20, pady=(0,100))

        scrollbar = tk.Scrollbar(lista_frame, orient=tk.VERTICAL)
        listbox = tk.Listbox(lista_frame, yscrollcommand=scrollbar.set, font=("Arial", 10), height=10)
        scrollbar.config(command=listbox.yview)

        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        for rev in range(maior_rev + 1):
            rev = str(rev + 1)
            listbox.insert(tk.END, rev.zfill(3))

        # Botao de emissao
        tk.Button(
            frame,
            text="Nova emissão",
            font= ("Arial", 10),
            width=20,
            command=lambda: self._on_emissao(caminho, nome)
        ).pack()

        # Botao de desfazer (apenas se houver emissão para desfazer)
        desfazer = DesfazerEmissao(caminho)
        if desfazer.pode_desfazer():
            tk.Button(
                frame,
                text=f"Desfazer última emissão\n({desfazer.nome_ultima_emissao()})",
                font=("Arial", 10),
                width=20,
                command=lambda: self._on_desfazer_emissao(caminho, nome),
            ).pack(pady=(15,0))
        
        # Botao exportacao
        tk.Button(
            frame,
            text="Exportar LD/GRD",
            font=("Arial", 10),
            width=20,
            command=lambda: self._on_exportacao(caminho, nome, listbox)
        ).pack(pady=(15,0))
        
        # Botao de voltar
        tk.Button(
            frame,
            text="Voltar",
            font=("Arial", 10),
            width=20,
            command=lambda: self._mostrar_frame(self._build_ui_find_project),
        ).pack(pady=(15, 0))

    def _on_desfazer_emissao(self, caminho, nome):
        desfazer = DesfazerEmissao(caminho)
        if not desfazer.pode_desfazer():
            messagebox.showinfo("Sem emissão", "Não há emissão para desfazer.")
            return

        if not desfazer.confirm_zip():
            messagebox.showerror("Erro na emissão", "O ultimo zip com a ultima emissão não existe, a emissão não será desfeita!")
            return

        confirmado = desfazer.janela_confirmacao(self.root)
        if not confirmado:
            return

        sucesso, erros = desfazer.executar()

        if sucesso:
            messagebox.showinfo(
                "Emissão desfeita",
                "A emissão foi desfeita com sucesso.\n"
                "Os arquivos foram restaurados para Para_Emissao."
            )
        else:
            messagebox.showerror(
                "Erro ao desfazer emissão",
                "Ocorreram erros:\n\n" + "\n".join(erros)
            )

        caminho, caminho_ld = self._get_caminhos(nome)
        arquivos_ld = self._get_arquivos_ld(caminho_ld)
        self._mostrar_frame(self._build_ui_project_decision, nome, caminho, arquivos_ld)

    def _on_emissao(self, caminho, nome):
        emis = Emission(caminho)
        emis.check_filename_pattern()
        dirs_to_create = emis.issued_directories()
        emis.confirm_files(dirs_to_create)
        emis.create_dirs(dirs_to_create)
        emis.ld_information = emis.get_ld_information()
        emis.check_open_files()
        emis.create_zip()
        emis.create_ld()
        emis.move_files()

        caminho, caminho_ld = self._get_caminhos(nome)
        arquivos_ld = self._get_arquivos_ld(caminho_ld)
        self._mostrar_frame(self._build_ui_project_decision, nome, caminho, arquivos_ld)

    def _on_exportacao(self, _, nome, listbox):
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione uma emissão primeiro.")
            return
        emissao = listbox.get(sel[0])
        _, caminho_ld = self._get_caminhos(nome)
        Exportar(caminho_ld, emissao)


if __name__ == "__main__":
    GerenciadorProjeto()
