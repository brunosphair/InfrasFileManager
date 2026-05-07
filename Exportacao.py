import os
import re
import tkinter as tk
from tkinter import messagebox
import openpyxl
import xlwings as xw
from pathlib import Path

FONTE = ("Arial", 10)

class Exportar():

    def __init__(self, base_path):
        self.base_path = Path(base_path)
        self.padrao_re = re.compile(r'^IFS-\d{4}-\d{3}-(\w{3}-\w{2}-\d{4}|\w-\w{2}-\d{5})(_R\d+)?$')

        self.planilha = self.encontrar_planilha()
        self.pasta_salvamento = self.caminho_salvamento()

        self.janela = tk.Tk()
        self.janela.configure(padx=50, pady=15)
        self.janela.resizable(False, False)

        self.frame_atual = None
        self._mostrar_frame(self.interface)

        self.janela.mainloop()
    
    def _mostrar_frame(self, builder, *args):
        '''
            Limpa o frame atual e constrói um novo usando a função builder fornecida.
        '''
        if self.frame_atual:
            self.frame_atual.destroy()
        self.frame_atual = tk.Frame(self.janela)
        self.frame_atual.grid(column=0, row=0)
        # Chama a classe enviada como parametro no _mostrar_frame
        builder(self.frame_atual, *args)


    def caminho_salvamento(self):
        path = self.base_path.parent.absolute()
        parent_path = path.parent.absolute()
        parent_parent_path = parent_path.parent.absolute()
        caminho = Path(os.path.join(parent_parent_path, "3_Emitidos", "00_LDs"))

        if not caminho.is_dir():
            caminho = Path(os.path.join(self.base_path, "Exportados"))

        if not caminho.is_dir():
            messagebox.showerror("Erro", "Pasta para salvar LD nao existente")
            raise ValueError("Pasta de salvamento nao existente", caminho)

        print(caminho)
        return caminho


    @staticmethod
    def extrair_revisao(nome_base):
        match = re.search(r'_R(\d+)$', nome_base)
        return int(match.group(1)) if match else 0
    
    def encontrar_planilha(self):
        candidatos = []

        for nome in os.listdir(self.base_path):
            raiz, ext = os.path.splitext(nome)
            if ext.lower() == ".xlsx" and self.padrao_re.match(raiz):
                candidatos.append((nome, self.extrair_revisao(raiz)))
        if not candidatos:
            messagebox.showwarning('Resultado', 'Nenhuma planilha encontrada com o padrão IFS.')
            return
        maior_rev = max(r for _, r in candidatos)
        planilha_maior_rev = [nome for nome, r in candidatos if r == maior_rev][0]

        return planilha_maior_rev
    
    def gerar_nomenclatura_grd(self, nmr_grd):
        if self.planilha[14] == "-": 
            nmr_grd = nmr_grd.zfill(5)
        else:
            nmr_grd = nmr_grd.zfill(4)

        nome_planilha_sem_ext = os.path.splitext(self.planilha)[0]
        list_planilha = nome_planilha_sem_ext.split("-")

        list_planilha[4] = "GR"
        list_planilha[5] = nmr_grd
        nomenclatura_grd = "-".join(list_planilha) + ".xlsx"

        return nomenclatura_grd


    def interface(self, frame):

        mensagem = tk.Label(frame, text="Selecione o que gostaria de fazer.", font=FONTE)
        mensagem.grid(column=1,row=1, columnspan=2, pady=(0,10))

        botao_exportar_ld = tk.Button(frame, text="Exportar LD",font=FONTE, command=self.exportar_ld)
        botao_exportar_ld.grid(column=1,row=2)

        botao_exportar_grd = tk.Button(frame, text="Exportar GRD",font=FONTE, command=lambda: self._mostrar_frame(self.interface_export_grd))
        botao_exportar_grd.grid(column=2, row=2)

        

    def exportar_ld(self):

        app = xw.App(visible=False)
        try:
            wb = app.books.open(str(self.base_path / self.planilha))

            for nome_aba in ["LD", "Capa"]:
                aba = wb.sheets[nome_aba]
                rng = aba.used_range
                rng.copy()
                rng.paste(paste="values")

            aba_ld = wb.sheets["LD"]
            aba_ld.api.Columns("BZ:XFD").Delete()

            abas_para_deletar = [s for s in wb.sheets if s.name not in ("LD", "Capa")]
            for s in abas_para_deletar:
                s.delete()

            caminho_salvamento_ld = os.path.join(self.pasta_salvamento, self.planilha)
            wb.save(caminho_salvamento_ld)
            wb.close()
        finally:
            app.quit()

        messagebox.showinfo("Resultado", "Planilha exportada com sucesso!")
        self.janela.destroy()

    def interface_export_grd(self, frame):

        mensagem = tk.Label(frame, text="Insira o numero da GRD que\nvoce gostaria de exportar", font=FONTE)
        mensagem.grid(column=0,row=0,columnspan=2)

        self.entrada_grd = tk.Entry(frame, font=FONTE)
        self.entrada_grd.grid(column=0,row=1, columnspan=2, pady=(10,0))

        botao_enviar = tk.Button(frame,
            text="Exportar",
            font=FONTE,
            command= self.exportar_grd
        ).grid(column=0,row=2, pady=(10,0))

        botao_voltar = tk.Button(
            frame,
            text="Voltar",
            font=FONTE,
            command=lambda: self._mostrar_frame(self.interface)
        ).grid(column=1,row=2, pady=(10,0))
    

    def validar_titulo_vazio(self, ws):
        contador = 26
        while True:
            celula_doc = ws[f"B{contador}"]
            if celula_doc.value is None:
                return False
            celula_titulo = ws[f"E{contador}"]
            if celula_titulo.value is None:
                return True
            contador += 1


    def exportar_grd(self):

        nmr_grd = self.entrada_grd.get().zfill(3)
        planilha = openpyxl.load_workbook(str(self.base_path / self.planilha), data_only=True)

        try: 
            int(nmr_grd)
            planilha[f"GRD-{nmr_grd}"]
        except ValueError:
            messagebox.showerror("Erro!", "Insira um número válido")
            return
        except KeyError:
            messagebox.showerror("Erro!", "GRD nao encontrada")
            return
        
        nome_aba_grd = f"GRD-{nmr_grd}"

        if self.validar_titulo_vazio(planilha[nome_aba_grd]):
            continuar = messagebox.askyesno(
                "Atenção",
                "Há documentos sem título na GRD. Deseja continuar com a exportação?"
            )
            if not continuar:
                return

        nome_planilha_grd_exportada = self.gerar_nomenclatura_grd(nmr_grd)

        app = xw.App(visible=False)
        try:
            wb = app.books.open(str(self.base_path / self.planilha))
            aba = wb.sheets[nome_aba_grd]

            # Copia o range inteiro e cola como valores (substitui fórmulas)
            rng = aba.used_range
            rng.copy()
            rng.paste(paste="values")

            # Deleta todas as outras abas
            for s in wb.sheets:
                if s.name != nome_aba_grd:
                    s.delete()

            caminho_salvamento_grd = os.path.join(self.pasta_salvamento, nome_planilha_grd_exportada)
            wb.save(caminho_salvamento_grd)
            wb.close()
        finally:
            app.quit()

        messagebox.showinfo("Resultado", "Planilha exportada com sucesso!")
        self.janela.destroy()

