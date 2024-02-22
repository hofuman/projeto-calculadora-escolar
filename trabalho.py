import tkinter as tk
from tkinter import ttk
import pandas as pd

class PrincipalRAD():
    def __init__(self, win):
        #componentes
        self.lblNome=tk.Label(win, text= 'Nome do Aluno:')
        self.lblNota1=tk.Label(win, text='Nota 1: ')
        self.lblNota2=tk.Label(win, text='Nota 2: ')
        self.lblMedia=tk.Label(win, text='Media: ')
        self.txtNome=tk.Entry(win, bd=3)
        self.txtNota1=tk.Entry(win)
        self.txtNota2=tk.Entry(win)
        self.btnCalcular=tk.Button(win, text="Calcular Média", command=self.fCalcularMedia)

        ### Componentes Treeview
        self.dadosColunas = ("Alunos", "Nota 1", "Nota 2", "Média", "Situação")
        self.treeMedia = ttk.Treeview(win,
                                      columns=self.dadosColunas,
                                      selectmode='browse')
        self.verscrlbar = ttk.Scrollbar(win,
                                        orient="vertical",
                                        command=self.treeMedia.yview)
        self.verscrlbar.pack(side='right', fill = 'y')
        self.treeMedia.configure(yscrollcommand=self.verscrlbar.set)

        self.treeMedia.heading('Alunos', text='Aluno')
        self.treeMedia.heading('Nota 1', text='Nota 1')
        self.treeMedia.heading('Nota 2', text='Nota 2')
        self.treeMedia.heading('Média', text='Média')
        self.treeMedia.heading('Situação', text='Situação')

        self.treeMedia.column('#0', stretch=tk.NO, minwidth=0, width=0)
        for col in self.dadosColunas:
            self.treeMedia.column(col, stretch=tk.NO, minwidth=0, width=100)

        self.treeMedia.pack(padx=10, pady=10)

        ### Posicionamento de janela

        self.lblNome.place(x=100, y=50)
        self.txtNome.place(x=200, y=50)

        self.lblNota1.place(x=100, y=100)
        self.txtNota1.place(x=200, y=100)

        self.lblNota2.place(x=100, y=150)
        self.txtNota2.place(x=200, y=150)

        self.btnCalcular.place(x=100, y=200)

        self.treeMedia.place(x=100, y=300)
        self.verscrlbar.place(x=805, y=300, height=225)

        self.id = 0
        self.iid = 0

        self.carregarDadosIniciais()

    ### Funções

    def carregarDadosIniciais(self):
        try:
            fsave = 'planilhaAlunos.xlsx'
            dados = pd.read_excel(fsave)
            print("****** dados disponíveis ********")
            print(dados)

            nn = len(dados['Alunos'])
            for i in range(nn):
                nome = dados['Alunos'][i]
                nota1 = str(dados['Nota 1'][i])
                nota2 = str(dados['Nota 2'][i])
                media = str(dados['Média'][i])
                situacao = dados['Situação'][i]

                self.treeMedia.insert('', 'end',
                                      values=(nome,
                                              nota1,
                                              nota2,
                                              media,
                                              situacao))
                self.iid += 1
                self.id += 1

        except:
            print('Ainda não existem dados para carregar.')

    ## Salvar dados para planilha exel

    def fSalvarDados(self):
        try:
            fsave = 'planilhaAlunos.xlsx'
            dados = []
            for line in self.treeMedia.get_children():
                lstDados = []
                for value in self.treeMedia.item(line)['values']:
                    lstDados.append(value)
                dados.append(lstDados)
            df = pd.DataFrame(dados, columns=self.dadosColunas)
            df.to_excel(fsave, index=False)
            print('Dados salvos!')

        except:
            print('Não foi possível salvar os dados')

    ## verificar média do aluno
    def fVerificarSituacao(self, nota1, nota2):
        media = (nota1 + nota2) / 2
        if media >= 7.0:
            situacao = 'Aprovado'
        elif media >= 5.0:
            situacao ='Em Recuperação'
        else:
            situacao = 'Reprovado'

        return media, situacao

    ## Imprimir os dados

    def fCalcularMedia(self):
        try:
            nome = self.txtNome.get()
            nota1 = float(self.txtNota1.get())
            nota2 = float(self.txtNota2.get())
            media, situacao = self.fVerificarSituacao(nota1, nota2)

            self.treeMedia.insert('', 'end',
                                  iid=self.iid,
                                  values=(nome,
                                          str(nota1),
                                          str(nota2),
                                          str(media),
                                          situacao))
            self.iid += 1
            self.id += 1

            self.fSalvarDados()
        except ValueError:
            print("Entre com valores válidos")
        finally:
            self.txtNome.delete(0, 'end')
            self.txtNota1.delete(0, 'end')
            self.txtNota2.delete(0, 'end')

### programa principal

janela = tk.Tk()
principal = PrincipalRAD(janela)
janela.title('Bem Vindo ao RAD')
janela.geometry("820x600+10+10")
janela.mainloop()
