import os
import openpyxl
from getpass import getuser
import shutil


class Desafio_RPA():
    def __init__(self):
        # Inicialização de Variáveis
        self.dir_path = fr"C:\Users\{getuser()}\Desktop\RPA-Artigo/"
        self.qtd = len(os.listdir(self.dir_path))
        self.lista_numeros = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
        self.doc = None
        self.aba1 = None
        self.nome_arquivo = None
        self.nome_arquivo_separado = None
        self.extensao = None
        self.contador = 2
        self.valid = None
        self.primeiro_caracter = None
        self.segundo_caracter = None

    def run(self):
        # Cria Planilha
        self.criar_sheet()
        # Loop de arquivos
        for self.nome_arquivo in os.listdir(self.dir_path):
            self.nome_arquivo_separado, self.extensao = os.path.splitext(self.nome_arquivo)
            self.primeiro_caracter = self.nome_arquivo_separado[0]
            self.segundo_caracter = self.nome_arquivo_separado[1]

            # Verifica se o arquivo é .pdf e o primeiro caracter é numeral
            if self.extensao == ".pdf":
                if self.primeiro_caracter in self.lista_numeros:
                    if self.segundo_caracter in self.lista_numeros:
                        self.valid = True
                        shutil.copy(self.dir_path + self.nome_arquivo, self.dir_path + "Página " + self.primeiro_caracter + self.segundo_caracter + ".pdf")
                        self.alimentar_planilha()
                    else:
                        shutil.copy(self.dir_path + self.nome_arquivo, self.dir_path + "Página " + self.primeiro_caracter + ".pdf")
                        self.alimentar_planilha()
                else:
                    print("Arquivo é .pdf porém não começa com Numeral")
            else:
                print("Não é Arquivo .pdf")

        # Salva Planilha
        self.salvar_plan()

    # Função de Criar Planilha
    def criar_sheet(self):
        self.doc = openpyxl.Workbook()
        self.aba1 = self.doc.active
        self.aba1['A1'].value = 'Nome do Documento'
        self.aba1['B1'].value = 'Status'

    # Função de Salvar Planilha
    def salvar_plan(self):
        self.doc.save(self.dir_path + 'Relatório De Execução.xlsx')

    # Função de Alimentar Planilha
    def alimentar_planilha(self):
        self.aba1['A' + str(self.contador)].value = self.nome_arquivo
        if self.valid:
            self.aba1['B' + str(self.contador)].value = "Página " + self.primeiro_caracter + self.segundo_caracter + ".pdf"
        else:
            self.aba1['B' + str(self.contador)].value = "Página " + self.primeiro_caracter + "-Modificado.pdf"
        self.contador += 1


# Iniciar
if __name__ == "__main__":
    Desafio_RPA().run()