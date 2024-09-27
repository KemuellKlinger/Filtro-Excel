import pandas as pd
from tkinter import *
from tkinter import filedialog, messagebox
from tabulate import tabulate


class DataHandler:
    """Classe responsável pela manipulação dos dados."""
    
    def __init__(self):
        self.tabela = None
        self.nome_colunas = None
        self.tabela_formatada = []

    def carregar_arquivo(self, file_path):
        """Carrega um arquivo Excel e armazena as colunas."""
        try:
            self.tabela = pd.read_excel(file_path)
            self.nome_colunas = self.tabela.columns
            return True
        except Exception as e:
            print(f"Erro ao carregar arquivo: {e}")
            return False

    def buscar_chave(self, chave):
        """Busca uma chave no DataFrame e retorna os resultados formatados."""
        if self.tabela is None:
            return "Nenhuma tabela carregada!"

        if len(chave) < 3:
            return "A chave deve ter ao menos 3 caracteres!"

        self.tabela_formatada = []
        for _, linha in self.tabela.iterrows():
            if chave in str(linha):
                self.tabela_formatada.append(linha.tolist())

        if not self.tabela_formatada:
            return "Dados não encontrados!"

        return tabulate(self.tabela_formatada, headers=self.nome_colunas, tablefmt='grid')

    def gerar_tabela(self, file_name="TabelaGerada.xlsx"):
        """Gera um arquivo Excel com os dados formatados."""
        if not self.tabela_formatada:
            return "Nenhum dado para gerar tabela!"
        try:
            df = pd.DataFrame(self.tabela_formatada, columns=self.nome_colunas)
            df.to_excel(file_name, index=False)
            return f"Tabela gerada com sucesso: {file_name}"
        except Exception as e:
            return f"Erro ao gerar a tabela: {e}"


