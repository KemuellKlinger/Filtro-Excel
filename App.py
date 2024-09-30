from DataHandler import *
class App:
    """Classe responsável pela interface gráfica e interação com o usuário."""
    
    def __init__(self, master):
        self.master = master
        self.master.title("Seleção de Arquivo")
        self.master.geometry("800x650")

        self.data_handler = DataHandler()

        # Criando os widgets da interface
        self.criar_interface()

    def criar_interface(self):
        """Configura a interface gráfica."""
        self.button_arquivo = Button(self.master, text="Selecionar Arquivo", command=self.open_file_dialog)
        self.button_arquivo.pack(pady=20)

        self.info = Label(self.master, text="Digite a chave para busca:")
        self.info.pack()

        self.chave = Entry(self.master)
        self.chave.pack()

        self.button_buscar = Button(self.master, text="Buscar", command=self.buscar)
        self.button_buscar.pack(pady=20)

        self.info3 = LabelFrame(self.master, text="RESULTADO", width=1000, height=500)
        self.info3.pack(expand=YES, fill=BOTH)

        self.info2 = Text(self.info3, wrap=WORD, state=DISABLED)
        self.info2.pack(expand=YES, fill=BOTH)

        self.button_gerar_tabela = Button(self.master, text="Gerar Tabela", command=self.gerar_tabela)
        self.button_gerar_tabela.pack(pady=20)

    def open_file_dialog(self):
        """Abre o diálogo para seleção de arquivo e processa o arquivo."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            if self.data_handler.carregar_arquivo(file_path):
                tabela_str = tabulate(self.data_handler.tabela, headers=self.data_handler.nome_colunas, tablefmt='grid')
                self.update_info_text(tabela_str)
            else:
                messagebox.showerror("Erro", "Falha ao carregar o arquivo. Verifique o formato.")

    def buscar(self):
        """Busca a chave na tabela carregada."""
        chave_texto = self.chave.get().strip()
        resultado = self.data_handler.buscar_chave(chave_texto)
        self.update_info_text(resultado)

    def gerar_tabela(self):
        """Gera um arquivo Excel com os resultados da busca."""
        resposta = self.data_handler.gerar_tabela()
        messagebox.showinfo("Resultado", resposta)

    def update_info_text(self, text):
        """Atualiza o campo de resultado."""
        self.info2.config(state=NORMAL)
        self.info2.delete(1.0, END)
        self.info2.insert(END, text)
        self.info2.config(state=DISABLED)

