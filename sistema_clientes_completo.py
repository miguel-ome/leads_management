import tkinter as tk
from tkinter import ttk
import openpyxl

ARQUIVO_EXCEL = 'clientes.xlsx'

class SistemaClientes:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Clientes")
        self.root.state('zoomed')

        self.dados = []
        self.filtros = {}
        self.colunas = [
            "ID", "Nome", "Rua", "Número", "Bairro", "Cidade", "Estado",
            "CEP", "Celular", "Email", "Salário", "Categoria", "Objetivo de Compra", "Observações"
        ]

        self.frame_topo = tk.Frame(self.root, bg="#D8F2F0")
        self.frame_topo.pack(fill=tk.X)

        self.canvas = tk.Canvas(self.frame_topo, height=60, bg="#D8F2F0")
        self.canvas.pack(fill=tk.X, side=tk.TOP)

        self.filter_entries = []

        self.tree = ttk.Treeview(self.root, columns=self.colunas, show="headings")
        for col in self.colunas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="w")

        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<Configure>", self.ajustar_filtros)
        for col in self.colunas:
            self.tree.heading(col, command=lambda c=col: self.treeview_column_resized(c))

        self.scroll_y = ttk.Scrollbar(self.tree, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scroll_y.set)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.frame_botoes = tk.Frame(self.root, bg="#465F87")
        self.frame_botoes.pack(fill=tk.X)

        self.btn_criar = tk.Button(self.frame_botoes, text="Criar", command=self.criar_cliente)
        self.btn_criar.pack(side=tk.LEFT, padx=10, pady=5)
        self.btn_editar = tk.Button(self.frame_botoes, text="Editar", command=self.editar_cliente)
        self.btn_editar.pack(side=tk.LEFT, padx=10)
        self.btn_excluir = tk.Button(self.frame_botoes, text="Excluir", command=self.excluir_cliente)
        self.btn_excluir.pack(side=tk.LEFT, padx=10)

        self.carregar_dados()
        self.criar_filtros()

    def carregar_dados(self):
        self.dados.clear()
        wb = openpyxl.load_workbook(ARQUIVO_EXCEL)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            self.dados.append(dict(zip(self.colunas, row)))
        self.atualizar_tabela()

    def atualizar_tabela(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        dados_filtrados = self.filtrar_dados()
        for dado in dados_filtrados:
            self.tree.insert("", tk.END, values=[dado.get(col, "") for col in self.colunas])

    def criar_filtros(self):
        for i, col in enumerate(self.colunas):
            lbl = tk.Label(self.canvas, text=col, font=("Arial", 9, "bold"), bg="#D8F2F0")
            lbl.place(x=i*100, y=0, width=100)
            entry = ttk.Entry(self.canvas)
            entry.place(x=i*100, y=20, width=100)
            entry.bind("<KeyRelease>", lambda e, c=col: self.atualizar_filtros(c, e.widget.get()))
            self.filter_entries.append(entry)

    def ajustar_filtros(self, event=None):
        x = 0
        for i, col in enumerate(self.colunas):
            width = self.tree.column(col, width=None)
            self.canvas.coords(self.canvas.winfo_children()[i*2], x, 0)
            self.canvas.itemconfig(self.canvas.winfo_children()[i*2], width=width)
            self.filter_entries[i].place(x=x, y=20, width=width)
            x += width

    def atualizar_filtros(self, coluna, valor):
        self.filtros[coluna] = valor
        self.atualizar_tabela()

    def filtrar_dados(self):
        resultado = self.dados
        for col, val in self.filtros.items():
            if val:
                resultado = [d for d in resultado if val.lower() in str(d.get(col, '')).lower()]
        return resultado

    def treeview_column_resized(self, col):
        self.ajustar_filtros()

    def criar_cliente(self):
        print("Criar cliente")

    def editar_cliente(self):
        print("Editar cliente")

    def excluir_cliente(self):
        print("Excluir cliente")

if __name__ == "__main__":
    root = tk.Tk()
    app = SistemaClientes(root)
    root.mainloop()
