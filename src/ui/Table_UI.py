from config import COLORS, FIELDS_TABLE_LEADS, WIDTH_COLUMN

import tkinter as tk
from tkinter import ttk, messagebox


class Table_UI:
  def __init__(self, app_ui):
    self.root = app_ui.root

    frame_tabela = tk.Frame(app_ui.main_frame, bg=COLORS["bg"])
    frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))

    # Scrollbars
    scrollbar_y = ttk.Scrollbar(frame_tabela, orient="vertical")
    scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

    scrollbar_x = ttk.Scrollbar(frame_tabela, orient="horizontal")
    scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

    # Tabela (Treeview)
    self.tabela = ttk.Treeview(
        frame_tabela,
        columns=FIELDS_TABLE_LEADS,
        show="headings",
        yscrollcommand=scrollbar_y.set,
        xscrollcommand=scrollbar_x.set
    )
    for campo in FIELDS_TABLE_LEADS:
        self.tabela.heading(campo, text=campo)
        self.tabela.column(campo, width=WIDTH_COLUMN[campo] * 10, anchor="w")

    self.tabela.pack(fill=tk.BOTH, expand=True)

    scrollbar_y.config(command=self.tabela.yview)
    scrollbar_x.config(command=self.tabela.xview)

    # Botões CRUD
    frame_botoes = tk.Frame(self.root, bg=COLORS["bg"])
    frame_botoes.pack(pady=10, padx=20)

    tk.Button(frame_botoes, text="Criar", command=lambda: print("criar"), bg=COLORS["btn"], width=12, fg='white', bd=0, highlightthickness=0).pack(side=tk.LEFT, padx=10)
    tk.Button(frame_botoes, text="Editar", command=lambda: print("editar"), bg=COLORS["btn"], width=12, fg='white', bd=0, highlightthickness=0).pack(side=tk.LEFT, padx=10)
    tk.Button(frame_botoes, text="Excluir", command=lambda: print("excluir"), bg=COLORS["btn"], width=12, fg='white', bd=0, highlightthickness=0).pack(side=tk.LEFT, padx=10)
