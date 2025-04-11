import tkinter as tk
from tkinter import ttk, messagebox

from config import COLORS, FIELDS_TABLE_LEADS, WIDTH_COLUMN
from services.Filter_service import Filter_Service

class Filter_UI:
  def __init__(self, app_ui):
    self.root = app_ui.root
    self.entries_filtro = []

    # Parte superior: filtros + botão Filtrar
    self.filters_frame = tk.Frame(app_ui.main_frame, bg=COLORS["bg"])
    self.filters_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

    for i, campo in enumerate(FIELDS_TABLE_LEADS):
        lbl = tk.Label(self.filters_frame, text=campo, bg=COLORS["btn"], fg='white')
        lbl.grid(row=0, column=i, sticky="ew")

        if campo == "Salário":
            frame_salario = tk.Frame(self.filters_frame, bg=COLORS["bg"])
            frame_salario.grid(row=1, column=i, sticky="ew", padx=1, pady=1)
            
            var_operador = tk.StringVar(value="=")
            operador = ttk.Combobox(frame_salario, textvariable=var_operador, width=2, state="readonly")
            operador['values'] = ("=", ">", "<", ">=", "<=")
            operador.grid(row=0, column=0)

            entry_valor = tk.Entry(frame_salario, width=15)
            entry_valor.grid(row=0, column=1)

            self.entries_filtro.append((var_operador, entry_valor))
        else:
            entry = tk.Entry(self.filters_frame, width=WIDTH_COLUMN[campo])
            entry.grid(row=1, column=i, sticky="ew", padx=1, pady=1)
            self.entries_filtro.append(entry)

        self.filters_frame.grid_columnconfigure(i, weight=1)


    botao_filtro = tk.Button(self.filters_frame, text="Filtrar", command=Filter_Service.filter_data(self), bg=COLORS["btn"], fg='white')
    botao_filtro.grid(row=1, column=len(FIELDS_TABLE_LEADS), padx=(10, 0))