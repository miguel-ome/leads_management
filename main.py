import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook
import os
import sys

if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

excel_path = os.path.join(BASE_DIR, "clientes.xlsx")

wb = load_workbook(excel_path)
ws = wb.active

larguras_personalizadas = {
    "ID": 40,
    "Nome": 200,
    "Rua": 150,
    "Número": 60,
    "Bairro": 120,
    "Cidade": 120,
    "Estado": 50,
    "CEP": 80,
    "Celular": 110,
    "Email": 180,
    "Salário": 80,
    "Categoria": 100,
    "Objetivo de Compra": 180,
    "Observações": 200,
}


# Paleta de cores
COR_FUNDO = "#465F87"
COR_HEADER = "#D8F2F0"
COR_BOTAO = "#89ABD9"
COR_TEXTO = "#000000"
COR_FILTRO_BG = "#2E3F5F"

colunas = ["ID", "Nome", "Rua", "Número", "Bairro", "Cidade", "Estado", "CEP",
           "Celular", "Email", "Salário", "Categoria", "Objetivo de Compra", "Observações"]

def carregar_clientes():
    return [
        [cell if cell is not None else "" for cell in row]
        for row in ws.iter_rows(min_row=2, values_only=True)
    ]


def salvar_cliente(dados):
    ws.append(dados)
    wb.save(excel_path)

def atualizar_cliente(linha, dados):
    for i, valor in enumerate(dados):
        ws.cell(row=linha, column=i+1).value = valor
    wb.save(excel_path)

def excluir_cliente(linha):
    ws.delete_rows(linha)
    wb.save(excel_path)

def centralizar_popup(janela, largura=500, altura=500):
    janela.update_idletasks()
    x = (janela.winfo_screenwidth() // 2) - (largura // 2)
    y = (janela.winfo_screenheight() // 2) - (altura // 2)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")

def criar_interface():
    def atualizar_tabela(filtros=None):
        for item in tree.get_children():
            tree.delete(item)

        for cliente in carregar_clientes():
            if filtros:
                match = True
                for i, entrada in filtros.items():
                    # Trata o campo de salário (índice 11)
                    if isinstance(entrada, tuple) and len(entrada) == 2:
                        operador, valor_input = entrada
                        operador_map = {
                            "Maior que": ">",
                            "Menor que": "<",
                            "Igual a": "="
                        }
                        operador_real = operador_map.get(operador)

                        if operador_real and valor_input.strip():
                            try:
                                salario_cliente = float(str(cliente[i]).replace(",", ".").strip() or "0")
                                valor_input = float(valor_input.replace(",", ".").strip())

                                if operador_real == ">" and not (salario_cliente > valor_input):
                                    match = False
                                elif operador_real == "<" and not (salario_cliente < valor_input):
                                    match = False
                                elif operador_real == "=" and not (salario_cliente == valor_input):
                                    match = False
                            except (ValueError, TypeError) as e:
                                print("Erro ao comparar salário:", e)
                                match = False
                        continue  # pula pro próximo filtro
                    # Trata os outros campos (strings)
                    if entrada:  # só filtra se houver algo
                        valor_input = str(entrada).strip().lower()
                        valor_cliente = str(cliente[i]).strip().lower()
                        print(f"Comparando cliente[{i}] = '{valor_cliente}' com filtro = {valor_input}")
                        if valor_input not in valor_cliente:
                            match = False
                            break
                if not match:
                    continue
            tree.insert("", "end", values=cliente)


    def abrir_formulario(cliente=None, linha=None):
        form = tk.Toplevel(root)
        form.title("Atualização de Cliente" if cliente else "Cadastro de Cliente")
        form.configure(bg=COR_FUNDO)

        container = tk.Frame(form, bg=COR_FUNDO)
        container.pack(expand=True)

        labels = ["Nome", "Rua", "Número", "Bairro", "Cidade", "Estado", "CEP",
                "Celular", "Email", "Salário", "Categoria", "Objetivo de Compra", "Observações"]
        entradas = {}

        for i, label in enumerate(labels):
            tk.Label(container, text=label, bg=COR_FUNDO, fg="white", font=("Arial", 10)).grid(
                row=i, column=0, padx=10, pady=5, sticky="e"
            )
            entrada = tk.Entry(container, width=30)
            entrada.grid(row=i, column=1, padx=10, pady=5)
            entradas[label] = entrada

        if cliente:
            for i, valor in enumerate(cliente[1:], start=1):  # Ignorando o ID
                entradas[labels[i - 1]].insert(0, valor)

        def salvar():
            dados = [entrada.get() for entrada in entradas.values()]
            if cliente:
                atualizar_cliente(linha, [cliente[0]] + dados)  # Mantendo o ID original
                messagebox.showinfo("Atualizado", "Cliente atualizado com sucesso!")
            else:
                novo_id = len(carregar_clientes()) + 1
                salvar_cliente([novo_id] + dados)
                messagebox.showinfo("Cadastrado", "Cliente cadastrado com sucesso!")
            atualizar_tabela()
            form.destroy()

        btn_salvar = tk.Button(
            container, text="Salvar", command=salvar,
            bg=COR_BOTAO, fg=COR_TEXTO, font=("Arial", 10)
        )
        btn_salvar.grid(row=len(labels), column=0, columnspan=2, pady=10)

        centralizar_popup(form, largura=400, altura=600)
        form.grab_set()

    def deletar_cliente():
        item = tree.selection()
        if not item:
            messagebox.showwarning("Seleção", "Selecione um cliente para excluir.")
            return
        confirm = messagebox.askyesno("Confirmação", "Deseja realmente excluir o cliente?")
        if confirm:
            index = tree.index(item)
            excluir_cliente(index + 2)
            atualizar_tabela()
            messagebox.showinfo("Excluído", "Cliente excluído com sucesso!")

    def editar_cliente():
        item = tree.selection()
        if not item:
            messagebox.showwarning("Seleção", "Selecione um cliente para editar.")
            return
        dados = tree.item(item)['values']
        linha = tree.index(item) + 2
        abrir_formulario(dados, linha)

    def aplicar_filtros():
        filtros = {}
        for i, entrada in entradas_filtro.items():
            if isinstance(entrada, tuple):  # campo do salário, com operador + valor
                operador, valor = entrada
                filtros[i] = (operador.get(), valor.get())

            else:
                filtros[i] = entrada.get()
        atualizar_tabela(filtros)



    def toggle_filtros():
        if filtro_frame.winfo_ismapped():
            filtro_frame.pack_forget()
        else:
            filtro_frame.pack(side="top", fill="x", pady=5)

    global root
    root = tk.Tk()
    root.title("Sistema de Clientes")
    root.configure(bg=COR_FUNDO)
    root.state('zoomed')

    btn_top = tk.Frame(root, bg=COR_FUNDO)
    btn_top.pack(side="top", fill="x", pady=(10, 0))

    tk.Button(btn_top, text="Filtros", font=("Arial", 10, "bold"),
              command=toggle_filtros, bg=COR_BOTAO, fg=COR_TEXTO,
              bd=0, highlightthickness=0, width=15, height=2).pack(side="left", padx=10)

    filtro_frame = tk.Frame(root, bg=COR_FILTRO_BG)
    entradas_filtro = {}

    filtros_internos_frame = tk.Frame(filtro_frame, bg=COR_FILTRO_BG)
    filtros_internos_frame.pack(anchor="center", pady=5)

    col = 0  # índice manual da coluna

    for i, coluna in enumerate(colunas):
        larguras_filtros = {
            "ID": 10,
            "Nome": 18,
            "Rua": 18,
            "Número": 6,
            "Bairro": 14,
            "Cidade": 14,
            "Estado": 5,
            "CEP": 10,
            "Celular": 12,
            "Email": 22,
            "Salário": 10,
            "Categoria": 12,
            "Objetivo de Compra": 20,
            "Observações": 25,
        }

        largura = larguras_filtros.get(coluna, 12)

        if coluna == "Salário":
            # Label ocupa duas colunas
            lbl = tk.Label(filtros_internos_frame, text=coluna, bg=COR_FILTRO_BG, fg="white", font=("Arial", 9))
            lbl.grid(row=0, column=col, columnspan=2, padx=5, pady=2)

            operador_combo = ttk.Combobox(
                filtros_internos_frame,
                values=["", "Maior que", "Menor que", "Igual a"],
                width=10,
                state="readonly"
            )
            operador_combo.set("Maior que")
            operador_combo.grid(row=1, column=col, padx=(5, 2), pady=2)

            entrada_salario = tk.Entry(filtros_internos_frame, width=10)
            entrada_salario.grid(row=1, column=col + 1, padx=(2, 5), pady=2)

            entradas_filtro[i] = (operador_combo, entrada_salario)
            col += 2  # pular uma coluna extra porque "Salário" ocupa duas
        else:
            lbl = tk.Label(filtros_internos_frame, text=coluna, bg=COR_FILTRO_BG, fg="white", font=("Arial", 9))
            lbl.grid(row=0, column=col, padx=5, pady=2)

            entrada = tk.Entry(filtros_internos_frame, width=largura)
            entrada.grid(row=1, column=col, padx=5, pady=2)
            entradas_filtro[i] = entrada
            col += 1

    btn_aplicar = tk.Button(filtros_internos_frame, text="Aplicar Filtros", command=aplicar_filtros,
                            font=("Arial", 10, "normal"), bg=COR_BOTAO, fg=COR_TEXTO,
                            bd=0, highlightthickness=0, width=15, height=2)
    btn_aplicar.grid(row=2, column=0, columnspan=len(colunas), pady=10)

    tree = ttk.Treeview(root, columns=colunas, show="headings")

    for col in colunas:
        tree.heading(col, text=col)
        largura = larguras_personalizadas.get(col, 100)
        tree.column(col, width=largura)


    tree.pack(expand=True, fill="both", padx=10, pady=10)

    btn_frame = tk.Frame(root, bg=COR_FUNDO)
    btn_frame.pack(side="bottom", fill="x", pady=10)

    tk.Button(btn_frame, text="Novo Cliente", font=("Arial", 10, "normal"), command=lambda: abrir_formulario(cliente=""),
              bg=COR_BOTAO, fg=COR_TEXTO, bd=0, highlightthickness=0, width=15, height=2).pack(side="left", padx=10)
    tk.Button(btn_frame, text="Editar Cliente", font=("Arial", 10, "normal"), command=editar_cliente,
              bg=COR_BOTAO, fg=COR_TEXTO, bd=0, highlightthickness=0, width=15, height=2).pack(side="left", padx=10)
    tk.Button(btn_frame, text="Excluir Cliente", font=("Arial", 10, "normal"), command=deletar_cliente,
              bg=COR_BOTAO, fg=COR_TEXTO, bd=0, highlightthickness=0, width=15, height=2).pack(side="left", padx=10)

    atualizar_tabela()
    root.mainloop()

criar_interface()
