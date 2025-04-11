import tkinter as tk

class Table_service:
  @staticmethod
  def atualizar_tabela(self, dados):
        for item in self.tabela.get_children():
            self.tabela.delete(item)
        for row in dados:
            # Substitui valores None por string vazia antes de inserir
            linha_formatada = [cell if cell is not None else "" for cell in row]
            self.tabela.insert("", tk.END, values=linha_formatada)