from config import COLORS
from src.services.Excel_service import Excel_Services
from .Table_UI import Table_UI
from .Filter_UI import Filter_UI

import tkinter as tk
from tkinter import ttk, messagebox

class App_UI:
  ### Atributes
  models = ["clientes.xlsx"]
  
  ### Constructor
  def __init__(self, root):
    ### Initialize app
    self.root = root
    self.root.title("Cadastro de Clientes")
    self.root.state('zoomed')  # Tela cheia
    self.root.geometry("1400x600")
    self.root.configure(bg=COLORS["bg"])

    self.main_frame = tk.Frame(self.root, bg=COLORS["bg"], bd=0, highlightthickness=0)
    self.main_frame.pack(fill=tk.BOTH, expand=True)

    ### Validations
    self.validate_DB()
    self.build_widgets()

  ### Methods
  def validate_DB(self):
    is_DB = Excel_Services.validate_exist_DB()
    if not is_DB:
      for model in self.models:
        Excel_Services.create_database_file(model)

  def build_widgets(self):
    Filter_UI(self)
    Table_UI(self)