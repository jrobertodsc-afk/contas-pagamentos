import os
import re
import sys
import shutil
import time
import threading
import sqlite3
import json
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from datetime import datetime, date
import zipfile
import tempfile
import smtplib

from config import *
from utils.data_processing import *
from utils.ocr_utils import *
from utils.email_service import *
from utils.file_manager import *
from utils.integrations_utils import *
from utils.extratos_processor import *
from utils.dashboard_service import *
from utils.formatters import *
from database import *
from cnab_generator import *


class AppUtilsMixin:

    def _stat_card(self, parent, label, value, color):
        # Agora usando cores dinâmicas passadas pelo app ou preto profundo por padrão
        bg_card = "#111111" # Superfície ultra escura
        frame = tk.Frame(parent, bg=bg_card, padx=20, pady=12, highlightbackground="#222222", highlightthickness=1)
        frame.pack(side="left", expand=True, fill="x", padx=(0, 10))
        tk.Label(frame, text=label, font=("Segoe UI", 8, "bold"), fg="#525252", bg=bg_card).pack(anchor="w")
        lbl = tk.Label(frame, text=value, font=("Segoe UI", 26, "bold"), fg=color, bg=bg_card)
        lbl.pack(anchor="w")
        return lbl


    def _contar_pdfs(self):
        if not os.path.exists(PASTA_ENTRADA): return 0
        count = 0
        for root, _, files in os.walk(PASTA_ENTRADA):
            if "SAIDA_ORGANIZADA" in root: continue
            count += sum(1 for f in files if f.lower().endswith('.pdf'))
        return count


    def _log(self, msg):
        if not hasattr(self, 'log_area'): return
        self.log_area.config(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        tag = "info"
        if "❌" in msg or "Erro" in msg: tag = "error"
        elif "✅" in msg: tag = "success"
        elif "⚠️" in msg: tag = "warning"
        
        linha = f"[{timestamp}] {msg}"
        self.log_area.insert("end", linha + "\n", tag)
        self.log_area.tag_configure("info", foreground="#888888")
        self.log_area.tag_configure("success", foreground="#a3e635")
        self.log_area.tag_configure("error", foreground="#f87171")
        self.log_area.tag_configure("warning", foreground="#fbbf24")
        self.log_area.see("end")
        self.log_area.config(state="disabled")
        self.root.update_idletasks()


    def _check_server_health(self):
        """Verifica se o servidor (Y:\) está acessível."""
        return os.path.exists(PASTA_SERVIDOR)


    def _status_light(self, parent, label):
        """Cria um indicador circular de status (luz)."""
        frame = tk.Frame(parent, bg=parent["bg"])
        frame.pack(side="left", padx=10)
        
        canvas = tk.Canvas(frame, width=12, height=12, bg=parent["bg"], highlightthickness=0)
        canvas.pack(side="left", padx=(0, 5))
        light = canvas.create_oval(2, 2, 10, 10, fill="#525252") # Cinza inicial
        
        tk.Label(frame, text=label, font=("Segoe UI", 7, "bold"), fg="#888888", bg=parent["bg"]).pack(side="left")
        return canvas, light


    def _atualizar_stats(self, processados, duplicados):
        self.lbl_processados.config(text=str(processados))
        self.lbl_duplicados.config(text=str(duplicados))
        self.root.update_idletasks()
