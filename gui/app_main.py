import os
import re
import sys
import shutil
import time
import threading
import sqlite3
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
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

from gui.mixins.aba_robo import AbaRoboMixin
from gui.mixins.aba_busca import AbaBuscaMixin
from gui.mixins.aba_extrato import AbaExtratoMixin
from gui.mixins.aba_contas_pagar import AbaContasPagarMixin
from gui.mixins.aba_recebiveis import AbaRecebiveisMixin
from gui.mixins.aba_autorizacao import AbaAutorizacaoMixin
from gui.mixins.aba_historico_status import AbaHistStatusMixin
from gui.mixins.aba_historico_fluxo import AbaHistFluxoMixin
from gui.mixins.aba_restituicoes import AbaRestituicoesMixin
from gui.mixins.utils_mixin import AppUtilsMixin

class ComSystemApp(
    ctk.CTk,
    AbaRoboMixin,
    AbaBuscaMixin,
    AbaExtratoMixin,
    AbaContasPagarMixin,
    AbaRecebiveisMixin,
    AbaAutorizacaoMixin,
    AbaHistStatusMixin,
    AbaHistFluxoMixin,
    AbaRestituicoesMixin,
    AppUtilsMixin
):
    def __init__(self):
        super().__init__()
        
        # ── CONFIGURAÇÕES BÁSICAS ──
        self.root = self # Compatibilidade com mixins antigos
        self.title("Com System Dashboard")
        self.geometry("1280x850")
        self.configure(fg_color="#000000")
        self.state('zoomed')
        
        self.rodando = False
        self._log_historico = []
        self._resultados_busca = []
        
        # Paleta de Cores
        self.colors = {
            "bg": "#000000",
            "surface": "#111111",
            "card": "#1A1A1A",
            "border": "#222222",
            "accent": "#CCFF00", # Lime Green Principal
            "accent_glow": "#E0FF33",
            "text": "#FFFFFF",
            "muted": "#525252",
            "error": "#FF3333",
            "warning": "#FFCC00"
        }
        
        self._build_ui()
        
        # Garante pastas de operação em background após o render inicial
        self.after(1000, ensure_directories)

    def _build_ui(self):
        # ── ESTILO GLOBAL (TTK) ──
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", 
                        background=self.colors["bg"], 
                        foreground=self.colors["text"], 
                        fieldbackground=self.colors["bg"], 
                        rowheight=35,
                        borderwidth=0,
                        font=("Segoe UI", 10))
        style.configure("Treeview.Heading", 
                        background=self.colors["surface"], 
                        foreground=self.colors["accent"], 
                        relief="flat",
                        font=("Segoe UI", 10, "bold"))
        style.map("Treeview", background=[("selected", self.colors["border"])])

        # ── HEADER (Barra Superior) ──
        self.header = ctk.CTkFrame(self, fg_color=self.colors["bg"], height=80, corner_radius=0)
        self.header.pack(fill="x", side="top")
        
        # Logo e Título
        logo_frame = ctk.CTkFrame(self.header, fg_color=self.colors["accent"], width=45, height=45, corner_radius=8)
        logo_frame.pack_propagate(False)
        logo_frame.pack(side="left", padx=(30, 15), pady=15)
        ctk.CTkLabel(logo_frame, text="CS", font=("Segoe UI", 14, "bold"), text_color="#000000").pack(expand=True)
        
        brand_f = ctk.CTkFrame(self.header, fg_color="transparent")
        brand_f.pack(side="left", pady=15)
        ctk.CTkLabel(brand_f, text="Com", font=("Segoe UI", 24, "bold"), text_color=self.colors["text"]).pack(side="left")
        ctk.CTkLabel(brand_f, text=" System", font=("Segoe UI", 24, "bold"), text_color=self.colors["accent"]).pack(side="left")
        
        # Status Operacional (Direita)
        status_f = ctk.CTkFrame(self.header, fg_color="transparent")
        status_f.pack(side="right", padx=30)
        
        self.status_dot = ctk.CTkFrame(status_f, width=10, height=10, corner_radius=5, fg_color=self.colors["accent"])
        self.status_dot.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(status_f, text="SYSTEM STATUS: OPERATIONAL", font=("Segoe UI", 10, "bold"), text_color=self.colors["accent"]).pack(side="left")

        # ── NAVEGAÇÃO (TABVIEW) ──
        # Customização para parecer abas de vidro
        self.tabview = ctk.CTkTabview(
            self, 
            fg_color="transparent",
            segmented_button_fg_color=self.colors["surface"],
            segmented_button_selected_color=self.colors["accent"],
            segmented_button_selected_hover_color=self.colors["accent_glow"],
            segmented_button_unselected_color=self.colors["card"],
            segmented_button_unselected_hover_color=self.colors["border"],
            text_color="#FFFFFF" # Texto branco para todas as abas (funciona bem no Lima e no Cinza)
        )
        self.tabview.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Configuração comum para todas as builds de abas
        aba_args = (
            self.colors["bg"], 
            self.colors["surface"], 
            self.colors["border"], 
            self.colors["accent"], 
            self.colors["accent"], 
            self.colors["warning"], 
            self.colors["text"], 
            self.colors["muted"]
        )

        # ── MONTAGEM DAS ABAS (Lazy Loading) ──
        self.tab_configs = {
            "Dashboard": self._build_aba_robo,
            "Buscar": self._build_aba_busca,
            "Fluxo": self._build_aba_extrato,
            "Logs": self._build_aba_historico_status,
            "Histórico": self._build_aba_historico_fluxo,
            "Cartões": self._build_aba_recebiveis,
            "Autorização": self._build_aba_autorizacao,
            "Manual": self._build_aba_manual,
            "Contas a Pagar": self._build_aba_contas_pagar,
            "Restituições": self._build_aba_restituicoes,
        }
        
        self.tabs_built = set()
        self.aba_args = aba_args

        # Cria os containers das abas, mas não constrói o conteúdo ainda
        for name in self.tab_configs.keys():
            self.tabview.add(name)

        # Constrói apenas a primeira aba (Dashboard) imediatamente
        self._build_tab("Dashboard")
        
        # Configura o evento de troca de aba para construir as outras sob demanda
        self.tabview.configure(command=self._on_tab_change)

    def _build_tab(self, name):
        if name in self.tabs_built: return
        
        tab_frame = self.tabview.tab(name)
        build_func = self.tab_configs[name]
        
        # Mostra um "Carregando..." temporário se não for a primeira
        if name != "Dashboard":
            loading = ctk.CTkLabel(tab_frame, text=f"Carregando {name}...", font=("Segoe UI", 12))
            loading.pack(expand=True)
            self.update_idletasks()
            loading.destroy()

        build_func(tab_frame, *self.aba_args)
        self.tabs_built.add(name)

    def _on_tab_change(self):
        selected_tab = self.tabview.get()
        self._build_tab(selected_tab)

    def _build_aba_manual(self, parent, *args):
        # Placeholder para a aba manual
        ctk.CTkLabel(parent, text="Módulo Manual em Desenvolvimento", font=("Segoe UI", 16)).pack(expand=True)

if __name__ == "__main__":
    app = ComSystemApp()
    app.mainloop()
