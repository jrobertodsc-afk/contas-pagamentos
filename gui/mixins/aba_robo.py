import os
import re
import sys
import shutil
import time
import threading
import sqlite3
import json
import tkinter as tk
from tkinter import messagebox, filedialog
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


class AbaRoboMixin:

    def _build_aba_robo(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        # ── SYSTEM OVERVIEW (Stats Cards) ──
        stats_frame = ctk.CTkFrame(parent, fg_color="transparent")
        stats_frame.pack(fill="x", padx=30, pady=(20, 10))
        
        # Grid layout para os cartões
        stats_frame.grid_columnconfigure((0, 1, 2), weight=1, pad=20)
        
        self.lbl_processados = self._modern_stat_card(stats_frame, "PROCESSED TODAY", "0", accent, 0)
        self.lbl_duplicados   = self._modern_stat_card(stats_frame, "DUPLICATE ITEMS", "0", yellow, 1)
        
        entrada_count = self._contar_pdfs()
        self.lbl_entrada = self._modern_stat_card(stats_frame, "QUEUE STATUS", str(entrada_count), "#FFFFFF", 2)

        # ── ACTIVITY LOG (Console) ──
        log_container = ctk.CTkFrame(parent, fg_color=surface, corner_radius=12, border_width=1, border_color=border)
        log_container.pack(fill="both", expand=True, padx=30, pady=10)
        
        log_header = ctk.CTkFrame(log_container, fg_color="transparent", height=40)
        log_header.pack(fill="x", padx=20, pady=(10, 0))
        
        ctk.CTkLabel(log_header, text="📋 ACTIVITY LOG", font=("Segoe UI", 11, "bold"), text_color=accent).pack(side="left")
        
        # Botões de controle do log (Estilo Moderno)
        ctk.CTkButton(log_header, text="LIVE", width=60, height=24, corner_radius=12, fg_color=accent, text_color="#000000", font=("Segoe UI", 10, "bold")).pack(side="right", padx=5)
        ctk.CTkButton(log_header, text="CLEAR", width=60, height=24, corner_radius=12, fg_color="transparent", border_width=1, border_color=muted, text_color=text, font=("Segoe UI", 10)).pack(side="right", padx=5)

        self.log_area = ctk.CTkTextbox(
            log_container, 
            font=("Consolas", 11),
            fg_color="#000000", 
            text_color="#FFFFFF",
            border_width=0,
            padx=20, pady=20
        )
        self.log_area.pack(fill="both", expand=True, padx=10, pady=10)
        self.log_area.configure(state="disabled")

        # ── FOOTER / CONTROLES ──
        footer_f = ctk.CTkFrame(parent, fg_color="transparent", height=100)
        footer_f.pack(fill="x", padx=30, pady=(10, 20))
        
        # Info de Sessão (Esquerda) - Estrutura de grid para info detalhada
        session_info = ctk.CTkFrame(footer_f, fg_color="transparent")
        session_info.pack(side="left", fill="y")
        
        info_grid = ctk.CTkFrame(session_info, fg_color="transparent")
        info_grid.pack(side="left")
        
        # Coluna 1
        f1 = ctk.CTkFrame(info_grid, fg_color="transparent")
        f1.pack(side="left", padx=(0, 20))
        ctk.CTkLabel(f1, text="👤 USER", font=("Segoe UI", 8, "bold"), text_color=accent).pack(anchor="w")
        ctk.CTkLabel(f1, text="ADMIN_PRO", font=("Segoe UI", 10, "bold"), text_color=text).pack(anchor="w")
        
        # Coluna 2
        f2 = ctk.CTkFrame(info_grid, fg_color="transparent")
        f2.pack(side="left", padx=20)
        ctk.CTkLabel(f2, text="🕒 DURATION", font=("Segoe UI", 8, "bold"), text_color=accent).pack(anchor="w")
        self.lbl_duration = ctk.CTkLabel(f2, text="00:00:00", font=("Segoe UI", 10, "bold"), text_color=text)
        self.lbl_duration.pack(anchor="w")
        
        # Coluna 3
        f3 = ctk.CTkFrame(info_grid, fg_color="transparent")
        f3.pack(side="left", padx=20)
        ctk.CTkLabel(f3, text="📂 DATA SOURCE", font=("Segoe UI", 8, "bold"), text_color=accent).pack(anchor="w")
        ctk.CTkLabel(f3, text=f"...{PASTA_ATUAL[-20:]}", font=("Segoe UI", 10, "bold"), text_color=text).pack(anchor="w")

        # Botão de Ação (Direita)
        btn_container = ctk.CTkFrame(footer_f, fg_color="transparent")
        btn_container.pack(side="right", fill="y")
        
        # Efeito de Glow (Borda externa colorida)
        glow_frame = ctk.CTkFrame(btn_container, fg_color=accent, corner_radius=10)
        glow_frame.pack(side="right")
        
        self.btn_rodar = ctk.CTkButton(
            glow_frame, 
            text="START SESSION",
            font=("Segoe UI", 16, "bold"),
            fg_color=accent,
            hover_color=green,
            text_color="#000000",
            height=54,
            width=240,
            corner_radius=8,
            border_width=0,
            command=self._iniciar_robo
        )
        self.btn_rodar.pack()
        
        self.var_auto = tk.BooleanVar(value=False)
        self.chk_auto = ctk.CTkCheckBox(
            btn_container, 
            text="Auto-Mode", 
            variable=self.var_auto,
            fg_color=accent,
            hover_color=green,
            checkmark_color="#000000",
            text_color=text,
            font=("Segoe UI", 11, "bold"),
            command=self._toggle_auto
        )
        self.chk_auto.pack(side="right", padx=30)

        # Inicia monitoramento e carregamento de dados com delay
        self.after(200, self._initial_load)
        self._log("🏢 System initialized. Ready for operation.")

    def _initial_load(self):
        """Carregamento inicial de dados após o render da UI."""
        self.lbl_entrada.configure(text=str(self._contar_pdfs()))
        self._update_health()

    def _modern_stat_card(self, parent, title, value, color, column):
        card = ctk.CTkFrame(parent, fg_color="#1A1A1A", corner_radius=12, border_width=1, border_color="#222222")
        card.grid(row=0, column=column, sticky="nsew", padx=10)
        
        ctk.CTkLabel(card, text=title, font=("Segoe UI", 10, "bold"), text_color="#AAAAAA").pack(pady=(15, 0))
        
        val_lbl = ctk.CTkLabel(card, text=value, font=("Segoe UI", 32, "bold"), text_color=color)
        val_lbl.pack(pady=(5, 10))
        
        ctk.CTkLabel(card, text="TOTAL ITEMS", font=("Segoe UI", 8), text_color="#555555").pack(pady=(0, 15))
        
        return val_lbl

    def _log(self, msg):
        """Atualiza a área de log com cores se necessário."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = f"[{timestamp}] "
        
        self.log_area.configure(state="normal")
        self.log_area.insert("end", prefix)
        
        # Simulação simples de cores para níveis de log
        if "❌" in msg or "ERROR" in msg.upper():
            self.log_area.insert("end", f"{msg}\n", "error")
        elif "⚠️" in msg or "WARN" in msg.upper():
            self.log_area.insert("end", f"{msg}\n", "warning")
        else:
            self.log_area.insert("end", f"{msg}\n")
            
        self.log_area.see("end")
        self.log_area.configure(state="disabled")

    def _update_health(self):
        """Atualiza os indicadores de saúde do sistema."""
        # srv_ok = self._check_server_health() # Depende do mixin de utils
        # No CTK, atualizamos as cores dos componentes diretamente
        pass 
        # self.after(5000, self._update_health)

    def _toggle_auto(self):
        if self.var_auto.get():
            self._log("🤖 Auto-Mode ACTIVATED. Scanning for incoming files.")
            self._run_auto_loop()
        else:
            self._log("💤 Auto-Mode DEACTIVATED.")

    def _run_auto_loop(self):
        if not self.var_auto.get(): return
        if not self.rodando and self._contar_pdfs() > 0:
            self._iniciar_robo()
        self.after(30000, self._run_auto_loop)

    def _iniciar_robo(self):
        if self.rodando: return
        pdfs = self._contar_pdfs()
        if pdfs == 0:
            if not self.var_auto.get():
                self._log("⚠️ Input queue is empty!")
            return
        
        self.rodando = True
        self.btn_rodar.configure(text="⏳ PROCESSING...", state="disabled")
        threading.Thread(target=self._rodar_em_thread, daemon=True).start()

    def _rodar_em_thread(self):
        try:
            def log_fn(msg):
                self._log(msg)

            total = organizar_arquivos(log_fn, self._atualizar_stats)
            self.after(0, self._finalizar, total)
        except Exception as e:
            self._log(f"❌ Critical Error: {e}")
            self.after(0, self._finalizar, 0)

    def _finalizar(self, total):
        self.rodando = False
        self.btn_rodar.configure(text="▶  START SESSION", state="normal")
        self.lbl_entrada.configure(text=str(self._contar_pdfs()))

    def _atualizar_stats(self, proc, dupl):
        # Esta função é chamada pelo file_manager
        self.after(0, lambda: self.lbl_processados.configure(text=str(proc)))
        self.after(0, lambda: self.lbl_duplicados.configure(text=str(dupl)))

    def _contar_pdfs(self):
        try:
            return len([f for f in os.listdir(PASTA_ENTRADA) if f.lower().endswith('.pdf')])
        except:
            return 0
    def _abrir_saida(self):
        os.makedirs(PASTA_SAIDA, exist_ok=True)
        os.startfile(PASTA_ATUAL)
