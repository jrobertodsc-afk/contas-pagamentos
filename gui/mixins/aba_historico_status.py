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


class AbaHistStatusMixin:

    # ==========================================================================
    # ── ABA 4: HISTÓRICO DE STATUS ─────────────────────────────────────────
    # ==========================================================================

    def _build_aba_historico_status(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        """Aba que lista todos os logs de execução anteriores do robô."""

        # ── TÍTULO ──
        hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.pack(fill="x", padx=25, pady=(20, 10))
        
        ctk.CTkLabel(hdr, text="📜 HISTÓRICO DE EXECUÇÕES", font=("Segoe UI", 16, "bold"), text_color=accent).pack(side="left")
        ctk.CTkLabel(parent, text="Todos os logs de processamento salvos automaticamente pelo robô.", font=("Segoe UI", 11), text_color=muted).pack(anchor="w", padx=25, pady=(0, 15))

        # ── FRAME PRINCIPAL ──
        main = ctk.CTkFrame(parent, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=25, pady=(0, 10))

        # ── LISTA DE LOGS (esquerda) ──
        frame_lista = ctk.CTkFrame(main, fg_color=surface, corner_radius=12, border_width=1, border_color=border, width=320)
        frame_lista.pack(side="left", fill="y", padx=(0, 15))
        frame_lista.pack_propagate(False)

        ctk.CTkLabel(frame_lista, text="SESSÕES SALVAS", font=("Segoe UI", 10, "bold"), text_color=accent).pack(pady=(15, 10))

        self._hist_log_listbox = tk.Listbox(
            frame_lista, font=("Consolas", 10),
            bg=surface, fg=text, selectbackground=accent,
            selectforeground="#000000", relief="flat", bd=0,
            activestyle="none", highlightthickness=0
        )
        self._hist_log_listbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self._hist_log_listbox.bind("<<ListboxSelect>>", self._carregar_log_selecionado)

        # ── DETALHE DO LOG (direita) ──
        frame_detalhe = ctk.CTkFrame(main, fg_color="transparent")
        frame_detalhe.pack(side="left", fill="both", expand=True)

        barra_det = ctk.CTkFrame(frame_detalhe, fg_color=surface, corner_radius=12, border_width=1, border_color=border)
        barra_det.pack(fill="x", pady=(0, 10))

        self._hist_lbl_resumo = ctk.CTkLabel(barra_det, text="SELECIONE UM LOG PARA VISUALIZAR", font=("Segoe UI", 10, "bold"), text_color=muted)
        self._hist_lbl_resumo.pack(side="left", padx=20, pady=10)

        ctk.CTkButton(
            barra_det, text="🗑 APAGAR LOG", font=("Segoe UI", 10, "bold"), 
            fg_color="#7f1d1d", text_color="#fca5a5", hover_color="#991b1b",
            width=120, height=32, command=self._apagar_log_selecionado
        ).pack(side="right", padx=15)

        log_container = ctk.CTkFrame(frame_detalhe, fg_color=bg, corner_radius=12, border_width=1, border_color=border)
        log_container.pack(fill="both", expand=True)

        self._hist_log_area = ctk.CTkTextbox(
            log_container, font=("Consolas", 10), fg_color=bg, text_color="#ffffff", 
            border_width=0, corner_radius=10, activate_scrollbars=True
        )
        self._hist_log_area.pack(fill="both", expand=True, padx=5, pady=5)

        # ── BARRA INFERIOR ──
        barra_inf = ctk.CTkFrame(parent, fg_color="transparent")
        barra_inf.pack(fill="x", padx=25, pady=10)

        ctk.CTkButton(
            barra_inf, text="🔄 ATUALIZAR LISTA", font=("Segoe UI", 11, "bold"), 
            fg_color=border, text_color=accent, width=180, height=40, command=self._carregar_lista_logs
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            barra_inf, text="🗑 LIMPAR TUDO", font=("Segoe UI", 11, "bold"), 
            fg_color="#7f1d1d", text_color="#fca5a5", width=180, height=40, command=self._apagar_todos_logs
        ).pack(side="left")

        self._hist_lbl_total = ctk.CTkLabel(barra_inf, text="", font=("Segoe UI", 11, "bold"), text_color=muted)
        self._hist_lbl_total.pack(side="right")

        self._carregar_lista_logs()

    def _carregar_lista_logs(self):
        """Varre a pasta atual buscando arquivos log_*.txt e preenche a listbox."""
        self._hist_log_listbox.delete(0, "end")
        self._hist_logs_paths = []

        try:
            arquivos = sorted(
                [f for f in os.listdir(PASTA_ATUAL) if f.startswith("log_") and f.endswith(".txt")],
                reverse=True
            )
        except Exception:
            arquivos = []

        for arq in arquivos:
            caminho = os.path.join(PASTA_ATUAL, arq)
            try:
                partes = arq.replace("log_", "").replace(".txt", "").split("_")
                d = partes[0]; h = partes[1]
                label = f"{d[6:8]}/{d[4:6]}/{d[0:4]} {h[0:2]}:{h[2:4]}"
            except Exception:
                label = arq

            self._hist_log_listbox.insert("end", f" 📄 {label}")
            self._hist_logs_paths.append(caminho)

        total = len(arquivos)
        self._hist_lbl_total.configure(text=f"{total} SESSÕES ENCONTRADAS")

        if not arquivos:
            self._hist_log_listbox.insert("end", " NENHUM LOG ENCONTRADO")

    def _carregar_log_selecionado(self, event=None):
        """Exibe o conteúdo do log selecionado na listbox."""
        sel = self._hist_log_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self._hist_logs_paths):
            return

        caminho = self._hist_logs_paths[idx]
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                conteudo = f.read()

            ok    = conteudo.count("✅")
            erros = conteudo.count("❌")
            dup   = conteudo.count("♻️")
            resumo_txt = f"✅ {ok} OK   ❌ {erros} ERROS   ♻️ {dup} DUP   —   {os.path.basename(caminho)}"
            self._hist_lbl_resumo.configure(text=resumo_txt, text_color=accent)

            self._hist_log_area.delete("1.0", "end")
            self._hist_log_area.insert("end", conteudo)
            self._hist_log_area.see("end")
        except Exception as e:
            self._hist_lbl_resumo.configure(text=f"ERRO AO ABRIR: {e}", text_color="#f87171")

    def _apagar_log_selecionado(self):
        """Apaga o log atualmente selecionado."""
        sel = self._hist_log_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(getattr(self, "_hist_logs_paths", [])):
            return
        caminho = self._hist_logs_paths[idx]
        try:
            os.remove(caminho)
            self._hist_log_area.delete("1.0", "end")
            self._hist_lbl_resumo.configure(text="LOG APAGADO.", text_color=yellow)
        except Exception as e:
            self._hist_lbl_resumo.configure(text=f"ERRO AO APAGAR: {e}", text_color="#f87171")
        self._carregar_lista_logs()

    def _apagar_todos_logs(self):
        """Apaga todos os arquivos log_*.txt da pasta atual."""
        apagados = 0
        for caminho in getattr(self, "_hist_logs_paths", []):
            try:
                os.remove(caminho)
                apagados += 1
            except Exception:
                pass
        self._hist_log_area.delete("1.0", "end")
        self._hist_lbl_resumo.configure(text=f"{apagados} LOG(S) APAGADO(S).", text_color=yellow)
        self._carregar_lista_logs()


    def _carregar_lista_logs(self):
        """Varre a pasta atual buscando arquivos log_*.txt e preenche a listbox."""
        self._hist_log_listbox.delete(0, "end")
        self._hist_logs_paths = []

        try:
            arquivos = sorted(
                [f for f in os.listdir(PASTA_ATUAL) if f.startswith("log_") and f.endswith(".txt")],
                reverse=True
            )
        except Exception:
            arquivos = []

        for arq in arquivos:
            caminho = os.path.join(PASTA_ATUAL, arq)
            # Extrai data do nome: log_YYYYMMDD_HHMMSS.txt
            try:
                partes = arq.replace("log_", "").replace(".txt", "").split("_")
                d = partes[0]; h = partes[1]
                label = f"{d[6:8]}/{d[4:6]}/{d[0:4]}  {h[0:2]}:{h[2:4]}:{h[4:6]}"
            except Exception:
                label = arq

            # Tenta extrair resumo (última linha do arquivo)
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    linhas = f.readlines()
                resumo = ""
                for linha in reversed(linhas):
                    if linha.strip():
                        resumo = linha.strip()[-45:]
                        break
                label += f"  │  {resumo}"
            except Exception:
                pass

            self._hist_log_listbox.insert("end", label)
            self._hist_logs_paths.append(caminho)

        total = len(arquivos)
        self._hist_lbl_total.config(text=f"{total} log(s) encontrado(s)")

        if not arquivos:
            self._hist_log_listbox.insert("end", "  Nenhum log encontrado.")


    def _carregar_log_selecionado(self, event=None):
        """Exibe o conteúdo do log selecionado na listbox."""
        sel = self._hist_log_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self._hist_logs_paths):
            return

        caminho = self._hist_logs_paths[idx]
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                conteudo = f.read()

            # Conta processados e erros
            ok    = conteudo.count("✅")
            erros = conteudo.count("❌")
            dup   = conteudo.count("♻️")
            resumo_txt = f"✅ {ok} OK   ❌ {erros} Erro(s)   ♻️ {dup} Dup.   —   {os.path.basename(caminho)}"
            self._hist_lbl_resumo.config(text=resumo_txt, fg="#e2e8f0")

            self._hist_log_area.config(state="normal")
            self._hist_log_area.delete("1.0", "end")
            self._hist_log_area.insert("end", conteudo)
            self._hist_log_area.see("end")
            self._hist_log_area.config(state="disabled")
        except Exception as e:
            self._hist_lbl_resumo.config(text=f"Erro ao abrir: {e}", fg="#f87171")


    def _apagar_log_selecionado(self):
        """Apaga o log atualmente selecionado."""
        sel = self._hist_log_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(getattr(self, "_hist_logs_paths", [])):
            return
        caminho = self._hist_logs_paths[idx]
        try:
            os.remove(caminho)
            self._hist_log_area.config(state="normal")
            self._hist_log_area.delete("1.0", "end")
            self._hist_log_area.config(state="disabled")
            self._hist_lbl_resumo.config(text="Log apagado.", fg="#fbbf24")
        except Exception as e:
            self._hist_lbl_resumo.config(text=f"Erro ao apagar: {e}", fg="#f87171")
        self._carregar_lista_logs()


    def _apagar_todos_logs(self):
        """Apaga todos os arquivos log_*.txt da pasta atual."""
        apagados = 0
        for caminho in getattr(self, "_hist_logs_paths", []):
            try:
                os.remove(caminho)
                apagados += 1
            except Exception:
                pass
        self._hist_log_area.config(state="normal")
        self._hist_log_area.delete("1.0", "end")
        self._hist_log_area.config(state="disabled")
        self._hist_lbl_resumo.config(text=f"{apagados} log(s) apagado(s).", fg="#fbbf24")
        self._carregar_lista_logs()
