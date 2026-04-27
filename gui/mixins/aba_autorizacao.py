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


class AbaAutorizacaoMixin:

    # ==========================================================================
    # ── ABA 7: AUTORIZAÇÃO DE PAGAMENTOS ──────────────────────────────────
    # ==========================================================================

    def _build_aba_autorizacao(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        self._aut_pagamentos = []   # lista de dicts com todos os pagamentos carregados
        self._aut_itens_tree = []   # espelho dos itens exibidos na treeview

        # ── TÍTULO ──
        hdr = tk.Frame(parent, bg=bg)
        hdr.pack(fill="x", padx=20, pady=(12, 4))
        tk.Label(hdr, text="📋  Relatório de Autorização de Pagamentos",
                 font=("Segoe UI", 13, "bold"), fg=accent, bg=bg).pack(side="left")
        self._aut_lbl_periodo = tk.Label(hdr, text="", font=("Segoe UI", 9),
                                          fg=yellow, bg=bg)
        self._aut_lbl_periodo.pack(side="left", padx=(16, 0))

        # ── SELEÇÃO DE ARQUIVOS ──
        frame_arqs = tk.Frame(parent, bg=surface, padx=14, pady=10)
        frame_arqs.pack(fill="x", padx=20, pady=(0, 8))

        # LALUA — Fornecedores/Impostos
        tk.Label(frame_arqs, text="📄 LALUA — Fornecedores/Impostos:",
                 fg=muted, bg=surface, font=("Segoe UI", 8)).grid(row=0, column=0, sticky="w", pady=2)
        self._aut_path_forn = tk.StringVar()
        tk.Entry(frame_arqs, textvariable=self._aut_path_forn, font=("Consolas", 8),
                 bg="#0a0f1e", fg=text, insertbackground=accent,
                 relief="flat", bd=3, width=58).grid(row=0, column=1, padx=(6, 6), sticky="w")
        tk.Button(frame_arqs, text="📂", font=("Segoe UI", 9), bg=surface, fg=accent,
                  relief="flat", bd=0, cursor="hand2",
                  command=lambda: self._aut_selecionar_arquivo(self._aut_path_forn)).grid(row=0, column=2)

        # SOLAR — Fornecedores/Impostos
        tk.Label(frame_arqs, text="📄 SOLAR — Fornecedores/Impostos:",
                 fg="#fbbf24", bg=surface, font=("Segoe UI", 8)).grid(row=1, column=0, sticky="w", pady=2)
        self._aut_path_solar = tk.StringVar()
        tk.Entry(frame_arqs, textvariable=self._aut_path_solar, font=("Consolas", 8),
                 bg="#0a0f1e", fg=text, insertbackground=accent,
                 relief="flat", bd=3, width=58).grid(row=1, column=1, padx=(6, 6), sticky="w")
        tk.Button(frame_arqs, text="📂", font=("Segoe UI", 9), bg=surface, fg="#fbbf24",
                  relief="flat", bd=0, cursor="hand2",
                  command=lambda: self._aut_selecionar_arquivo(self._aut_path_solar)).grid(row=1, column=2)
        tk.Label(frame_arqs, text="(opcional)", fg="#475569", bg=surface,
                 font=("Consolas", 7)).grid(row=1, column=3, padx=4)

        # Folha de Salários
        tk.Label(frame_arqs, text="📄 Folha de Salários:",
                 fg=muted, bg=surface, font=("Segoe UI", 8)).grid(row=2, column=0, sticky="w", pady=2)
        self._aut_path_folha = tk.StringVar()
        tk.Entry(frame_arqs, textvariable=self._aut_path_folha, font=("Consolas", 8),
                 bg="#0a0f1e", fg=text, insertbackground=accent,
                 relief="flat", bd=3, width=58).grid(row=2, column=1, padx=(6, 6), sticky="w")
        tk.Button(frame_arqs, text="📂", font=("Segoe UI", 9), bg=surface, fg=accent,
                  relief="flat", bd=0, cursor="hand2",
                  command=lambda: self._aut_selecionar_arquivo(self._aut_path_folha)).grid(row=2, column=2)
        tk.Label(frame_arqs, text="(opcional)", fg="#475569", bg=surface,
                 font=("Consolas", 7)).grid(row=2, column=3, padx=4)

        # ── FILTRO DE PERÍODO ──
        frame_per = tk.Frame(parent, bg=surface, padx=14, pady=8)
        frame_per.pack(fill="x", padx=20, pady=(0, 8))

        tk.Label(frame_per, text="Período:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).pack(side="left")
        self._aut_de = tk.Entry(frame_per, font=("Segoe UI", 9), bg="#0a0f1e", fg=text,
                                 insertbackground=accent, relief="flat", bd=3, width=12)
        self._aut_de.pack(side="left", padx=(6, 4))
        tk.Label(frame_per, text="até", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).pack(side="left", padx=4)
        self._aut_ate = tk.Entry(frame_per, font=("Segoe UI", 9), bg="#0a0f1e", fg=text,
                                  insertbackground=accent, relief="flat", bd=3, width=12)
        self._aut_ate.pack(side="left", padx=(0, 12))

        tk.Button(frame_per, text="📅 Sugerir período automático",
                  font=("Segoe UI", 8), bg="#1e293b", fg=accent,
                  relief="flat", bd=0, padx=8, pady=4, cursor="hand2",
                  command=self._aut_sugerir_periodo).pack(side="left", padx=(0, 12))

        tk.Button(frame_per, text="⚡ Carregar & Classificar",
                  font=("Segoe UI", 10, "bold"), bg=accent, fg="#0f172a",
                  relief="flat", bd=0, padx=14, pady=5, cursor="hand2",
                  command=self._aut_carregar).pack(side="left", padx=(0, 8))

        self._aut_lbl_status = tk.Label(frame_per, text="", font=("Segoe UI", 9),
                                         fg=muted, bg=surface)
        self._aut_lbl_status.pack(side="left")

        # ── LABEL INSTRUÇÃO ──
        tk.Label(parent, text="✏️  Revise e corrija antes de gerar — clique para editar:",
                 font=("Segoe UI", 8, "bold"), fg=text, bg=bg).pack(anchor="w", padx=20, pady=(2, 1))

        # ── RODAPÉ FIXO (pack ANTES da treeview para garantir espaço) ──

        # Botões gerar/reclassificar
        frame_bot = tk.Frame(parent, bg=bg, pady=4)
        frame_bot.pack(fill="x", padx=20, side="bottom")

        tk.Button(frame_bot, text="💸 Transferências Filiais",
                  font=("Segoe UI", 8), bg="#1e2a3a", fg="#38bdf8",
                  relief="flat", bd=0, padx=10, pady=7, cursor="hand2",
                  command=self._gerar_pdf_transferencias_filiais
                  ).pack(side="right", padx=(0, 6))

        tk.Button(frame_bot, text="📄  GERAR PDF",
                  font=("Segoe UI", 10, "bold"), bg="#7c3aed", fg="white",
                  relief="flat", bd=0, padx=18, pady=7, cursor="hand2",
                  command=self._aut_gerar_pdf).pack(side="right")

        tk.Button(frame_bot, text="🏢 PDF Transferências Filial→Matriz",
                  font=("Segoe UI", 9, "bold"), bg="#9d174d", fg="white",
                  relief="flat", bd=0, padx=12, pady=7, cursor="hand2",
                  command=self._gerar_pdf_transferencia_filial).pack(side="right", padx=(0,6))

        tk.Button(frame_bot, text="🏦 PDF Transferências Filiais",
                  font=("Segoe UI", 9, "bold"), bg="#0c4a6e", fg="#38bdf8",
                  relief="flat", bd=0, padx=14, pady=7, cursor="hand2",
                  command=self._gerar_pdf_transferencias).pack(side="right", padx=(0,6))

        tk.Button(frame_bot, text="🔄 Reclassificar",
                  font=("Segoe UI", 8), bg=surface, fg=muted,
                  relief="flat", bd=0, padx=10, pady=7, cursor="hand2",
                  command=self._aut_popular_treeview).pack(side="right", padx=(0, 6))

        tk.Button(frame_bot, text="📊 Relatório Excel",
                  font=("Segoe UI", 9, "bold"), bg="#10b981", fg="white",
                  relief="flat", bd=0, padx=14, pady=7, cursor="hand2",
                  command=self._gerar_relatorio_pagamentos_excel).pack(side="right", padx=(0,6))

        # Resumo será exibido no rodapé da tabela

        # Painel de edição rápida
        frame_ed = tk.Frame(parent, bg=surface, padx=12, pady=5)
        frame_ed.pack(fill="x", padx=20, side="bottom")

        tk.Label(frame_ed, text="✏️ Editar selecionado(s):",
                 fg=accent, bg=surface, font=("Segoe UI", 8, "bold")).pack(side="left", padx=(0, 8))

        tk.Label(frame_ed, text="Responsável:", fg=muted, bg=surface,
                 font=("Segoe UI", 8)).pack(side="left", padx=(0, 3))
        self._aut_edit_resp = ttk.Combobox(frame_ed, values=RESPONSAVEIS_LISTA,
                                            state="readonly", width=15, font=("Segoe UI", 8))
        self._aut_edit_resp.pack(side="left", padx=(0, 10))

        tk.Label(frame_ed, text="Categoria:", fg=muted, bg=surface,
                 font=("Segoe UI", 8)).pack(side="left", padx=(0, 3))

        # Frame para categoria + busca rápida
        frame_cat = tk.Frame(frame_ed, bg=surface)
        frame_cat.pack(side="left", padx=(0, 10))

        self._aut_busca_cat = tk.StringVar()
        entry_busca = tk.Entry(frame_cat, textvariable=self._aut_busca_cat,
                               font=("Segoe UI", 8), bg="#0a0f1e", fg=text,
                               insertbackground=accent, relief="flat", bd=2,
                               width=28)
        entry_busca.pack(fill="x")
        tk.Label(frame_cat, text="🔍 digite para filtrar",
                 fg="#475569", bg=surface, font=("Consolas", 6)).pack(anchor="w")

        self._aut_edit_cat = ttk.Combobox(frame_cat, values=CATEGORIAS_LISTA,
                                           state="readonly", width=28, font=("Segoe UI", 8))
        self._aut_edit_cat.pack(fill="x")

        def _filtrar_cats(*args):
            termo = self._aut_busca_cat.get().upper()
            if not termo:
                self._aut_edit_cat["values"] = CATEGORIAS_LISTA
            else:
                filtradas = [c for c in CATEGORIAS_LISTA if termo in c.upper()]
                self._aut_edit_cat["values"] = filtradas
                if filtradas:
                    self._aut_edit_cat.set(filtradas[0])

        self._aut_busca_cat.trace_add("write", _filtrar_cats)

        tk.Label(frame_ed, text="Empresa:", fg=muted, bg=surface,
                 font=("Segoe UI", 8)).pack(side="left", padx=(0, 3))
        self._aut_edit_emp = ttk.Combobox(frame_ed, values=["LALUA", "SOLAR"],
                                           state="readonly", width=7, font=("Segoe UI", 8))
        self._aut_edit_emp.pack(side="left", padx=(0, 8))

        tk.Button(frame_ed, text="✅ Aplicar",
                  font=("Segoe UI", 8, "bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=10, pady=3, cursor="hand2",
                  command=self._aut_aplicar_edicao).pack(side="left", padx=(0, 6))

        tk.Button(frame_ed, text="🗑 Remover",
                  font=("Segoe UI", 8), bg="#7f1d1d", fg="#fca5a5",
                  relief="flat", bd=0, padx=8, pady=3, cursor="hand2",
                  command=self._aut_remover_linha).pack(side="left")

        # ── RODAPÉ COM TOTAL (FIXO NO FUNDO) ──
        frame_resumo_fixo = tk.Frame(parent, bg="#000000", padx=20, pady=10)
        frame_resumo_fixo.pack(side="bottom", fill="x")
        
        self._aut_lbl_resumo = tk.Label(frame_resumo_fixo, text="💰 TOTAL: R$ 0,00 | 📦 0 itens",
                                         font=("Segoe UI", 12, "bold"), fg=accent, bg="#111111",
                                         padx=15, pady=8, highlightbackground=accent, highlightthickness=1)
        self._aut_lbl_resumo.pack(side="left")

        # ── TREEVIEW ──
        cols = ("Favorecido", "CNPJ", "Tipo", "Data", "Valor", "Responsável", "Categoria", "Empresa")
        self._aut_tree = ttk.Treeview(parent, columns=cols, show="headings", height=20)
        larguras = [280, 140, 120, 85, 100, 140, 200, 80]
        for col, w in zip(cols, larguras):
            self._aut_tree.heading(col, text=col, command=lambda c=col: self._aut_ordenar(c))
            self._aut_tree.column(col, width=w, anchor="w")
        self._aut_tree.column("Valor", anchor="e")

        self._aut_tree.tag_configure("ok", background="#152d1f", foreground="#4ade80")
        self._aut_tree.tag_configure("pendente", background="#2d2d15", foreground="#fbbf24")
        self._aut_tree.tag_configure("a_classificar", background="#2d1515", foreground="#f87171")
        self._aut_tree.tag_configure("manual", background="#1e1b4b", foreground="#818cf8")

        scroll_aut = ttk.Scrollbar(parent, orient="vertical", command=self._aut_tree.yview)
        self._aut_tree.configure(yscrollcommand=scroll_aut.set)

        self._aut_tree.pack(side="top", fill="both", expand=True, padx=(20, 20), pady=(0, 0))
        scroll_aut.place(in_=self._aut_tree, relx=1.0, rely=0, relheight=1.0, anchor="ne")

        self._aut_tree.bind("<Double-1>",        self._aut_editar_linha)
        self._aut_tree.bind("<<TreeviewSelect>>", self._aut_editar_linha)

        # Sugere período ao iniciar
        self._aut_sugerir_periodo()


    def _aut_atualizar_resumo(self):
        """Atualiza o label de total no rodapé."""
        total = sum(p.get("valor", 0) for p in self._aut_itens_tree)
        count = len(self._aut_itens_tree)
        self._aut_lbl_resumo.config(
            text=f"💰 TOTAL DO PERÍODO: R$ {total:,.2f} | 📦 {count} itens",
            fg="#4ade80" if total > 0 else "#94a3b8"
        )


    def _aut_remover_linha(self):
        """Remove linha(s) selecionada(s) da lista."""
        sels = self._aut_tree.selection()
        if not sels:
            return
        indices = sorted([self._aut_tree.index(s) for s in sels], reverse=True)
        for idx in indices:
            if 0 <= idx < len(self._aut_itens_tree):
                self._aut_itens_tree.pop(idx)
        self._aut_pagamentos = list(self._aut_itens_tree)
        self._aut_popular_treeview()


    def _aut_ordenar(self, coluna):
        """Ordena a treeview ao clicar no cabeçalho."""
        mapa = {
            "Favorecido": "nome", "CNPJ": "cnpj", "Tipo": "tipo",
            "Data": "data", "Valor": "valor", "Responsável": "responsavel",
            "Categoria": "categoria", "Empresa": "empresa"
        }
        chave = mapa.get(coluna, "nome")
        if not hasattr(self, "_aut_sort_col"):
            self._aut_sort_col = None
            self._aut_sort_rev = False
        if self._aut_sort_col == chave:
            self._aut_sort_rev = not self._aut_sort_rev
        else:
            self._aut_sort_col = chave
            self._aut_sort_rev = False
        self._aut_itens_tree.sort(
            key=lambda x: str(x.get(chave, "")),
            reverse=self._aut_sort_rev
        )
        self._aut_pagamentos = list(self._aut_itens_tree)
        self._aut_popular_treeview()


    def _aut_selecionar_arquivo(self, var):
        from tkinter import filedialog
        caminho = filedialog.askopenfilename(
            filetypes=[("Excel/XLS", "*.xls *.xlsx"), ("Todos", "*.*")]
        )
        if caminho:
            var.set(caminho)


    def _aut_sugerir_periodo(self):
        ini, fim, label = calcular_periodo_sugerido()
        self._aut_de.delete(0, "end")
        self._aut_de.insert(0, ini.strftime("%d/%m/%Y"))
        self._aut_ate.delete(0, "end")
        self._aut_ate.insert(0, fim.strftime("%d/%m/%Y"))
        self._aut_lbl_periodo.config(text=f"📅 Sugerido: {label}")


    def _aut_carregar(self):
        path_forn  = self._aut_path_forn.get().strip()
        path_solar = self._aut_path_solar.get().strip()
        path_folha = self._aut_path_folha.get().strip()

        if not path_forn and not path_solar:
            self._aut_lbl_status.config(text="⚠️ Selecione ao menos um arquivo.", fg="#fbbf24")
            return

        try:
            de_str  = self._aut_de.get().strip()
            ate_str = self._aut_ate.get().strip()
            dt_de   = datetime.strptime(de_str,  "%d/%m/%Y") if de_str  else datetime(2000,1,1)
            dt_ate  = datetime.strptime(ate_str, "%d/%m/%Y") if ate_str else datetime(2099,1,1)
        except:
            self._aut_lbl_status.config(text="⚠️ Data inválida.", fg="#fbbf24")
            return

        self._aut_lbl_status.config(text="⏳ Carregando...", fg="#38bdf8")
        self.root.update_idletasks()

        def carregar():
            try:
                pagamentos = []

                # LALUA
                if path_forn and os.path.exists(path_forn):
                    itens = ler_itau_pagamentos(path_forn)
                    for p in itens: p["empresa"] = "LALUA"
                    pagamentos += itens

                # SOLAR
                if path_solar and os.path.exists(path_solar):
                    itens_solar = ler_itau_pagamentos(path_solar)
                    for p in itens_solar: p["empresa"] = "SOLAR"
                    pagamentos += itens_solar

                # Folha
                if path_folha and os.path.exists(path_folha):
                    pagamentos += ler_itau_folha(path_folha)

                # Filtra
                filtrados = []
                for p in pagamentos:
                    if p.get("data_obj"):
                        if dt_de <= p["data_obj"] <= dt_ate:
                            filtrados.append(p)
                    else:
                        filtrados.append(p)

                self._aut_pagamentos = filtrados
                pre_count = len(pagamentos)
                
                def finalizar():
                    self._aut_lbl_status.config(
                        text=f"✅ {pre_count} total | {len(filtrados)} no período", fg="#4ade80")
                    self._aut_popular_treeview()
                
                self.root.after(0, finalizar)

            except Exception as e:
                self.root.after(0, lambda: self._aut_lbl_status.config(
                    text=f"❌ Erro: {str(e)}", fg="#f87171"))

        threading.Thread(target=carregar, daemon=True).start()


    def _aut_popular_treeview(self):
        for row in self._aut_tree.get_children():
            self._aut_tree.delete(row)

        self._aut_itens_tree = list(self._aut_pagamentos)
        sem_cat = 0
        total   = 0.0

        for p in self._aut_itens_tree:
            cat = p.get("categoria", "")
            resp = p.get("responsavel", "")
            tag = "ok" if (cat and cat != "A CLASSIFICAR" and resp) else \
                  "a_classificar" if cat == "A CLASSIFICAR" else "pendente"
            if tag != "ok":
                sem_cat += 1
            total += p["valor"]

            self._aut_tree.insert("", "end", tags=(tag,), values=(
                p["nome"][:35],
                p["cnpj"],
                p["tipo"],
                p["data"],
                f"R$ {p['valor']:,.2f}",
                resp,
                cat[:40],
                p.get("empresa", "LALUA"),
            ))

        ok_count = len(self._aut_itens_tree) - sem_cat
        self._aut_lbl_resumo.config(
            text=f"💰 TOTAL: R$ {total:,.2f}   |   📦 {len(self._aut_itens_tree)} itens   |   ⚠️ {sem_cat} pendentes",
            fg="#a3e635" if sem_cat == 0 else "#fbbf24"
        )


    def _aut_editar_linha(self, event=None):
        """Preenche o painel de edição com os valores da linha clicada."""
        sel = self._aut_tree.selection()
        if not sel:
            return
        idx = self._aut_tree.index(sel[0])
        p   = self._aut_itens_tree[idx]
        self._aut_edit_resp.set(p.get("responsavel", ""))
        self._aut_edit_cat.set(p.get("categoria", ""))
        self._aut_edit_emp.set(p.get("empresa", "LALUA"))


    def _aut_aplicar_edicao(self):
        """Aplica a edição do painel ao(s) item(ns) selecionado(s)."""
        sels = self._aut_tree.selection()
        if not sels:
            return
        resp = self._aut_edit_resp.get().strip()
        cat  = self._aut_edit_cat.get().strip()
        emp  = self._aut_edit_emp.get().strip()

        for sel in sels:
            idx = self._aut_tree.index(sel)
            if resp: self._aut_itens_tree[idx]["responsavel"] = resp
            if cat:  self._aut_itens_tree[idx]["categoria"]   = cat
            if emp:  self._aut_itens_tree[idx]["empresa"]      = emp

        # Sincroniza com _aut_pagamentos
        self._aut_pagamentos = list(self._aut_itens_tree)
        self._aut_popular_treeview()


    def _gerar_pdf_transferencias(self):
        """PDF de transferências de filiais para a matriz — padrão BOAH."""
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib import colors
            from reportlab.lib.units import cm, mm
            from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                            Paragraph, Spacer, HRFlowable,
                                            KeepTogether, PageBreak)
            from reportlab.lib.styles import ParagraphStyle
            from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
        except ImportError:
            self._aut_lbl_resumo.config(
                text="❌ pip install reportlab", fg="#f87171")
            return

        # Coleta transferências filial→matriz dos pagamentos carregados
        transferencias = [p for p in self._aut_pagamentos
                          if "FILIAL" in p.get("tipo","").upper()
                          or "TRANSFERENCIA MATRIZ" in p.get("categoria","").upper()
                          or "FILIAL" in p.get("categoria","").upper()]

        # Se não houver, abre formulário para informar manualmente
        if not transferencias:
            self._aut_lbl_resumo.config(
                text="ℹ️ Nenhuma transferência filial→matriz. "
                     "Use '➕ Inserção Manual' com tipo 'PIX FILIAL→MATRIZ'.",
                fg="#fbbf24")
            return

        # ── PALETA BOAH ──
        PRETO     = colors.HexColor("#1A1A1A")
        CREME     = colors.HexColor("#F5F0E8")
        VINHO     = colors.HexColor("#8B1A2B")
        VINHO_CLR = colors.HexColor("#C41E3A")
        BEGE      = colors.HexColor("#D4C5A9")
        CINZA_CLR = colors.HexColor("#F0ECE4")
        CINZA_HDR = colors.HexColor("#3A3A3A")
        AZUL_FILIAL = colors.HexColor("#0c4a6e")
        AZUL_CLR  = colors.HexColor("#0ea5e9")
        BRANCO    = colors.white

        PAGE = landscape(A4)
        os.makedirs(PASTA_ATUAL, exist_ok=True)
        data_hoje = datetime.now().strftime("%d_%m_%Y_%H%M")
        caminho_pdf = os.path.join(PASTA_ATUAL,
                                   f"TRANSFERENCIAS_FILIAIS_{data_hoje}.pdf")

        doc = SimpleDocTemplate(
            caminho_pdf,
            pagesize=PAGE,
            leftMargin=1.8*cm, rightMargin=1.8*cm,
            topMargin=1.4*cm,  bottomMargin=1.4*cm,
            title="BOAH — Transferências Filiais → Matriz"
        )

        de_str  = self._aut_de.get().strip()
        ate_str = self._aut_ate.get().strip()

        def P(txt, fs=8, bold=False, cor=None, align=TA_LEFT):
            fn  = "Helvetica-Bold" if bold else "Helvetica"
            clr = cor or PRETO
            return Paragraph(str(txt),
                ParagraphStyle(f"p{abs(hash(str(txt)+str(fs)+str(bold))%999999)}",
                    fontSize=fs, fontName=fn, textColor=clr,
                    alignment=align, leading=fs+2))

        def _cabecalho():
            cab = [[
                Table([[
                    [P("BOAH", fs=26, bold=True)],
                    [P("LALUA COMERCIO DE MODAS LTDA  ·  SOLAR SERVIÇOS DE APOIO ADMIN",
                       fs=7.5, cor=colors.HexColor("#555"))],
                ]], colWidths=[doc.width*0.5]),
                Table([[
                    [P("TRANSFERÊNCIAS FILIAIS → MATRIZ", fs=10, bold=True,
                       cor=AZUL_FILIAL, align=TA_RIGHT)],
                    [P(f"Período: {de_str}  a  {ate_str}", fs=8,
                       cor=colors.HexColor("#555"), align=TA_RIGHT)],
                    [P(f"Emitido em {datetime.now().strftime('%d/%m/%Y às %H:%M')}",
                       fs=7, cor=colors.HexColor("#888"), align=TA_RIGHT)],
                ]], colWidths=[doc.width*0.5]),
            ]]
            t = Table(cab, colWidths=[doc.width*0.5, doc.width*0.5])
            t.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), CREME),
                ("TOPPADDING",    (0,0), (-1,-1), 10),
                ("BOTTOMPADDING", (0,0), (-1,-1), 10),
                ("LEFTPADDING",   (0,0), (0,-1),  14),
                ("RIGHTPADDING",  (1,0), (1,-1),  14),
                ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ]))
            return [t, HRFlowable(width="100%", thickness=2,
                                   color=AZUL_CLR, spaceAfter=8)]

        # Agrupa por filial (nome do beneficiário contém a filial)
        # Tenta extrair nome da filial do nome do pagamento
        FILIAIS_CONHECIDAS = [
            "BARRA", "HORTO", "VILAS", "PASEO", "SDB",
            "PARALELA", "MATRIZ", "ONLINE", "ECOMMERCE"
        ]

        from collections import defaultdict
        por_filial = defaultdict(list)
        for p in transferencias:
            nome_up = p.get("nome","").upper()
            filial = "FILIAL NÃO IDENTIFICADA"
            for f in FILIAIS_CONHECIDAS:
                if f in nome_up:
                    filial = f
                    break
            # Fallback — usa o próprio nome
            if filial == "FILIAL NÃO IDENTIFICADA":
                filial = p.get("nome","FILIAL").upper()
            por_filial[filial].append(p)

        story = []
        story.extend(_cabecalho())

        total_geral = 0.0
        cw = [doc.width*p for p in [0.22, 0.10, 0.08, 0.12, 0.12, 0.18, 0.18]]
        HDRS = ["Favorecido / Identificação", "Data", "Valor (R$)",
                "Responsável", "Empresa", "Categoria", "Observação"]

        for filial_nome in sorted(por_filial.keys()):
            itens = por_filial[filial_nome]
            subtotal = sum(i["valor"] for i in itens)
            total_geral += subtotal

            # Cabeçalho da filial
            cab_filial = [[
                P(f"  🏢  FILIAL: {filial_nome}", fs=9, bold=True,
                  cor=BRANCO, align=TA_LEFT),
                P(f"{len(itens)} transferência(s)  ·  R$ {subtotal:,.2f}",
                  fs=8.5, bold=True, cor=CREME, align=TA_RIGHT),
            ]]
            tbl_cf = Table(cab_filial,
                           colWidths=[doc.width*0.6, doc.width*0.4])
            tbl_cf.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), AZUL_FILIAL),
                ("TOPPADDING",    (0,0), (-1,-1), 7),
                ("BOTTOMPADDING", (0,0), (-1,-1), 7),
                ("LEFTPADDING",   (0,0), (0,-1),  10),
                ("RIGHTPADDING",  (1,0), (1,-1),  10),
                ("LINEBELOW",     (0,0), (-1,-1), 2.5, AZUL_CLR),
                ("LINEBEFORE",    (0,0), (0,-1),  4, AZUL_CLR),
            ]))
            story.append(tbl_cf)

            # Linhas de dados
            hdr_r = [P(h, fs=6.5, bold=True, cor=BRANCO) for h in HDRS]
            rows  = [hdr_r]
            for item in sorted(itens, key=lambda x: x.get("data","")):
                rows.append([
                    P(item["nome"][:40], fs=7.5),
                    P(item.get("data",""), fs=7, align=TA_CENTER),
                    P(f"R$ {item['valor']:,.2f}", fs=7.5, bold=True,
                      cor=AZUL_FILIAL, align=TA_RIGHT),
                    P(item.get("responsavel",""), fs=7, align=TA_CENTER),
                    P(item.get("empresa",""), fs=7, align=TA_CENTER),
                    P(item.get("categoria",""), fs=7),
                    P(item.get("observacao",""), fs=7),
                ])

            # Subtotal da filial
            nr = len(rows)
            rows.append([
                P(f"SUBTOTAL  {filial_nome}", fs=8, bold=True),
                "", "",
                P(f"R$ {subtotal:,.2f}", fs=9, bold=True,
                  cor=AZUL_FILIAL, align=TA_RIGHT),
                "", "", "",
            ])

            tbl = Table(rows, colWidths=cw, repeatRows=1)
            nrr = len(rows)
            tbl.setStyle(TableStyle([
                ("BACKGROUND",    (0,0),      (-1,0),      CINZA_HDR),
                ("TEXTCOLOR",     (0,0),      (-1,0),      BRANCO),
                ("ROWBACKGROUNDS",(0,1),      (-1,nrr-2),  [BRANCO, CINZA_CLR]),
                ("ALIGN",         (1,1),      (1,nrr-2),   "CENTER"),
                ("ALIGN",         (2,1),      (2,nrr-2),   "RIGHT"),
                ("ALIGN",         (3,1),      (4,nrr-2),   "CENTER"),
                ("BACKGROUND",    (0,nrr-1),  (-1,nrr-1),  BEGE),
                ("FONTNAME",      (0,nrr-1),  (-1,nrr-1),  "Helvetica-Bold"),
                ("SPAN",          (0,nrr-1),  (2,nrr-1)),
                ("SPAN",          (3,nrr-1),  (6,nrr-1)),
                ("ALIGN",         (3,nrr-1),  (6,nrr-1),   "RIGHT"),
                ("LINEBEFORE",    (0,0),      (0,-1),       3, AZUL_CLR),
                ("LINEBELOW",     (0,0),      (-1,0),       0.5, BEGE),
                ("LINEBELOW",     (0,1),      (-1,nrr-2),   0.3, BEGE),
                ("TOPPADDING",    (0,0),      (-1,-1),      3),
                ("BOTTOMPADDING", (0,0),      (-1,-1),      3),
                ("LEFTPADDING",   (0,0),      (-1,-1),      4),
                ("RIGHTPADDING",  (0,0),      (-1,-1),      4),
            ]))
            story.append(tbl)
            story.append(Spacer(1, 10))

        # ── TOTAL GERAL ──
        story.append(HRFlowable(width="100%", thickness=1.5,
                                 color=AZUL_CLR, spaceBefore=4, spaceAfter=4))
        tgd = [[
            P("TOTAL GERAL TRANSFERÊNCIAS", fs=11, bold=True, cor=BRANCO),
            P(f"R$ {total_geral:,.2f}", fs=12, bold=True,
              cor=CREME, align=TA_RIGHT),
        ]]
        tbl_tg = Table(tgd, colWidths=[doc.width*0.6, doc.width*0.4])
        tbl_tg.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), AZUL_FILIAL),
            ("ALIGN",         (1,0), (1,-1),  "RIGHT"),
            ("TOPPADDING",    (0,0), (-1,-1),  10),
            ("BOTTOMPADDING", (0,0), (-1,-1),  10),
            ("LEFTPADDING",   (0,0), (-1,-1),  12),
            ("RIGHTPADDING",  (1,0), (1,-1),   12),
        ]))
        story.append(tbl_tg)
        story.append(Spacer(1, 8))
        story.append(P(
            f"Documento interno · Uso exclusivo Contas a Pagar BOAH/SOLAR  ·  "
            f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            fs=6.5, cor=colors.HexColor("#999999"), align=TA_CENTER))

        doc.build(story)
        self._aut_lbl_resumo.config(
            text=f"✅ PDF gerado: {os.path.basename(caminho_pdf)}", fg="#4ade80")
        try:
            os.startfile(caminho_pdf)
        except: pass


    def _gerar_pdf_transferencia_filial(self):
        """PDF individual por filial — uma página por filial com
        tabela de conferência de fornecedores (20 linhas) e área de assinatura."""
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib import colors
            from reportlab.lib.units import cm, mm
            from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                            Paragraph, Spacer, HRFlowable, PageBreak)
            from reportlab.lib.styles import ParagraphStyle
            from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
        except ImportError:
            self._aut_lbl_resumo.config(
                text="❌ pip install reportlab", fg="#f87171")
            return

        # Coleta transferências
        transferencias = [p for p in self._aut_pagamentos
                          if "FILIAL" in p.get("tipo","").upper()
                          or "TRANSFERENCIA MATRIZ" in p.get("categoria","").upper()]

        if not transferencias:
            self._aut_lbl_resumo.config(
                text="ℹ️ Nenhuma transferência filial→matriz registrada.",
                fg="#fbbf24")
            return

        # ── PALETA ──
        PRETO     = colors.HexColor("#1A1A1A")
        CREME     = colors.HexColor("#F5F0E8")
        VINHO     = colors.HexColor("#8B1A2B")
        VINHO_CLR = colors.HexColor("#C41E3A")
        BEGE      = colors.HexColor("#D4C5A9")
        CINZA_CLR = colors.HexColor("#F0ECE4")
        CINZA_HDR = colors.HexColor("#3A3A3A")
        AZUL      = colors.HexColor("#0c4a6e")
        AZUL_CLR  = colors.HexColor("#0ea5e9")
        BRANCO    = colors.white

        PAGE = A4  # Retrato para caber na impressora de qualquer filial
        os.makedirs(PASTA_ATUAL, exist_ok=True)
        data_hoje = datetime.now().strftime("%d_%m_%Y_%H%M")
        caminho_pdf = os.path.join(PASTA_ATUAL,
                                   f"TRANSFERENCIA_FILIAL_{data_hoje}.pdf")

        doc = SimpleDocTemplate(
            caminho_pdf,
            pagesize=PAGE,
            leftMargin=1.8*cm, rightMargin=1.8*cm,
            topMargin=1.4*cm,  bottomMargin=1.4*cm,
            title="BOAH — Transferência Filial → Matriz"
        )

        de_str  = self._aut_de.get().strip()
        ate_str = self._aut_ate.get().strip()

        def P(txt, fs=8, bold=False, cor=None, align=TA_LEFT):
            fn  = "Helvetica-Bold" if bold else "Helvetica"
            clr = cor or PRETO
            return Paragraph(str(txt),
                ParagraphStyle(f"p{abs(hash(str(txt)+str(fs)+str(bold))%999999)}",
                    fontSize=fs, fontName=fn, textColor=clr,
                    alignment=align, leading=fs+2))

        # Agrupa por filial
        FILIAIS_CONHECIDAS = [
            "BARRA","HORTO","VILAS","PASEO","SDB","PARALELA",
            "MATRIZ","ONLINE","ECOMMERCE"
        ]
        from collections import defaultdict
        por_filial = defaultdict(list)
        for p in transferencias:
            nome_up = p.get("nome","").upper()
            filial  = "NÃO IDENTIFICADA"
            for f in FILIAIS_CONHECIDAS:
                if f in nome_up:
                    filial = f
                    break
            if filial == "NÃO IDENTIFICADA":
                filial = p.get("nome","FILIAL").upper()
            por_filial[filial].append(p)

        story  = []
        pagina = 0

        for filial_nome in sorted(por_filial.keys()):
            itens    = por_filial[filial_nome]
            subtotal = sum(i["valor"] for i in itens)

            if pagina > 0:
                story.append(PageBreak())
            pagina += 1

            # ── CABEÇALHO ──
            cab = [[
                Table([[
                    [P("BOAH", fs=22, bold=True)],
                    [P("LALUA COMERCIO DE MODAS LTDA  ·  SOLAR SERVIÇOS DE APOIO ADMIN",
                       fs=7)],
                ]], colWidths=[doc.width*0.5]),
                Table([[
                    [P("COMPROVANTE DE TRANSFERÊNCIA", fs=9, bold=True,
                       cor=AZUL, align=TA_RIGHT)],
                    [P(f"Filial Remetente: {filial_nome}", fs=10, bold=True,
                       cor=VINHO, align=TA_RIGHT)],
                    [P(f"Período: {de_str}  a  {ate_str} | "
                       f"{datetime.now().strftime('%d/%m/%Y %H:%M')}",
                       fs=7, cor=colors.HexColor("#888"), align=TA_RIGHT)],
                ]], colWidths=[doc.width*0.5]),
            ]]
            t_cab = Table(cab, colWidths=[doc.width*0.5, doc.width*0.5])
            t_cab.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), CREME),
                ("TOPPADDING",    (0,0), (-1,-1), 10),
                ("BOTTOMPADDING", (0,0), (-1,-1), 10),
                ("LEFTPADDING",   (0,0), (0,-1),  14),
                ("RIGHTPADDING",  (1,0), (1,-1),  14),
                ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ]))
            story.append(t_cab)
            story.append(HRFlowable(width="100%", thickness=3,
                                     color=VINHO_CLR, spaceAfter=8))

            # ── IDENTIFICAÇÃO DA FILIAL ──
            id_filial = [[
                P(f"FILIAL REMETENTE:", fs=8, bold=True, cor=AZUL),
                P(filial_nome, fs=14, bold=True, cor=VINHO),
                P("DATA DO ENVIO:", fs=8, bold=True, cor=AZUL, align=TA_RIGHT),
                P(itens[0].get("data","__/__/____"), fs=11, bold=True,
                  cor=PRETO, align=TA_RIGHT),
            ]]
            t_id = Table(id_filial,
                         colWidths=[doc.width*0.2, doc.width*0.4,
                                     doc.width*0.2, doc.width*0.2])
            t_id.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), BEGE),
                ("TOPPADDING",    (0,0), (-1,-1), 10),
                ("BOTTOMPADDING", (0,0), (-1,-1), 10),
                ("LEFTPADDING",   (0,0), (-1,-1), 10),
                ("RIGHTPADDING",  (0,0), (-1,-1), 10),
                ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
                ("BOX",           (0,0), (-1,-1), 1.5, VINHO),
            ]))
            story.append(t_id)
            story.append(Spacer(1, 8))

            # ── TABELA DE TRANSFERÊNCIAS ──
            story.append(P("DETALHAMENTO DAS TRANSFERÊNCIAS", fs=8,
                           bold=True, cor=AZUL))
            story.append(Spacer(1, 4))

            cw_t = [doc.width*p for p in [0.34, 0.12, 0.14, 0.18, 0.22]]
            HDRS_T = ["Favorecido / Descrição", "Data", "Valor (R$)",
                      "Responsável", "Categoria"]
            hdr_r = [P(h, fs=7, bold=True, cor=BRANCO) for h in HDRS_T]
            rows  = [hdr_r]
            for item in sorted(itens, key=lambda x: x.get("data","")):
                rows.append([
                    P(item["nome"][:50], fs=8),
                    P(item.get("data",""), fs=8, align=TA_CENTER),
                    P(f"R$ {item['valor']:,.2f}", fs=8, bold=True,
                      cor=AZUL, align=TA_RIGHT),
                    P(item.get("responsavel",""), fs=7.5, align=TA_CENTER),
                    P(item.get("categoria",""), fs=7.5),
                ])

            # Linha total
            nr  = len(rows)
            rows.append([
                P("TOTAL TRANSFERIDO", fs=9, bold=True),
                "",
                P(f"R$ {subtotal:,.2f}", fs=10, bold=True,
                  cor=VINHO, align=TA_RIGHT),
                "", "",
            ])
            nrr = len(rows)

            tbl_t = Table(rows, colWidths=cw_t, repeatRows=1)
            tbl_t.setStyle(TableStyle([
                ("BACKGROUND",    (0,0),     (-1,0),     CINZA_HDR),
                ("ROWBACKGROUNDS",(0,1),     (-1,nr-1),  [BRANCO, CINZA_CLR]),
                ("ALIGN",         (1,1),     (1,nr-1),   "CENTER"),
                ("ALIGN",         (2,1),     (2,nr-1),   "RIGHT"),
                ("ALIGN",         (3,1),     (3,nr-1),   "CENTER"),
                ("BACKGROUND",    (0,nrr-1), (-1,nrr-1), CREME),
                ("FONTNAME",      (0,nrr-1), (-1,nrr-1), "Helvetica-Bold"),
                ("SPAN",          (0,nrr-1), (1,nrr-1)),
                ("SPAN",          (2,nrr-1), (4,nrr-1)),
                ("ALIGN",         (2,nrr-1), (4,nrr-1),  "RIGHT"),
                ("BOX",           (0,0),     (-1,-1),     1, BEGE),
                ("LINEBELOW",     (0,0),     (-1,0),      0.5, BEGE),
                ("LINEBELOW",     (0,1),     (-1,nr-1),   0.3, BEGE),
                ("TOPPADDING",    (0,0),     (-1,-1),     4),
                ("BOTTOMPADDING", (0,0),     (-1,-1),     4),
                ("LEFTPADDING",   (0,0),     (-1,-1),     5),
                ("RIGHTPADDING",  (0,0),     (-1,-1),     5),
                ("RIGHTPADDING",  (2,nrr-1), (4,nrr-1),  8),
            ]))
            story.append(tbl_t)
            story.append(Spacer(1, 12))

            # ── CONFERÊNCIA DE FORNECEDORES (20 linhas) ──
            story.append(P("CONFERÊNCIA DE FORNECEDORES / PAGAMENTOS DA FILIAL",
                           fs=8, bold=True, cor=AZUL))
            story.append(Spacer(1, 4))

            cw_c = [doc.width*p for p in [0.05, 0.35, 0.14, 0.13, 0.13, 0.20]]
            HDRS_C = ["Nº", "Fornecedor / Beneficiário", "CNPJ/CPF",
                      "Vencimento", "Valor (R$)", "Categoria"]
            hdr_c = [P(h, fs=7, bold=True, cor=BRANCO) for h in HDRS_C]
            rows_c = [hdr_c]

            # 20 linhas em branco para preenchimento manual
            for i in range(1, 21):
                bg_linha = BRANCO if i % 2 == 0 else CINZA_CLR
                rows_c.append([
                    P(str(i), fs=7.5, cor=colors.HexColor("#999"),
                      align=TA_CENTER),
                    P("", fs=8), P("", fs=8), P("", fs=8),
                    P("", fs=8), P("", fs=8),
                ])

            # Linha total conferência
            rows_c.append([
                P("TOTAL", fs=8, bold=True), "", "",
                P("", fs=8),
                P("R$", fs=8, bold=True, align=TA_RIGHT),
                "",
            ])

            tbl_c = Table(rows_c, colWidths=cw_c, repeatRows=1)
            nr_c  = len(rows_c)
            tbl_c.setStyle(TableStyle([
                ("BACKGROUND",    (0,0),      (-1,0),       CINZA_HDR),
                ("TEXTCOLOR",     (0,0),      (-1,0),       BRANCO),
                ("ROWBACKGROUNDS",(0,1),      (-1,nr_c-2),
                                  [BRANCO, CINZA_CLR]),
                ("ALIGN",         (0,0),      (0,-1),       "CENTER"),
                ("ALIGN",         (3,1),      (4,nr_c-2),   "CENTER"),
                ("BACKGROUND",    (0,nr_c-1), (-1,nr_c-1),  BEGE),
                ("FONTNAME",      (0,nr_c-1), (-1,nr_c-1),  "Helvetica-Bold"),
                ("SPAN",          (0,nr_c-1), (3,nr_c-1)),
                ("BOX",           (0,0),      (-1,-1),       1, BEGE),
                ("LINEBELOW",     (0,0),      (-1,-1),       0.3, BEGE),
                ("TOPPADDING",    (0,0),      (-1,-1),       5),
                ("BOTTOMPADDING", (0,0),      (-1,-1),       5),
                ("LEFTPADDING",   (0,0),      (-1,-1),       5),
                ("RIGHTPADDING",  (0,0),      (-1,-1),       5),
            ]))
            story.append(tbl_c)
            story.append(Spacer(1, 14))

            # ── ÁREA DE ASSINATURAS ──
            linha_ass = "_" * 35
            ass = [[
                Table([[
                    [P(linha_ass, fs=8, cor=colors.HexColor("#999"))],
                    [P(f"Responsável — {filial_nome}", fs=7.5, bold=True,
                       cor=AZUL, align=TA_CENTER)],
                    [P("Data: ____/____/________", fs=7.5,
                       cor=colors.HexColor("#666"), align=TA_CENTER)],
                ]], colWidths=[doc.width*0.32]),
                Table([[
                    [P(linha_ass, fs=8, cor=colors.HexColor("#999"))],
                    [P("Conferido por — Contas a Pagar BOAH", fs=7.5,
                       bold=True, cor=AZUL, align=TA_CENTER)],
                    [P("Data: ____/____/________", fs=7.5,
                       cor=colors.HexColor("#666"), align=TA_CENTER)],
                ]], colWidths=[doc.width*0.32]),
                Table([[
                    [P(linha_ass, fs=8, cor=colors.HexColor("#999"))],
                    [P("Aprovado por — Financeiro BOAH", fs=7.5,
                       bold=True, cor=VINHO, align=TA_CENTER)],
                    [P("Data: ____/____/________", fs=7.5,
                       cor=colors.HexColor("#666"), align=TA_CENTER)],
                ]], colWidths=[doc.width*0.32]),
            ]]
            t_ass = Table(ass, colWidths=[doc.width*0.34]*3)
            t_ass.setStyle(TableStyle([
                ("ALIGN",         (0,0), (-1,-1), "CENTER"),
                ("VALIGN",        (0,0), (-1,-1), "BOTTOM"),
                ("TOPPADDING",    (0,0), (-1,-1), 4),
                ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ]))
            story.append(t_ass)
            story.append(Spacer(1, 6))
            story.append(HRFlowable(width="100%", thickness=0.5, color=BEGE))
            story.append(P(
                f"Documento interno · BOAH/SOLAR  ·  "
                f"Filial: {filial_nome}  ·  "
                f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                fs=6.5, cor=colors.HexColor("#999"), align=TA_CENTER))

        doc.build(story)
        n = len(por_filial)
        self._aut_lbl_resumo.config(
            text=f"✅ PDF gerado — {n} filial(is) | "
                 f"R$ {sum(i['valor'] for i in transferencias):,.2f}",
            fg="#4ade80")
        try:
            os.startfile(caminho_pdf)
        except: pass


    def _aut_gerar_pdf(self):
        """Gera o PDF de autorização de pagamentos no estilo BOAH."""
        if not self._aut_pagamentos:
            self._aut_lbl_status.config(text="⚠️ Carregue os pagamentos primeiro.", fg="#fbbf24")
            return

        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib import colors
            from reportlab.lib.units import cm
            from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                            Paragraph, Spacer, HRFlowable)
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
        except ImportError:
            self._aut_lbl_status.config(
                text="❌ reportlab não instalado. Execute: pip install reportlab", fg="#f87171")
            return

        de_str  = self._aut_de.get().strip()
        ate_str = self._aut_ate.get().strip()

        # Agrupa por empresa → data → tipo de pagamento
        from collections import defaultdict
        # estrutura: empresas[emp][data_str][tipo] = [itens]
        empresas = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        for p in self._aut_pagamentos:
            emp  = p.get("empresa", "LALUA")
            data = p.get("data", "")
            tipo = p.get("tipo", "Outros")
            empresas[emp][data][tipo].append(p)

        ORDEM_TIPO = [
            "Boleto Itaú",
            "Boleto outros bancos",
            "Boletos ouros bancos",       # variação de digitação do Itaú
            "Concessionária",
            "Tributos com código de barras",
            "DARF código de barras",      # Receita Federal
            "DAS código barras",          # Simples Nacional
            "IPTU/ISS e outros tributos",
            "PIX Transferências",
            "Pix Transferencia",          # variação
            "PIX Qr Code",
            "TED",
            "Folha de pagamento",
            "Folha de pagamentos",
            "Outros",
        ]

        os.makedirs(PASTA_ATUAL, exist_ok=True)
        data_hoje = datetime.now().strftime("%d_%m_%Y")
        nome_arq  = f"AUTORIZACAO_PAGAMENTOS_{data_hoje}.pdf"
        caminho_pdf = os.path.join(PASTA_ATUAL, nome_arq)

        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib import colors
            from reportlab.lib.units import cm, mm
            from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                            Paragraph, Spacer, HRFlowable,
                                            KeepTogether, PageBreak)
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
        except ImportError:
            self._aut_lbl_status.config(
                text="❌ reportlab não instalado. Execute: pip install reportlab", fg="#f87171")
            return

        de_str  = self._aut_de.get().strip()
        ate_str = self._aut_ate.get().strip()

        from collections import defaultdict
        empresas = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        for p in self._aut_pagamentos:
            emp  = p.get("empresa", "LALUA")
            data = p.get("data", "")
            tipo = p.get("tipo", "Outros")
            empresas[emp][data][tipo].append(p)

        ORDEM_TIPO = [
            "Boleto Itaú", "Boleto outros bancos", "Boletos ouros bancos",
            "Concessionária", "Tributos com código de barras",
            "DARF código de barras", "DAS código barras",
            "IPTU/ISS e outros tributos",
            "PIX Transferências", "Pix Transferencia",
            "PIX Qr Code", "TED",
            "Folha de pagamento", "Folha de pagamentos", "Outros",
        ]

        os.makedirs(PASTA_ATUAL, exist_ok=True)
        data_hoje = datetime.now().strftime("%d_%m_%Y")
        nome_arq  = f"AUTORIZACAO_PAGAMENTOS_{data_hoje}.pdf"
        caminho_pdf = os.path.join(PASTA_ATUAL, nome_arq)

        try:
            # ── PAISAGEM A4 ──
            PAGE = landscape(A4)
            doc = SimpleDocTemplate(
                caminho_pdf,
                pagesize=PAGE,
                leftMargin=1.8*cm, rightMargin=1.8*cm,
                topMargin=1.4*cm,  bottomMargin=1.4*cm,
                title=f"BOAH — Autorização de Pagamentos {de_str}"
            )

            # ══ PALETA BOAH ══
            PRETO     = colors.HexColor("#1A1A1A")
            CREME     = colors.HexColor("#F5F0E8")
            VINHO     = colors.HexColor("#8B1A2B")
            VINHO_CLR = colors.HexColor("#C41E3A")
            BEGE      = colors.HexColor("#D4C5A9")
            CINZA_CLR = colors.HexColor("#F0ECE4")
            CINZA_HDR = colors.HexColor("#3A3A3A")
            BRANCO    = colors.white

            # ── ESTILOS ──
            st_logo = ParagraphStyle("logo", fontSize=28, fontName="Times-Bold",
                                      textColor=PRETO, alignment=TA_LEFT, leading=30)
            st_emp  = ParagraphStyle("emp",  fontSize=8,  fontName="Helvetica",
                                      textColor=colors.HexColor("#555555"), alignment=TA_LEFT)
            st_td   = ParagraphStyle("td",   fontSize=7.5,fontName="Helvetica",
                                      textColor=PRETO, alignment=TA_LEFT)
            st_td_sm= ParagraphStyle("tdsm", fontSize=7,  fontName="Helvetica",
                                      textColor=PRETO, alignment=TA_LEFT)
            st_th   = ParagraphStyle("th",   fontSize=6.5,fontName="Helvetica-Bold",
                                      textColor=BRANCO, alignment=TA_CENTER)

            def P(txt, st=None, fs=7.5, bold=False, cor=None, align=TA_LEFT):
                if st: return Paragraph(str(txt), st)
                fn  = "Helvetica-Bold" if bold else "Helvetica"
                clr = cor or PRETO
                return Paragraph(str(txt),
                                  ParagraphStyle(f"x{abs(hash(str(txt)+str(fs)+str(bold)))%999999}",
                                                  fontSize=fs, fontName=fn,
                                                  textColor=clr, alignment=align))

            # 6 colunas — Tipo fica no cabeçalho do bloco, não na tabela
            col_props = [0.26, 0.15, 0.08, 0.11, 0.16, 0.24]
            HDRS = ["Favorecido / Beneficiário", "CPF/CNPJ",
                    "Data", "Valor (R$)", "Responsável", "Categoria"]

            def _cabecalho_boah():
                cab = [[
                    Table([[
                        [P("BOAH", st=st_logo)],
                        [P("LALUA COMERCIO DE MODAS LTDA  ·  SOLAR SERVIÇOS DE APOIO ADMIN",
                           st=st_emp)],
                    ]], colWidths=[doc.width * 0.5]),
                    Table([[
                        [P("AUTORIZAÇÃO DE PAGAMENTOS", bold=True, fs=10,
                           cor=PRETO, align=TA_RIGHT)],
                        [P(f"Período: {de_str}  a  {ate_str}", fs=8,
                           cor=colors.HexColor("#555555"), align=TA_RIGHT)],
                        [P(f"Emitido em {datetime.now().strftime('%d/%m/%Y às %H:%M')}",
                           fs=7, cor=colors.HexColor("#888888"), align=TA_RIGHT)],
                    ]], colWidths=[doc.width * 0.5]),
                ]]
                t = Table(cab, colWidths=[doc.width*0.5, doc.width*0.5])
                t.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,-1), CREME),
                    ("TOPPADDING",    (0,0), (-1,-1), 10),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 10),
                    ("LEFTPADDING",   (0,0), (0,-1),  14),
                    ("RIGHTPADDING",  (1,0), (1,-1),  14),
                    ("VALIGN",        (0,0), (-1,-1),  "MIDDLE"),
                ]))
                return [t, HRFlowable(width="100%", thickness=2,
                                       color=VINHO_CLR, spaceAfter=6)]

            def _linha_total(label, valor, bg, txt_cor=BRANCO, fs=8):
                cw = [doc.width*x for x in col_props]
                d  = [[P(label, bold=True, cor=txt_cor, fs=fs),
                        "", "",
                        P(f"R$ {valor:,.2f}", bold=True, cor=txt_cor,
                          fs=fs+0.5, align=TA_RIGHT),
                        "", ""]]
                t = Table(d, colWidths=cw)
                t.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,-1), bg),
                    ("SPAN",          (0,0), (2,0)),
                    ("SPAN",          (3,0), (5,0)),
                    ("ALIGN",         (3,0), (5,0), "RIGHT"),
                    ("TOPPADDING",    (0,0), (-1,-1), 5),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 5),
                    ("LEFTPADDING",   (0,0), (-1,-1), 8),
                    ("RIGHTPADDING",  (3,0), (5,0),   8),
                ]))
                return t

            story = []
            total_geral    = 0.0
            primeira_emp   = True
            empresas_lista = [e for e in ["LALUA", "SOLAR"] if e in empresas]

            for emp_nome in empresas_lista:
                nome_emp_full = ("LALUA COMERCIO DE MODAS LTDA" if emp_nome == "LALUA"
                                 else "SOLAR SERVIÇOS DE APOIO ADMIN")

                # Quebra de página entre empresas — cada uma começa numa página nova
                if not primeira_emp:
                    story.append(PageBreak())
                primeira_emp = False

                # Cabeçalho BOAH no topo de cada página de empresa
                story.extend(_cabecalho_boah())

                total_empresa = 0.0

                datas_ord = sorted(
                    empresas[emp_nome].keys(),
                    key=lambda d: (datetime.strptime(d, "%d/%m/%Y")
                                   if re.match(r'\d{2}/\d{2}/\d{4}', d)
                                   else datetime.max)
                )

                for data_dia in datas_ord:
                    tipos_dia = empresas[emp_nome][data_dia]
                    total_dia = 0.0

                    for tipo in ORDEM_TIPO:
                        if tipo not in tipos_dia:
                            continue
                        itens    = tipos_dia[tipo]
                        subtotal = sum(i["valor"] for i in itens)
                        total_dia     += subtotal
                        total_empresa += subtotal
                        total_geral   += subtotal

                        # Cabeçalho do bloco: empresa (esq) + data · tipo (dir) — LINHA ÚNICA
                        cab_bloco = [[
                            P(f"  {nome_emp_full}", bold=True, fs=8.5,
                              cor=BRANCO, align=TA_LEFT),
                            P(f"{data_dia}  ·  {tipo}", bold=True, fs=8.5,
                              cor=CREME, align=TA_RIGHT),
                        ]]
                        tbl_cb = Table(cab_bloco,
                                       colWidths=[doc.width*0.55, doc.width*0.45])
                        tbl_cb.setStyle(TableStyle([
                            ("BACKGROUND",    (0,0), (-1,-1), CINZA_HDR),
                            ("TOPPADDING",    (0,0), (-1,-1), 5),
                            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
                            ("LEFTPADDING",   (0,0), (0,-1),  8),
                            ("RIGHTPADDING",  (1,0), (1,-1),  8),
                            ("LINEBELOW",     (0,0), (-1,-1), 2, VINHO_CLR),
                            ("LINEBEFORE",    (0,0), (0,-1),  3, VINHO_CLR),
                        ]))
                        story.append(tbl_cb)

                        # Cabeçalho das colunas
                        hdr_r = [P(h, st=st_th) for h in HDRS]
                        rows  = [hdr_r]

                        for item in sorted(itens, key=lambda x: x["nome"]):
                            rows.append([
                                P(item["nome"][:55], st=st_td),
                                P(item["cnpj"],              fs=7),
                                P(item["data"],              fs=7,   align=TA_CENTER),
                                P(f"R$ {item['valor']:,.2f}",fs=7.5, bold=True, align=TA_RIGHT),
                                P(item.get("responsavel",""),fs=7,   align=TA_CENTER),
                                P(item.get("categoria",""),  st=st_td_sm),
                            ])

                        # Linha subtotal
                        nr = len(rows)
                        rows.append([
                            P(tipo, bold=True, cor=PRETO, fs=7.5),
                            "", "",
                            P(f"R$ {subtotal:,.2f}", bold=True,
                              cor=VINHO, fs=8.5, align=TA_RIGHT),
                            "", "",
                        ])

                        cw  = [doc.width*x for x in col_props]
                        tbl = Table(rows, colWidths=cw, repeatRows=1)
                        nrr = len(rows)
                        tbl.setStyle(TableStyle([
                            # Cabeçalho colunas
                            ("BACKGROUND",    (0,0), (-1,0), CINZA_HDR),
                            ("TEXTCOLOR",     (0,0), (-1,0), BRANCO),
                            ("ALIGN",         (0,0), (-1,0), "CENTER"),
                            ("FONTSIZE",      (0,0), (-1,0), 6.5),
                            # Dados — linhas alternadas
                            ("FONTSIZE",      (0,1), (-1,nrr-2), 7.5),
                            ("ROWBACKGROUNDS",(0,1), (-1,nrr-2), [BRANCO, CINZA_CLR]),
                            # Alinhamentos
                            ("ALIGN",         (2,1), (2,nrr-2), "CENTER"),  # Data
                            ("ALIGN",         (3,1), (3,nrr-2), "RIGHT"),   # Valor
                            ("ALIGN",         (4,1), (4,nrr-2), "CENTER"),  # Responsável
                            # Subtotal
                            ("BACKGROUND",    (0,nrr-1), (-1,nrr-1), CREME),
                            ("FONTNAME",      (0,nrr-1), (-1,nrr-1), "Helvetica-Bold"),
                            ("FONTSIZE",      (0,nrr-1), (-1,nrr-1), 7.5),
                            ("SPAN",          (0,nrr-1), (2,nrr-1)),
                            ("SPAN",          (3,nrr-1), (5,nrr-1)),
                            ("ALIGN",         (3,nrr-1), (5,nrr-1), "RIGHT"),
                            ("LINEBEFORE",    (0,0),     (0,-1),     3, VINHO_CLR),
                            # Grid elegante
                            ("LINEBELOW",     (0,0), (-1,0),      0.5, BEGE),
                            ("LINEBELOW",     (0,1), (-1,nrr-2),  0.3, BEGE),
                            # Padding
                            ("TOPPADDING",    (0,0), (-1,-1), 3),
                            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
                            ("LEFTPADDING",   (0,0), (-1,-1), 4),
                            ("RIGHTPADDING",  (0,0), (-1,-1), 4),
                            ("RIGHTPADDING",  (3,nrr-1), (5,nrr-1), 6),
                        ]))
                        story.append(tbl)
                        story.append(Spacer(1, 3))

                    # Total do dia
                    story.append(_linha_total(
                        f"Total  {data_dia}", total_dia, BEGE, txt_cor=PRETO, fs=8))
                    story.append(Spacer(1, 6))

                # Total da empresa
                story.append(_linha_total(
                    f"Total  {nome_emp_full}", total_empresa, PRETO, txt_cor=BRANCO, fs=9))
                story.append(Spacer(1, 10))

            # ── TOTAL GERAL (última página) ─────────────────────────────────
            story.append(HRFlowable(width="100%", thickness=1.5,
                                     color=VINHO_CLR, spaceBefore=2, spaceAfter=2))
            cw  = [doc.width*x for x in col_props]
            tgd = [[
                P("TOTAL GERAL", bold=True, fs=11, cor=BRANCO),
                "", "",
                P(f"R$ {total_geral:,.2f}", bold=True, fs=12,
                  cor=CREME, align=TA_RIGHT),
                "", "",
            ]]
            tbl_tg = Table(tgd, colWidths=cw)
            tbl_tg.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), VINHO),
                ("SPAN",          (0,0), (2,0)),
                ("SPAN",          (3,0), (5,0)),
                ("ALIGN",         (3,0), (5,0), "RIGHT"),
                ("TOPPADDING",    (0,0), (-1,-1), 10),
                ("BOTTOMPADDING", (0,0), (-1,-1), 10),
                ("LEFTPADDING",   (0,0), (-1,-1), 10),
                ("RIGHTPADDING",  (3,0), (5,0),   10),
            ]))
            story.append(tbl_tg)
            story.append(Spacer(1, 8))
            story.append(HRFlowable(width="100%", thickness=0.5, color=BEGE))
            story.append(P(
                "Documento interno · Uso exclusivo Contas a Pagar BOAH/SOLAR  ·  "
                f"Gerado automaticamente em {datetime.now().strftime('%d/%m/%Y %H:%M')}  ·  "
                "Robô Financeiro BOAH v17 — Roberto Cerqueira",
                fs=6, cor=colors.HexColor("#AAAAAA"), align=TA_CENTER
            ))

            doc.build(story)

            salvos = salvar_pagamentos_autorizados(self._aut_pagamentos)
            self._aut_lbl_status.config(
                text=f"✅ PDF gerado e {salvos} pagamentos salvos!", fg="#4ade80")
            os.startfile(caminho_pdf)

        except Exception as e:
            self._aut_lbl_status.config(text=f"❌ Erro ao gerar PDF: {e}", fg="#f87171")


    def _gerar_pdf_transferencias_filiais(self):
        """Gera PDF de Transferências Filiais → Matriz no padrão BOAH."""
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib import colors
            from reportlab.lib.units import cm, mm
            from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                            Paragraph, Spacer, HRFlowable,
                                            KeepTogether, PageBreak)
            from reportlab.lib.styles import ParagraphStyle
            from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
        except ImportError:
            from tkinter import messagebox
            messagebox.showerror("Erro",
                "reportlab não instalado.\nExecute: pip install reportlab")
            return

        # ── PALETA BOAH (mesma da Autorização) ──
        PRETO     = colors.HexColor("#1A1A1A")
        CREME     = colors.HexColor("#F5F0E8")
        VINHO     = colors.HexColor("#8B1A2B")
        VINHO_CLR = colors.HexColor("#C41E3A")
        BEGE      = colors.HexColor("#D4C5A9")
        CINZA_CLR = colors.HexColor("#F0ECE4")
        CINZA_HDR = colors.HexColor("#3A3A3A")
        BRANCO    = colors.white
        VERDE     = colors.HexColor("#1A6B3A")
        AZUL_PIX  = colors.HexColor("#1A3A6B")

        def P(txt, fs=8, bold=False, cor=None, align=TA_LEFT):
            fn  = "Helvetica-Bold" if bold else "Helvetica"
            clr = cor or PRETO
            return Paragraph(str(txt),
                ParagraphStyle(f"p{abs(hash(str(txt)+str(fs)+str(bold)))%999999}",
                               fontSize=fs, fontName=fn,
                               textColor=clr, alignment=align,
                               leading=fs*1.3))

        # Coleta transferências da aba autorização (tipo PIX FILIAL→MATRIZ)
        # e também do campo manual marcado como transferência
        filiais_map = {
            "LALUA - Barra":   "BARRA",
            "LALUA - Horto":   "HORTO",
            "LALUA - Vilas":   "VILAS",
            "LALUA - Paseo":   "PASEO",
            "LALUA - SDB":     "SDB",
            "LALUA - Matriz":  "MATRIZ",
            "SOLAR - Geral":   "SOLAR",
        }

        # Pede dados via janela modal
        from tkinter import Toplevel, StringVar, BooleanVar
        win = Toplevel(self.root)
        win.title("Relatório de Transferências Filiais → Matriz")
        win.geometry("560x460")
        win.configure(bg="#0f172a")
        win.grab_set()

        bg_w    = "#000000"
        surf_w  = "#111111"
        acc_w   = "#a3e635"
        txt_w   = "#e2e8f0"
        muted_w = "#94a3b8"

        tk.Label(win, text="💸  Relatório de Transferências Filiais → Matriz",
                 font=("Segoe UI",11,"bold"), fg=acc_w, bg=bg_w
                 ).pack(anchor="w", padx=20, pady=(14,4))
        tk.Label(win,
                 text="Preencha os dados de cada filial que enviou PIX para a matriz.",
                 font=("Segoe UI",8), fg=muted_w, bg=bg_w
                 ).pack(anchor="w", padx=20, pady=(0,8))

        frm_campos = tk.Frame(win, bg=surf_w, padx=16, pady=12)
        frm_campos.pack(fill="x", padx=20)

        def lbl(t, r, c):
            tk.Label(frm_campos, text=t, fg=muted_w, bg=surf_w,
                     font=("Segoe UI",8)).grid(row=r, column=c, sticky="w", pady=3, padx=(0,6))

        def ent(r, c, w=14, val=""):
            e = tk.Entry(frm_campos, font=("Segoe UI",9),
                         bg="#0a0f1e", fg=txt_w,
                         insertbackground=acc_w, relief="flat", bd=2, width=w)
            e.grid(row=r, column=c, sticky="w", padx=(0,10))
            if val: e.insert(0, val)
            return e

        # Cabeçalho
        for i, h in enumerate(["Filial","Data","Valor Transferido","Referência/Obs"]):
            tk.Label(frm_campos, text=h, fg=acc_w, bg=surf_w,
                     font=("Segoe UI",7,"bold")).grid(row=0, column=i, sticky="w", padx=(0,10))

        FILIAIS_PADRAO = [
            "LALUA - Barra", "LALUA - Horto", "LALUA - Vilas",
            "LALUA - Paseo", "LALUA - SDB",
        ]

        linhas_filial = []
        data_hoje = datetime.now().strftime("%d/%m/%Y")

        for i, fil in enumerate(FILIAIS_PADRAO):
            fil_var  = tk.StringVar(value=fil)
            ttk.Combobox(frm_campos, textvariable=fil_var,
                         values=list(filiais_map.keys()),
                         state="normal", width=16, font=("Segoe UI",8)
                         ).grid(row=i+1, column=0, sticky="w", padx=(0,10), pady=2)
            dt_e   = ent(i+1, 1, 11, data_hoje)
            _aplicar_mascara_data(dt_e)
            val_e  = ent(i+1, 2, 13)
            _aplicar_mascara_valor(val_e)
            obs_e  = ent(i+1, 3, 20)
            linhas_filial.append((fil_var, dt_e, val_e, obs_e))

        # Campos gerais
        frm_geral = tk.Frame(win, bg=surf_w, padx=16, pady=8)
        frm_geral.pack(fill="x", padx=20, pady=(6,0))

        tk.Label(frm_geral, text="Período referência:", fg=muted_w, bg=surf_w,
                 font=("Segoe UI",8)).grid(row=0, column=0, sticky="w")
        de_var  = tk.StringVar(value=self._aut_de.get() if hasattr(self,"_aut_de") else data_hoje)
        ate_var = tk.StringVar(value=self._aut_ate.get() if hasattr(self,"_aut_ate") else data_hoje)
        tk.Entry(frm_geral, textvariable=de_var, font=("Segoe UI",9),
                 bg="#0a0f1e", fg=txt_w, insertbackground=acc_w,
                 relief="flat", bd=2, width=11).grid(row=0, column=1, padx=(4,8))
        tk.Label(frm_geral, text="até", fg=muted_w, bg=surf_w,
                 font=("Segoe UI",8)).grid(row=0, column=2)
        tk.Entry(frm_geral, textvariable=ate_var, font=("Segoe UI",9),
                 bg="#0a0f1e", fg=txt_w, insertbackground=acc_w,
                 relief="flat", bd=2, width=11).grid(row=0, column=3, padx=(4,0))

        lbl_status = tk.Label(win, text="", font=("Segoe UI",9),
                               fg="#4ade80", bg=bg_w)
        lbl_status.pack(anchor="w", padx=20, pady=4)

        def _gerar():
            try:
                # Coleta dados das linhas da modal
                transferencias = []
                for (fil_var, dt_e, val_e, obs_e) in linhas_filial:
                    filial = fil_var.get().strip()
                    dt     = dt_e.get().strip()
                    val_s  = val_e.get().strip()
                    obs    = obs_e.get().strip()
                    if not filial or not val_s:
                        continue
                    try:
                        valor = _parse_valor(val_s)
                        if valor <= 0: continue
                    except: continue
                    transferencias.append({
                        "filial": filial,
                        "data":   dt,
                        "valor":  valor,
                        "obs":    obs,
                    })

                if not transferencias:
                    lbl_status.config(text="⚠️ Informe ao menos uma transferência.",
                                      fg="#fbbf24")
                    return

                de_str  = de_var.get().strip()
                ate_str = ate_var.get().strip()
                total   = sum(t["valor"] for t in transferencias)

                # ── GERA O PDF ──
                data_hoje_str = datetime.now().strftime("%d_%m_%Y")
                nome_arq = f"TRANSFERENCIAS_FILIAIS_{data_hoje_str}.pdf"
                caminho  = os.path.join(PASTA_ATUAL, nome_arq)

                PAGE = landscape(A4)
                doc  = SimpleDocTemplate(
                    caminho, pagesize=PAGE,
                    leftMargin=1.8*cm, rightMargin=1.8*cm,
                    topMargin=1.4*cm,  bottomMargin=1.4*cm,
                    title=f"BOAH — Transferências Filiais → Matriz {de_str}"
                )

                story = []

                # ── CABEÇALHO ──
                cab = [[
                    Table([[
                        [P("BOAH", fs=28, bold=True, cor=PRETO)],
                        [P("LALUA COMERCIO DE MODAS LTDA  ·  SOLAR SERVIÇOS DE APOIO ADMIN",
                           fs=8, cor=colors.HexColor("#555555"))],
                    ]], colWidths=[doc.width*0.5]),
                    Table([[
                        [P("TRANSFERÊNCIAS FILIAIS → MATRIZ", bold=True, fs=10,
                           cor=PRETO, align=TA_RIGHT)],
                        [P(f"Período: {de_str}  a  {ate_str}", fs=8,
                           cor=colors.HexColor("#555555"), align=TA_RIGHT)],
                        [P(f"Emitido em {datetime.now().strftime('%d/%m/%Y às %H:%M')}",
                           fs=7, cor=colors.HexColor("#888888"), align=TA_RIGHT)],
                    ]], colWidths=[doc.width*0.5]),
                ]]
                t_cab = Table(cab, colWidths=[doc.width*0.5, doc.width*0.5])
                t_cab.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,-1), CREME),
                    ("TOPPADDING",    (0,0), (-1,-1), 10),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 10),
                    ("LEFTPADDING",   (0,0), (0,-1),  14),
                    ("RIGHTPADDING",  (1,0), (1,-1),  14),
                    ("VALIGN",        (0,0), (-1,-1),  "MIDDLE"),
                ]))
                story.append(t_cab)
                story.append(HRFlowable(width="100%", thickness=2,
                                         color=VINHO_CLR, spaceAfter=8))

                # ── BLOCO: RESUMO DAS TRANSFERÊNCIAS RECEBIDAS ──
                cb_bloco = [[
                    P("  DETALHAMENTO DAS TRANSFERÊNCIAS RECEBIDAS", bold=True,
                      fs=10, cor=BRANCO, align=TA_LEFT),
                    P("Filiais → Conta Matriz LALUA · Ag. 0334 · Cc. 98775-7",
                      fs=8.5, cor=CREME, align=TA_RIGHT),
                ]]
                t_cb = Table(cb_bloco, colWidths=[doc.width*0.55, doc.width*0.45])
                t_cb.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,-1), CINZA_HDR),
                    ("TOPPADDING",    (0,0), (-1,-1), 8),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 8),
                    ("LEFTPADDING",   (0,0), (0,-1),  10),
                    ("RIGHTPADDING",  (1,0), (1,-1),  10),
                    ("LINEBELOW",     (0,0), (-1,-1), 2, VINHO_CLR),
                    ("LINEBEFORE",    (0,0), (0,-1),  4, VINHO_CLR),
                ]))
                story.append(t_cb)

                # Tabela de transferências (Expandida)
                cw_transf = [doc.width*0.28, doc.width*0.12,
                              doc.width*0.15, doc.width*0.25, doc.width*0.20]
                hdrs_t = ["Filial Remetente", "Data", "Valor (R$)",
                          "Referência / Observação", "Status Conferência"]
                rows_t = [[P(h, fs=8.5, bold=True, cor=BRANCO, align=TA_CENTER)
                           for h in hdrs_t]]

                for tr in sorted(transferencias, key=lambda x: x["filial"]):
                    rows_t.append([
                        P(f"  {tr['filial']}", fs=10, bold=True, cor=AZUL_PIX),
                        P(tr["data"], fs=10, align=TA_CENTER),
                        P(f"R$ {tr['valor']:,.2f}", fs=11, bold=True,
                          cor=VERDE, align=TA_RIGHT),
                        P(tr["obs"] or "—", fs=9),
                        P("( )", fs=11, align=TA_CENTER),
                    ])

                # Linha total transferido
                rows_t.append([
                    P("TOTAL GERAL TRANSFERIDO", bold=True, fs=11, cor=BRANCO),
                    "", "",
                    P(f"R$ {total:,.2f}", bold=True, fs=12.5,
                      cor=CREME, align=TA_RIGHT),
                    "",
                ])

                nr = len(rows_t)
                # Altura das linhas para preencher melhor o espaço (mínimo 30pt por linha de dados)
                h_rows = [None] + [32]*(nr-2) + [38] 
                
                tbl_t = Table(rows_t, colWidths=cw_transf, rowHeights=h_rows, repeatRows=1)
                tbl_t.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,0), CINZA_HDR),
                    ("TEXTCOLOR",     (0,0), (-1,0), BRANCO),
                    ("ROWBACKGROUNDS",(0,1), (-1,nr-2), [BRANCO, CINZA_CLR]),
                    ("BACKGROUND",    (0,nr-1), (-1,nr-1), PRETO),
                    ("TEXTCOLOR",     (0,nr-1), (-1,nr-1), BRANCO),
                    ("SPAN",          (0,nr-1), (2,nr-1)),
                    ("SPAN",          (3,nr-1), (4,nr-1)),
                    ("ALIGN",         (3,nr-1), (4,nr-1), "RIGHT"),
                    ("ALIGN",         (1,1), (1,nr-2), "CENTER"),
                    ("ALIGN",         (2,1), (2,nr-2), "RIGHT"),
                    ("ALIGN",         (4,1), (4,nr-2), "CENTER"),
                    ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
                    ("LINEBEFORE",    (0,0), (0,-1), 4, VINHO_CLR),
                    ("LINEBELOW",     (0,0), (-1,0), 0.5, BEGE),
                    ("LINEBELOW",     (0,1), (-1,nr-2), 0.3, BEGE),
                    ("TOPPADDING",    (0,0), (-1,-1), 8),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 8),
                    ("LEFTPADDING",   (0,0), (-1,-1), 8),
                    ("RIGHTPADDING",  (0,0), (-1,-1), 8),
                ]))
                story.append(tbl_t)
                story.append(Spacer(1, 25))

                # ── ÁREA DE OBSERVAÇÕES (Opcional, para preenchimento manual) ──
                story.append(HRFlowable(width="100%", thickness=1, color=BEGE))
                story.append(P("OBSERVAÇÕES ADICIONAIS:", fs=8, bold=True, cor=colors.HexColor("#666666")))
                story.append(Spacer(1, 40)) # Espaço em branco para anotação manual
                story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#DDDDDD")))
                story.append(Spacer(1, 25))

                # ── ASSINATURAS ──
                story.append(HRFlowable(width="100%", thickness=0.5, color=BEGE))
                story.append(Spacer(1, 15))

                cw_ass = [doc.width*0.25] * 4
                ass_data = [[
                    Table([[P("_"*28, fs=8), P("Conferido por / Matriz",
                              fs=7.5, cor=colors.HexColor("#777777"), align=TA_CENTER)]],
                           colWidths=[doc.width*0.25]),
                    Table([[P("_"*28, fs=8), P("Aprovado por / Financeiro",
                              fs=7.5, cor=colors.HexColor("#777777"), align=TA_CENTER)]],
                           colWidths=[doc.width*0.25]),
                    Table([[P("_"*28, fs=8), P("Responsável / Filial",
                              fs=7.5, cor=colors.HexColor("#777777"), align=TA_CENTER)]],
                           colWidths=[doc.width*0.25]),
                    Table([[P("_"*28, fs=8), P(f"Data: ____/____/______",
                              fs=7.5, cor=colors.HexColor("#777777"), align=TA_CENTER)]],
                           colWidths=[doc.width*0.25]),
                ]]
                tbl_ass = Table(ass_data, colWidths=cw_ass)
                tbl_ass.setStyle(TableStyle([
                    ("ALIGN",  (0,0), (-1,-1), "CENTER"),
                    ("VALIGN", (0,0), (-1,-1), "BOTTOM"),
                ]))
                story.append(tbl_ass)
                story.append(Spacer(1, 10))

                # ── RODAPÉ ──
                story.append(HRFlowable(width="100%", thickness=0.5, color=BEGE))
                story.append(P(
                    "Documento interno · Uso exclusivo Contas a Pagar BOAH/SOLAR  ·  "
                    f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}  ·  "
                    "Robô Financeiro BOAH v17 — Roberto Cerqueira",
                    fs=6, cor=colors.HexColor("#AAAAAA"), align=TA_CENTER
                ))

                doc.build(story)

                lbl_status.config(
                    text=f"✅ PDF gerado: {nome_arq}", fg="#4ade80")
                win.after(1200, win.destroy)
                os.startfile(caminho)

            except Exception as e:
                lbl_status.config(text=f"❌ Erro: {e}", fg="#f87171")

        # Botões da janela modal
        frm_btns = tk.Frame(win, bg=bg_w)
        frm_btns.pack(fill="x", padx=20, pady=(8,12))
        tk.Button(frm_btns, text="📄 Gerar PDF",
                  font=("Segoe UI",10,"bold"), bg="#7c3aed", fg="white",
                  relief="flat", bd=0, padx=18, pady=7, cursor="hand2",
                  command=_gerar).pack(side="right")
        tk.Button(frm_btns, text="Cancelar",
                  font=("Segoe UI",9), bg=surf_w, fg=muted_w,
                  relief="flat", bd=0, padx=12, pady=7, cursor="hand2",
                  command=win.destroy).pack(side="right", padx=(0,8))


    def _enviar_para_autorizacao(self):
        """Abre a aba de Autorização com os comprovantes selecionados como referência."""
        selecionados = self.tree.selection()
        if not selecionados:
            itens = self._resultados_busca[:50]  # máx 50
        else:
            indices = [self.tree.index(s) for s in selecionados]
            itens = [self._resultados_busca[i] for i in indices]

        if not itens:
            self.lbl_status_envio.config(
                text="⚠️ Nenhum item selecionado.", fg="#fbbf24")
            return

        # Converte para o formato da aba autorização e injeta
        pgtos = []
        for item in itens:
            pgtos.append({
                "nome":        item['nome'],
                "cnpj":        "",
                "tipo":        "Comprovante",
                "data":        item['data'].strftime("%d/%m/%Y"),
                "data_obj":    item['data'],
                "valor":       item['valor'],
                "status":      "Aprovada",
                "responsavel": "",
                "categoria":   item['categoria'],
                "empresa":     item.get('empresa','LALUA'),
            })

        self._aut_pagamentos = pgtos
        self.root.after(0, self._aut_popular_treeview)

        # Vai para aba Autorização
        try:
            self.notebook.select(6)  # índice da aba Autorização
        except: pass

        self.lbl_status_envio.config(
            text=f"✅ {len(pgtos)} item(ns) enviados para Autorização de Pagamentos",
            fg="#f472b6")

    def _gerar_relatorio_pagamentos_excel(self):
        """Gera e salva o relatório analítico dos pagamentos no banco de dados."""
        from utils.relatorios_processor import gerar_excel_pagamentos_autorizados
        
        caminho_saida = filedialog.asksaveasfilename(
            title="Salvar Relatório Analítico de Pagamentos",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"Relatorio_Pagamentos_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        
        if not caminho_saida:
            return
            
        sucesso, msg = gerar_excel_pagamentos_autorizados(caminho_saida)
        
        if sucesso:
            self._aut_lbl_resumo.config(text=f"✅ Relatório salvo: {os.path.basename(caminho_saida)}", fg="#4ade80")
            os.startfile(caminho_saida)
        else:
            self._aut_lbl_resumo.config(text=f"❌ Erro: {msg}", fg="#f87171")


    def _build_aba_manual(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        """Aba dedicada para lançamentos manuais."""
        hdr = tk.Frame(parent, bg=bg)
        hdr.pack(fill="x", padx=20, pady=(12, 10))
        tk.Label(hdr, text="➕ Novo Lançamento Manual",
                 font=("Segoe UI", 13, "bold"), fg=accent, bg=bg).pack(side="left")

        frame_manual = tk.Frame(parent, bg=surface, padx=20, pady=20)
        frame_manual.pack(fill="x", padx=20, pady=10)

        def _lbl(t):
            return tk.Label(frame_manual, text=t, fg=muted, bg=surface, font=("Segoe UI", 10))
        def _ent(w=30):
            return tk.Entry(frame_manual, font=("Segoe UI", 10), bg="#0a0f1e", fg=text,
                            insertbackground=accent, relief="flat", bd=4, width=w)

        _lbl("Nome do Beneficiário:").grid(row=0, column=0, sticky="w", pady=8)
        self._man_nome = _ent(40)
        self._man_nome.grid(row=0, column=1, sticky="w", padx=10, columnspan=3)

        _lbl("Empresa:").grid(row=1, column=0, sticky="w", pady=8)
        self._man_emp_var = tk.StringVar(value="LALUA")
        ttk.Combobox(frame_manual, textvariable=self._man_emp_var, values=["LALUA","SOLAR"],
                     state="readonly", width=15).grid(row=1, column=1, sticky="w", padx=10)

        _lbl("Tipo de Pagamento:").grid(row=1, column=2, sticky="w", pady=8)
        self._man_tipo_var = tk.StringVar(value="PIX")
        ttk.Combobox(frame_manual, textvariable=self._man_tipo_var,
                     values=["PIX","TED","BOLETO","DARF","GPS","DAS","CONCESSIONÁRIA","PIX FILIAL→MATRIZ"],
                     state="readonly", width=25).grid(row=1, column=3, sticky="w", padx=10)

        _lbl("Data:").grid(row=2, column=0, sticky="w", pady=8)
        self._man_data = _ent(15)
        self._man_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self._man_data.grid(row=2, column=1, sticky="w", padx=10)

        _lbl("Valor (R$):").grid(row=2, column=2, sticky="w", pady=8)
        self._man_valor = _ent(15)
        self._man_valor.grid(row=2, column=3, sticky="w", padx=10)

        _lbl("Categoria:").grid(row=3, column=0, sticky="w", pady=8)
        self._man_cat_var = tk.StringVar()
        ttk.Combobox(frame_manual, textvariable=self._man_cat_var, values=CATEGORIAS_LISTA,
                     state="normal", width=38).grid(row=3, column=1, sticky="w", padx=10, columnspan=2)

        _lbl("Responsável:").grid(row=4, column=0, sticky="w", pady=8)
        self._man_resp_var = tk.StringVar()
        ttk.Combobox(frame_manual, textvariable=self._man_resp_var, values=RESPONSAVEIS_LISTA,
                     state="readonly", width=20).grid(row=4, column=1, sticky="w", padx=10)

        _lbl("Observação:").grid(row=5, column=0, sticky="w", pady=8)
        self._man_obs = _ent(40)
        self._man_obs.grid(row=5, column=1, sticky="w", padx=10, columnspan=3)

        self._man_status = tk.Label(frame_manual, text="", font=("Segoe UI", 9), fg=yellow, bg=surface)
        self._man_status.grid(row=6, column=0, columnspan=4, pady=15)

        def _adicionar():
            nome = self._man_nome.get().strip()
            valor_s = self._man_valor.get().strip()
            data = self._man_data.get().strip()
            if not nome or not valor_s:
                self._man_status.config(text="⚠️ Nome e Valor são obrigatórios!", fg=yellow)
                return
            try:
                valor = _parse_valor(valor_s)
                data_obj = datetime.strptime(data, "%d/%m/%Y")
                self._aut_pagamentos.append({
                    "nome": nome, "cnpj": "", "tipo": self._man_tipo_var.get(),
                    "data": data, "data_obj": data_obj, "valor": valor,
                    "status": "Aprovada", "responsavel": self._man_resp_var.get() or "ADM/FINANCEIRO",
                    "categoria": self._man_cat_var.get() or "A CLASSIFICAR",
                    "empresa": self._man_emp_var.get(), "observacao": self._man_obs.get(), "manual": True
                })
                self._man_status.config(text=f"✅ Adicionado: {nome} - R$ {valor:,.2f}", fg=green)
                self._man_nome.delete(0, "end")
                self._man_valor.delete(0, "end")
                self._man_obs.delete(0, "end")
                self._aut_popular_treeview()
            except Exception as e:
                self._man_status.config(text=f"❌ Erro: {str(e)}", fg="#f87171")

        tk.Button(frame_manual, text="➕ ADICIONAR AO RELATÓRIO", font=("Segoe UI", 10, "bold"),
                  bg=green, fg="#0f172a", relief="flat", bd=0, padx=25, pady=10, cursor="hand2",
                  command=_adicionar).grid(row=7, column=0, columnspan=4, pady=10)
