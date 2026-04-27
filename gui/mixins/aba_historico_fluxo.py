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


class AbaHistFluxoMixin:

    # ==========================================================================
    # ── ABA 5: HISTÓRICO DE FLUXO ──────────────────────────────────────────
    # ==========================================================================

    def _build_aba_historico_fluxo(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        """Aba que lê todos os dashboards Excel gerados e monta histórico de saldo por conta."""

        # ── TÍTULO ──
        hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.pack(fill="x", padx=25, pady=(20, 10))
        
        ctk.CTkLabel(hdr, text="📈 HISTÓRICO DE FLUXO", font=("Segoe UI", 16, "bold"), text_color=accent).pack(side="left")
        ctk.CTkLabel(parent, text="Acumula os saldos de todos os extratos processados na biblioteca.", font=("Segoe UI", 11), text_color=muted).pack(anchor="w", padx=25, pady=(0, 15))

        # ── CONTROLES ──
        ctrl = ctk.CTkFrame(parent, fg_color=surface, corner_radius=12, border_width=1, border_color=border)
        ctrl.pack(fill="x", padx=25, pady=(0, 10))

        ctk.CTkLabel(ctrl, text="CONTA:", font=("Segoe UI", 9, "bold"), text_color=muted).pack(side="left", padx=(15, 5))
        self._fluxo_conta_var = tk.StringVar(value="Todas")
        self._fluxo_conta_cb = ctk.CTkOptionMenu(ctrl, variable=self._fluxo_conta_var, values=["Todas"], fg_color=bg, button_color=border, width=180, height=32, command=lambda x: self._filtrar_fluxo())
        self._fluxo_conta_cb.pack(side="left", padx=5, pady=10)

        ctk.CTkButton(
            ctrl, text="🔄 ATUALIZAR", font=("Segoe UI", 10, "bold"), 
            fg_color=accent, text_color="#000000", hover_color=green,
            width=120, height=32, command=self._carregar_historico_fluxo
        ).pack(side="left", padx=15)

        ctk.CTkButton(
            ctrl, text="📥 EXCEL", font=("Segoe UI", 10, "bold"), 
            fg_color=border, text_color=accent, hover_color="#222",
            width=100, height=32, command=self._exportar_fluxo_excel
        ).pack(side="left")

        self._fluxo_lbl_status = ctk.CTkLabel(ctrl, text="", font=("Segoe UI", 11), text_color=muted)
        self._fluxo_lbl_status.pack(side="right", padx=15)

        # ── CARDS DE RESUMO ──
        self._fluxo_cards_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self._fluxo_cards_frame.pack(fill="x", padx=25, pady=(0, 10))

        # ── TREEVIEW DE HISTÓRICO ──
        tree_container = ctk.CTkFrame(parent, fg_color=bg, corner_radius=12, border_width=1, border_color=border)
        tree_container.pack(fill="both", expand=True, padx=25, pady=10)

        colunas = ("Período", "Conta", "Empresa", "Saldo Anterior", "Créditos", "Débitos", "Saldo Final", "Variação")
        self._fluxo_tree = ttk.Treeview(tree_container, columns=colunas, show="headings")
        
        style = ttk.Style()
        style.configure("Fluxo.Treeview", background=bg, foreground=text, fieldbackground=bg, rowheight=35, borderwidth=0, font=("Segoe UI", 10))
        style.configure("Fluxo.Treeview.Heading", background=surface, foreground=accent, relief="flat", font=("Segoe UI", 10, "bold"))
        style.map("Fluxo.Treeview", background=[("selected", border)])
        self._fluxo_tree.configure(style="Fluxo.Treeview")

        larguras = [100, 180, 90, 120, 120, 120, 120, 110]
        for col, w in zip(colunas, larguras):
            self._fluxo_tree.heading(col, text=col.upper(), command=lambda c=col: self._ordenar_fluxo(c))
            self._fluxo_tree.column(col, width=w, anchor="e" if col in ("Saldo Anterior", "Créditos", "Débitos", "Saldo Final", "Variação") else "w")

        self._fluxo_tree.tag_configure("positivo", foreground="#4ade80")
        self._fluxo_tree.tag_configure("negativo", foreground="#f87171")
        self._fluxo_tree.tag_configure("neutro",   foreground="#94a3b8")
        self._fluxo_tree.tag_configure("matriz",   background="#1c1033")

        scroll_f = ctk.CTkScrollbar(tree_container, orientation="vertical", command=self._fluxo_tree.yview)
        self._fluxo_tree.configure(yscrollcommand=scroll_f.set)
        
        self._fluxo_tree.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scroll_f.pack(side="right", fill="y", padx=(0, 5), pady=5)

        # ── TOTAIS ──
        self._fluxo_lbl_totais = ctk.CTkLabel(parent, text="", font=("Consolas", 11, "bold"), text_color=muted)
        self._fluxo_lbl_totais.pack(anchor="e", padx=30, pady=(0, 20))

        # Dados internos
        self._fluxo_dados = []
        self._fluxo_sort_col = None
        self._fluxo_sort_rev = False

        self._carregar_historico_fluxo()


    def _carregar_historico_fluxo(self):
        """Lê todos os Excel de Dashboard em EXTRATOS_PROCESSADOS e monta o histórico."""
        if not EXCEL_OK:
            self._fluxo_lbl_status.config(text="⚠️ openpyxl não instalado.", fg="#fbbf24")
            return

        os.makedirs(PASTA_EXTRATOS, exist_ok=True)
        arquivos = sorted([
            f for f in os.listdir(PASTA_EXTRATOS)
            if f.lower().endswith(".xlsx") and f.startswith("Dashboard_")
        ])

        if not arquivos:
            self._fluxo_lbl_status.config(
                text="Nenhum dashboard encontrado em EXTRATOS_PROCESSADOS.", fg="#fbbf24")
            return

        self._fluxo_lbl_status.config(text=f"⏳ Lendo {len(arquivos)} dashboard(s)...", fg="#38bdf8")
        self.root.update_idletasks()

        def carregar():
            linhas = []
            contas_vistas = set()
            try:
                for arq in arquivos:
                    caminho = os.path.join(PASTA_EXTRATOS, arq)
                    # Extrai período do nome do arquivo: Dashboard_NOME_YYYYMMDD_HHMM.xlsx
                    try:
                        partes_nome = arq.replace("Dashboard_", "").replace(".xlsx", "").split("_")
                        # Últimas duas partes são data e hora: YYYYMMDD_HHMM
                        data_arq = partes_nome[-2]
                        periodo_label = f"{data_arq[6:8]}/{data_arq[4:6]}/{data_arq[0:4]}"
                    except Exception:
                        periodo_label = arq[:16]

                    try:
                        wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
                    except Exception:
                        continue

                    # Procura aba "📊 Resumo Geral" ou primeira aba
                    aba_resumo = None
                    for nome_aba in wb.sheetnames:
                        if "Resumo" in nome_aba or "resumo" in nome_aba:
                            aba_resumo = wb[nome_aba]
                            break
                    if aba_resumo is None:
                        aba_resumo = wb.active

                    # Lê linhas de dados (ignora header e totais)
                    rows = list(aba_resumo.iter_rows(values_only=True))
                    for row in rows[2:]:  # pula título (linha 1) e cabeçalho (linha 2)
                        if not row or not row[0]:
                            continue
                        conta_nome = str(row[0]).strip()
                        if conta_nome in ("TOTAL GERAL", "", "None", "Conta"):
                            continue
                        try:
                            empresa   = str(row[1]).strip() if row[1] else ""
                            saldo_ant = float(row[2]) if row[2] is not None else 0.0
                            creditos  = float(row[3]) if row[3] is not None else 0.0
                            debitos   = float(row[4]) if row[4] is not None else 0.0
                            saldo_fin = float(row[5]) if row[5] is not None else 0.0
                        except Exception:
                            continue

                        variacao = saldo_fin - saldo_ant
                        linhas.append({
                            "periodo":    periodo_label,
                            "conta":      conta_nome,
                            "empresa":    empresa,
                            "saldo_ant":  saldo_ant,
                            "creditos":   creditos,
                            "debitos":    debitos,
                            "saldo_fin":  saldo_fin,
                            "variacao":   variacao,
                            "arquivo":    arq,
                        })
                        contas_vistas.add(conta_nome)

                    wb.close()

            except Exception as e:
                self.root.after(0, lambda: self._fluxo_lbl_status.config(
                    text=f"❌ Erro: {e}", fg="#f87171"))
                return

            self._fluxo_dados = linhas
            contas_lista = ["Todas"] + sorted(contas_vistas)

            def atualizar_ui():
                self._fluxo_conta_cb.config(values=contas_lista)
                self._fluxo_conta_var.set("Todas")
                self._filtrar_fluxo()
                self._fluxo_lbl_status.config(
                    text=f"✅ {len(arquivos)} dashboard(s) | {len(linhas)} linha(s)", fg="#4ade80")
                self._atualizar_cards_fluxo()

            self.root.after(0, atualizar_ui)

        threading.Thread(target=carregar, daemon=True).start()


    def _filtrar_fluxo(self):
        """Aplica filtro de conta e atualiza a treeview."""
        conta_filtro = self._fluxo_conta_var.get()

        for row in self._fluxo_tree.get_children():
            self._fluxo_tree.delete(row)

        dados = self._fluxo_dados
        if conta_filtro != "Todas":
            dados = [d for d in dados if d["conta"] == conta_filtro]

        total_deb = total_cred = total_var = 0.0
        for d in dados:
            variacao = d["variacao"]
            tag = "positivo" if variacao > 0 else ("negativo" if variacao < 0 else "neutro")
            if "Matriz" in d["conta"] or "98775" in d["conta"]:
                tag = "matriz"

            self._fluxo_tree.insert("", "end", tags=(tag,), values=(
                d["periodo"],
                d["conta"],
                d["empresa"],
                f"R$ {d['saldo_ant']:>12,.2f}",
                f"R$ {d['creditos']:>12,.2f}",
                f"R$ {d['debitos']:>12,.2f}",
                f"R$ {d['saldo_fin']:>12,.2f}",
                f"{'▲' if variacao >= 0 else '▼'} R$ {abs(variacao):,.2f}",
            ))
            total_deb  += d["debitos"]
            total_cred += d["creditos"]
            total_var  += variacao

        sinal = "▲" if total_var >= 0 else "▼"
        self._fluxo_lbl_totais.config(
            text=f"Total Créditos: R$ {total_cred:,.2f}   |   "
                 f"Total Débitos: R$ {total_deb:,.2f}   |   "
                 f"Variação líquida: {sinal} R$ {abs(total_var):,.2f}"
        )


    def _ordenar_fluxo(self, coluna):
        """Ordena a treeview ao clicar no cabeçalho."""
        mapa = {
            "Período": "periodo", "Conta": "conta", "Empresa": "empresa",
            "Saldo Anterior": "saldo_ant", "Créditos": "creditos",
            "Débitos": "debitos", "Saldo Final": "saldo_fin", "Variação": "variacao"
        }
        chave = mapa.get(coluna, "periodo")
        if self._fluxo_sort_col == chave:
            self._fluxo_sort_rev = not self._fluxo_sort_rev
        else:
            self._fluxo_sort_col = chave
            self._fluxo_sort_rev = False

        conta_filtro = self._fluxo_conta_var.get()
        dados = self._fluxo_dados if conta_filtro == "Todas" else \
                [d for d in self._fluxo_dados if d["conta"] == conta_filtro]
        self._fluxo_dados_ordenados = sorted(dados, key=lambda x: x[chave],
                                              reverse=self._fluxo_sort_rev)
        # Reaplica com dados ordenados temporariamente
        _backup = self._fluxo_dados
        _conta  = self._fluxo_conta_var.get()
        self._fluxo_dados = self._fluxo_dados_ordenados
        self._fluxo_conta_var.set("Todas")  # força reload completo
        self._filtrar_fluxo()
        self._fluxo_dados = _backup
        self._fluxo_conta_var.set(_conta)


    def _atualizar_cards_fluxo(self, parent_bg="#000000", surface="#111111", border="#222222", accent="#CCFF00", text="#FFFFFF", muted="#525252"):
        """Monta os stat cards de saldo mais recente por conta."""
        for w in self._fluxo_cards_frame.winfo_children():
            w.destroy()

        if not self._fluxo_dados:
            return

        # Para cada conta, pega o saldo final mais recente
        ultimos = {}
        for d in self._fluxo_dados:
            c = d["conta"]
            if c not in ultimos or d["periodo"] > ultimos[c]["periodo"]:
                ultimos[c] = d

        for conta, d in sorted(ultimos.items()):
            card = ctk.CTkFrame(self._fluxo_cards_frame, fg_color=surface, corner_radius=12, border_width=1, border_color=border)
            card.pack(side="left", expand=True, fill="x", padx=(0, 10), pady=5)
            
            nome_curto = conta.split("-")[-1].strip()[:14] if "-" in conta else conta[:14]
            ctk.CTkLabel(card, text=nome_curto.upper(), font=("Segoe UI", 9, "bold"), text_color=muted).pack(anchor="w", padx=15, pady=(10, 0))
            
            saldo = d["saldo_fin"]
            cor_saldo = accent if saldo >= 0 else "#FF3333"
            if saldo < 100_000 and "Matriz" in conta:
                cor_saldo = "#FFCC00"
                
            ctk.CTkLabel(card, text=f"R$ {saldo:,.0f}", font=("Segoe UI", 16, "bold"), text_color=cor_saldo).pack(anchor="w", padx=15, pady=(0, 2))
            
            var = d["variacao"]
            sinal = "▲" if var >= 0 else "▼"
            cor_var = accent if var >= 0 else "#FF3333"
            ctk.CTkLabel(card, text=f"{sinal} {abs(var):,.0f} ({d['periodo']})", font=("Consolas", 10), text_color=cor_var).pack(anchor="w", padx=15, pady=(0, 12))


    def _exportar_fluxo_excel(self):
        """Exporta o histórico completo de fluxo para um Excel."""
        if not EXCEL_OK:
            self._fluxo_lbl_status.config(text="⚠️ openpyxl não instalado.", fg="#fbbf24")
            return
        if not self._fluxo_dados:
            self._fluxo_lbl_status.config(text="⚠️ Carregue os extratos primeiro.", fg="#fbbf24")
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Histórico de Fluxo"

            borda = Border(
                left=Side(style='thin', color='CCCCCC'),
                right=Side(style='thin', color='CCCCCC'),
                top=Side(style='thin', color='CCCCCC'),
                bottom=Side(style='thin', color='CCCCCC')
            )

            headers = ["Período", "Conta", "Empresa",
                       "Saldo Anterior", "Créditos", "Débitos", "Saldo Final", "Variação"]
            larguras = [12, 28, 12, 16, 16, 16, 16, 16]

            ws.merge_cells("A1:H1")
            c = ws.cell(row=1, column=1,
                        value=f"HISTÓRICO DE FLUXO DE CAIXA — gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
            c.font = Font(bold=True, color="FFFFFF", size=12)
            c.fill = PatternFill("solid", fgColor="1A1A2E")
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 26

            for col, (h, w) in enumerate(zip(headers, larguras), 1):
                cell = ws.cell(row=2, column=col, value=h)
                cell.font = Font(bold=True, color="FFFFFF", size=10)
                cell.fill = PatternFill("solid", fgColor="2D3748")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = borda
                ws.column_dimensions[get_column_letter(col)].width = w

            fmt_r = 'R$ #,##0.00'
            for linha, d in enumerate(sorted(self._fluxo_dados,
                                             key=lambda x: (x["periodo"], x["conta"])), 3):
                var = d["variacao"]
                cor_bg = "D4EDDA" if var >= 0 else "FDECEA"
                vals = [d["periodo"], d["conta"], d["empresa"],
                        d["saldo_ant"], d["creditos"], d["debitos"], d["saldo_fin"], var]
                for col, v in enumerate(vals, 1):
                    cell = ws.cell(row=linha, column=col, value=v)
                    cell.border = borda
                    cell.font = Font(size=9)
                    if col >= 4:
                        cell.number_format = fmt_r
                        cell.alignment = Alignment(horizontal="right")
                    if col == 8:
                        cell.fill = PatternFill("solid", fgColor=cor_bg[2:] if var >= 0 else "FDECEA")

            ws.freeze_panes = "A3"
            ws.auto_filter.ref = f"A2:H{len(self._fluxo_dados) + 2}"

            nome_saida = os.path.join(PASTA_ATUAL,
                                       f"Historico_Fluxo_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
            wb.save(nome_saida)
            os.startfile(nome_saida)
            self._fluxo_lbl_status.config(text=f"✅ Exportado: {os.path.basename(nome_saida)}", fg="#4ade80")

        except Exception as e:
            self._fluxo_lbl_status.config(text=f"❌ Erro ao exportar: {e}", fg="#f87171")
