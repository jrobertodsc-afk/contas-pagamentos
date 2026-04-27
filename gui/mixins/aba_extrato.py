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


class AbaExtratoMixin:

    def _build_aba_extrato(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        # ── TÍTULO ──
        ctk.CTkLabel(parent, text="📊 FLUXO DE CAIXA CONSOLIDADO", font=("Segoe UI", 16, "bold"), text_color=accent).pack(anchor="w", padx=25, pady=(20, 5))
        ctk.CTkLabel(parent, text="Integração de extratos bancários (Itaú) e adquirentes (Getnet) para geração de dashboard.", font=("Segoe UI", 11), text_color=muted).pack(anchor="w", padx=25, pady=(0, 15))

        # ── SEÇÃO ITAÚ ──
        frame_itau = ctk.CTkFrame(parent, fg_color=surface, corner_radius=12, border_width=1, border_color=border)
        frame_itau.pack(fill="x", padx=25, pady=(0, 10))
        
        ctk.CTkLabel(frame_itau, text="🏦 EXTRATO ITAÚ (PDF)", font=("Segoe UI", 10, "bold"), text_color=accent).pack(anchor="w", padx=15, pady=(10, 5))

        row_itau = ctk.CTkFrame(frame_itau, fg_color="transparent")
        row_itau.pack(fill="x", padx=15, pady=(0, 15))

        self.extrato_caminho = tk.StringVar()
        self.extrato_entry = ctk.CTkEntry(row_itau, textvariable=self.extrato_caminho, font=("Consolas", 11), fg_color=bg, border_color=border, placeholder_text="Selecione o PDF do extrato...", width=600, height=35)
        self.extrato_entry.pack(side="left", padx=(0, 10))
        
        ctk.CTkButton(row_itau, text="📂 ABRIR PDF", font=("Segoe UI", 10, "bold"), fg_color=border, text_color=text, width=120, height=35, command=self._selecionar_extrato).pack(side="left")

        # ── SEÇÃO GETNET ──
        frame_getnet = ctk.CTkFrame(parent, fg_color=surface, corner_radius=12, border_width=1, border_color=border)
        frame_getnet.pack(fill="x", padx=25, pady=(0, 10))
        
        ctk.CTkLabel(frame_getnet, text="💳 ADQUIRENTES GETNET (XLSX)", font=("Segoe UI", 10, "bold"), text_color=yellow).pack(anchor="w", padx=15, pady=(10, 5))

        self._getnet_paths = []
        self._getnet_container = ctk.CTkFrame(frame_getnet, fg_color="transparent")
        self._getnet_container.pack(fill="x", padx=15, pady=(0, 5))

        def _add_getnet_linha(path=""):
            idx = len(self._getnet_paths)
            var = tk.StringVar(value=path)
            self._getnet_paths.append(var)
            
            row_f = ctk.CTkFrame(self._getnet_container, fg_color="transparent")
            row_f.pack(fill="x", pady=2)
            
            ctk.CTkLabel(row_f, text=f"FILIAL {idx+1}:", font=("Segoe UI", 9, "bold"), text_color=muted, width=70).pack(side="left")
            ctk.CTkEntry(row_f, textvariable=var, font=("Consolas", 11), fg_color=bg, border_color=border, width=530, height=32).pack(side="left", padx=10)
            ctk.CTkButton(row_f, text="📂", font=("Segoe UI", 10, "bold"), fg_color=border, text_color=yellow, width=40, height=32, command=lambda v=var: _sel_getnet(v)).pack(side="left")

        def _sel_getnet(var):
            from tkinter import filedialog
            p = filedialog.askopenfilename(title="Selecionar extrato Getnet", filetypes=[("Excel", "*.xlsx *.xls")])
            if p: var.set(p)

        _add_getnet_linha() 

        ctk.CTkButton(frame_getnet, text="+ ADICIONAR FILIAL", font=("Segoe UI", 9, "bold"), fg_color="transparent", text_color=yellow, border_width=1, border_color=yellow, hover_color="#2d2d00", width=150, height=28, command=_add_getnet_linha).pack(anchor="w", padx=15, pady=(0, 15))

        # ── BOTÕES DE AÇÃO ──
        frame_bts = ctk.CTkFrame(parent, fg_color="transparent")
        frame_bts.pack(fill="x", padx=25, pady=5)

        ctk.CTkButton(frame_bts, text="⚡ GERAR FLUXO CONSOLIDADO", font=("Segoe UI", 12, "bold"), fg_color="#7c3aed", hover_color="#6d28d9", text_color="white", height=45, width=250, command=self._processar_fluxo_consolidado).pack(side="left", padx=(0, 15))
        ctk.CTkButton(frame_bts, text="📊 APENAS ITAÚ", font=("Segoe UI", 11), fg_color=surface, border_width=1, border_color=border, text_color=accent, height=45, width=150, command=self._processar_extrato).pack(side="left")

        # ── ENVIO POR EMAIL ──
        frame_email = ctk.CTkFrame(parent, fg_color=surface, corner_radius=12)
        frame_email.pack(fill="x", padx=25, pady=10)
        
        ctk.CTkLabel(frame_email, text="E-MAIL PARA DASHBOARD:", font=("Segoe UI", 9, "bold"), text_color=muted).pack(side="left", padx=(15, 10))
        self.extrato_email = ctk.CTkEntry(frame_email, font=("Segoe UI", 11), fg_color=bg, border_color=border, width=300, height=35)
        self.extrato_email.insert(0, EMAIL_DESTINO_EXTRATO)
        self.extrato_email.pack(side="left", pady=10)
        
        ctk.CTkButton(frame_email, text="📤 ENVIAR POR E-MAIL", font=("Segoe UI", 10, "bold"), fg_color="#1e3a2f", text_color="#4ade80", height=35, width=180, command=self._enviar_dashboard_email).pack(side="left", padx=15)

        # ── STATUS ──
        self.extrato_status = ctk.CTkLabel(parent, text="", font=("Segoe UI", 11, "bold"), text_color=muted)
        self.extrato_status.pack(anchor="w", padx=25)

        # ── PREVIEW (Treeview estilizada) ──
        tree_container = ctk.CTkFrame(parent, fg_color=bg, corner_radius=12, border_width=1, border_color=border)
        tree_container.pack(fill="both", expand=True, padx=25, pady=(5, 15))

        cols = ("Origem", "Data", "Descrição / Favorecido", "Tipo", "Entrada", "Saída")
        self.extrato_tree = ttk.Treeview(tree_container, columns=cols, show="headings")
        
        style = ttk.Style()
        style.configure("Extrato.Treeview", background=bg, foreground=text, fieldbackground=bg, rowheight=30, borderwidth=0, font=("Segoe UI", 9))
        style.configure("Extrato.Treeview.Heading", background=surface, foreground=accent, relief="flat", font=("Segoe UI", 9, "bold"))
        
        self.extrato_tree.configure(style="Extrato.Treeview")

        larguras = [120, 90, 350, 160, 110, 110]
        for col, w in zip(cols, larguras):
            self.extrato_tree.heading(col, text=col.upper())
            self.extrato_tree.column(col, width=w, anchor="e" if col in ("Entrada","Saída") else "w")

        self.extrato_tree.tag_configure("deb",  background="#2d1515", foreground="#f87171")
        self.extrato_tree.tag_configure("cred", background="#152d1f", foreground="#4ade80")
        self.extrato_tree.tag_configure("getnet",background="#1a1510", foreground="#fbbf24")

        scroll_ext = ctk.CTkScrollbar(tree_container, orientation="vertical", command=self.extrato_tree.yview)
        self.extrato_tree.configure(yscrollcommand=scroll_ext.set)
        
        self.extrato_tree.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scroll_ext.pack(side="right", fill="y", padx=(0, 5), pady=5)


    def _selecionar_extrato(self):
        from tkinter import filedialog
        caminho = filedialog.askopenfilename(
            title="Selecionar extrato Itaú",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if caminho:
            self.extrato_caminho.set(caminho)


    def _processar_extrato(self):
        caminho = self.extrato_caminho.get().strip()
        if not caminho or not os.path.exists(caminho):
            self.extrato_status.config(text="⚠️ Selecione um arquivo PDF válido.", fg="#fbbf24")
            return

        self.extrato_status.config(text="⏳ Lendo extrato...", fg="#38bdf8")
        self.root.update_idletasks()

        def processar():
            try:
                contas, periodo = ler_extrato_pdf(caminho)
                if not contas:
                    self.root.after(0, lambda: self.extrato_status.config(
                        text="⚠️ Nenhuma conta encontrada no PDF.", fg="#fbbf24"))
                    return

                # Atualiza preview
                for row in self.extrato_tree.get_children():
                    self.extrato_tree.delete(row)

                total_lanc = 0
                for num_conta, dados in contas.items():
                    for l in dados['lancamentos']:
                        tag = "deb" if l['debito'] else "cred"
                        deb_str  = f"R$ {l['valor']:,.2f}" if l['debito']  else ""
                        cred_str = f"R$ {l['valor']:,.2f}" if not l['debito'] else ""
                        self.extrato_tree.insert("", "end", tags=(tag,), values=(
                            dados['nome'], l['data'], l['descricao'],
                            l['tipo'], deb_str, cred_str
                        ))
                        total_lanc += 1

                # Salva na pasta local EXTRATOS_PROCESSADOS
                os.makedirs(PASTA_EXTRATOS, exist_ok=True)
                nome_base = os.path.splitext(os.path.basename(caminho))[0]
                nome_excel = f"Dashboard_{nome_base}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                caminho_excel = os.path.join(PASTA_EXTRATOS, nome_excel)

                ok, resultado = gerar_dashboard_contas_pagar(contas, periodo, caminho_excel)

                if ok:
                    # Backup na rede
                    try:
                        import subprocess
                        os.makedirs(PASTA_BACKUP_EXTRATOS, exist_ok=True) if not os.path.exists(PASTA_BACKUP_EXTRATOS) else None
                        subprocess.run(
                            ["robocopy", PASTA_EXTRATOS, PASTA_BACKUP_EXTRATOS, "/XO", "/NP", "/LOG+:NUL"],
                            capture_output=True, timeout=60
                        )
                        backup_ok = " | ✅ Backup rede OK"
                    except: backup_ok = " | ⚠️ Backup rede falhou"

                    msg = f"✅ Dashboard gerado! {total_lanc} lançamentos | {len(contas)} conta(s){backup_ok}"
                    cor = "#4ade80"
                    os.startfile(resultado)
                else:
                    msg = f"❌ Erro ao gerar dashboard: {resultado}"
                    cor = "#f87171"

                self.root.after(0, lambda: self.extrato_status.config(text=msg, fg=cor))

            except Exception as e:
                self.root.after(0, lambda: self.extrato_status.config(
                    text=f"❌ Erro: {e}", fg="#f87171"))

        threading.Thread(target=processar, daemon=True).start()


    def _processar_fluxo_consolidado(self):
        """Lê extrato Itaú (opcional) + arquivos Getnet e gera fluxo de caixa consolidado."""
        paths_getnet = [v.get().strip() for v in self._getnet_paths if v.get().strip()]
        path_itau    = self.extrato_caminho.get().strip()

        if not paths_getnet and not path_itau:
            self.extrato_status.config(
                text="⚠️ Selecione ao menos um arquivo (extrato Itaú ou Getnet).", fg="#fbbf24")
            return

        self.extrato_status.config(text="⏳ Processando...", fg="#38bdf8")
        self.root.update_idletasks()

        def processar():
            try:
                # Lê Getnet
                recebiveis = []
                for path in paths_getnet:
                    if os.path.exists(path):
                        itens = ler_getnet_excel(path)
                        recebiveis.extend(itens)

                # Lê Itaú
                contas  = {}
                periodo = datetime.now().strftime("%d/%m/%Y")
                if path_itau and os.path.exists(path_itau):
                    contas, periodo = ler_extrato_pdf(path_itau)

                # Preview na treeview
                for row in self.extrato_tree.get_children():
                    self.extrato_tree.delete(row)

                # Mostra Getnet no preview
                for r in sorted(recebiveis, key=lambda x: x["data_venda"] or __import__('datetime').date.min):
                    self.extrato_tree.insert("", "end", tags=("getnet",), values=(
                        f"Getnet/{r['filial']}",
                        r["data_venda"].strftime("%d/%m/%Y") if r["data_venda"] else "",
                        f"{r['bandeira']} {r['forma']}",
                        r["modalidade"],
                        f"R$ {r['liquido']:,.2f}",
                        "",
                    ))

                # Mostra Itaú no preview
                for num_conta, dados in contas.items():
                    for l in dados.get("lancamentos", []):
                        tag = "deb" if l["debito"] else "cred"
                        self.extrato_tree.insert("", "end", tags=(tag,), values=(
                            dados["nome"], l["data"], l["descricao"], l["tipo"],
                            f"R$ {l['valor']:,.2f}" if not l["debito"] else "",
                            f"R$ {l['valor']:,.2f}" if l["debito"] else "",
                        ))

                # Gera Excel
                os.makedirs(PASTA_EXTRATOS, exist_ok=True)
                nome_arq = f"FluxoCaixa_Consolidado_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                caminho_excel = os.path.join(PASTA_EXTRATOS, nome_arq)

                ok, resultado = gerar_fluxo_consolidado(contas, recebiveis, periodo, caminho_excel)

                n_getnet = len(recebiveis)
                n_itau   = sum(len(d.get("lancamentos",[])) for d in contas.values())
                msg = (f"✅ Fluxo gerado! {n_getnet} transações Getnet | "
                       f"{n_itau} lançamentos Itaú → {nome_arq}")
                self.root.after(0, lambda: self.extrato_status.config(text=msg, fg="#4ade80"))
                if ok:
                    os.startfile(resultado)

            except Exception as e:
                import traceback
                self.root.after(0, lambda: self.extrato_status.config(
                    text=f"❌ Erro: {e}", fg="#f87171"))

        threading.Thread(target=processar, daemon=True).start()


    def _enviar_dashboard_email(self):
        """Envia o último dashboard gerado por e-mail."""
        destino = self.extrato_email.get().strip()
        if not destino:
            self.extrato_status.config(text="⚠️ Informe o e-mail de destino.", fg="#fbbf24")
            return

        # Busca o Excel mais recente na pasta de extratos
        os.makedirs(PASTA_EXTRATOS, exist_ok=True)
        arquivos = [f for f in os.listdir(PASTA_EXTRATOS) if f.endswith('.xlsx')]
        if not arquivos:
            self.extrato_status.config(text="⚠️ Nenhum dashboard gerado ainda. Processe o extrato primeiro.", fg="#fbbf24")
            return

        ultimo = max(arquivos, key=lambda f: os.path.getmtime(os.path.join(PASTA_EXTRATOS, f)))
        caminho_excel = os.path.join(PASTA_EXTRATOS, ultimo)

        self.extrato_status.config(text=f"⏳ Enviando dashboard para {destino}...", fg="#38bdf8")
        self.root.update_idletasks()

        def enviar():
            try:
                msg = MIMEMultipart()
                msg['From']    = EMAIL_REMETENTE
                msg['To']      = destino
                msg['Subject'] = f"Dashboard Contas a Pagar — {datetime.now().strftime('%d/%m/%Y')}"

                corpo = f"""<html><body style='font-family:Arial,sans-serif;'>
                <div style='background:#1a1a2e;padding:20px;border-radius:8px;'>
                  <h2 style='color:#38bdf8;margin:0;'>📊 Dashboard Contas a Pagar</h2>
                  <p style='color:#94a3b8;'>Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}</p>
                </div>
                <p>Segue em anexo o dashboard diário de contas a pagar com:</p>
                <ul>
                  <li>Posição financeira consolidada de todas as contas</li>
                  <li>Saídas do dia por categoria (Folha, Fornecedores, Impostos, etc.)</li>
                  <li>Top 10 maiores pagamentos</li>
                  <li>Recebimentos por filial e status de transferência para Matriz</li>
                  <li>Alertas de saldo crítico</li>
                </ul>
                <p style='font-size:11px;color:#888;'><i>Robô Financeiro BOAH/SOLAR — Contas a Pagar</i></p>
                </body></html>"""

                msg.attach(MIMEText(corpo, 'html', 'utf-8'))

                # Copia para arquivo temporário para evitar erro se Excel estiver aberto
                import tempfile
                with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                    tmp_path = tmp.name
                shutil.copy2(caminho_excel, tmp_path)
                try:
                    with open(tmp_path, 'rb') as f:
                        parte = MIMEApplication(f.read(), _subtype='xlsx')
                        parte.add_header('Content-Disposition', 'attachment', filename=ultimo)
                        msg.attach(parte)
                finally:
                    try: os.remove(tmp_path)
                    except: pass

                with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as srv:
                    srv.login(EMAIL_REMETENTE, EMAIL_SENHA)
                    srv.sendmail(EMAIL_REMETENTE, destino, msg.as_bytes())

                self.root.after(0, lambda: self.extrato_status.config(
                    text=f"✅ Dashboard enviado para {destino}!", fg="#4ade80"))
            except Exception as e:
                self.root.after(0, lambda: self.extrato_status.config(
                    text=f"❌ Erro ao enviar: {e}", fg="#f87171"))

        threading.Thread(target=enviar, daemon=True).start()
