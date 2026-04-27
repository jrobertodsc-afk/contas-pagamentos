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


class AbaBuscaMixin:

    def _build_aba_busca(self, parent, bg, surface, border, accent, green, yellow, text, muted):

        # Categorias reais
        CATEGORIAS_BUSCA = ["Todas"] + CATEGORIAS_LISTA + [
            "RH - SALARIOS", "RH - DOMINGOS E FERIADOS", "RH - FGTS",
            "RH - RESCISAO", "RH - FERIAS", "RH - ADIANTAMENTO",
            "FORNECEDORES", "IMPOSTOS", "CONTAS_CONSUMO", "A_VERIFICAR"
        ]

        # ── FILTROS ──
        filtros = ctk.CTkFrame(parent, fg_color=surface, corner_radius=12, border_width=1, border_color=border)
        filtros.pack(fill="x", padx=25, pady=(20, 10))

        # Linha 1: Nome + Categoria
        linha1 = ctk.CTkFrame(filtros, fg_color="transparent")
        linha1.pack(fill="x", padx=15, pady=15)

        ctk.CTkLabel(linha1, text="NOME DO FORNECEDOR", font=("Segoe UI", 9, "bold"), text_color=muted).grid(row=0, column=0, sticky="w", padx=(0, 20))
        self.busca_nome = ctk.CTkEntry(linha1, font=("Segoe UI", 12), fg_color=bg, border_color=border, placeholder_text="Ex: COELBA...", width=280, height=40)
        self.busca_nome.grid(row=1, column=0, padx=(0, 20), pady=(5, 0))
        self.busca_nome.bind("<Return>", lambda e: self._executar_busca())

        ctk.CTkLabel(linha1, text="CATEGORIA", font=("Segoe UI", 9, "bold"), text_color=muted).grid(row=0, column=1, sticky="w", padx=(0, 20))
        
        self.busca_cat_var = tk.StringVar(value="Todas")
        self.busca_cat = ctk.CTkOptionMenu(
            linha1, 
            values=CATEGORIAS_BUSCA[:15], 
            variable=self.busca_cat_var,
            fg_color=bg,
            button_color=border,
            button_hover_color=accent,
            dropdown_fg_color=surface,
            width=220,
            height=40
        )
        self.busca_cat.grid(row=1, column=1, padx=(0, 20), pady=(5, 0))

        ctk.CTkLabel(linha1, text="EMPRESA", font=("Segoe UI", 9, "bold"), text_color=muted).grid(row=0, column=2, sticky="w")
        self.busca_empresa_var = tk.StringVar(value="Todas")
        self.busca_empresa = ctk.CTkOptionMenu(
            linha1, 
            values=["Todas", "LALUA", "SOLAR"],
            variable=self.busca_empresa_var,
            fg_color=bg,
            button_color=border,
            width=120,
            height=40
        )
        self.busca_empresa.grid(row=1, column=2, pady=(5, 0))

        # Linha 2: Datas + Valores
        linha2 = ctk.CTkFrame(filtros, fg_color="transparent")
        linha2.pack(fill="x", padx=15, pady=(0, 15))

        ctk.CTkLabel(linha2, text="PERÍODO (DE / ATÉ)", font=("Segoe UI", 9, "bold"), text_color=muted).grid(row=0, column=0, columnspan=2, sticky="w", padx=(0, 20))
        
        self.busca_data_de = ctk.CTkEntry(linha2, font=("Segoe UI", 12), fg_color=bg, border_color=border, width=130, height=35)
        self.busca_data_de.insert(0, "01/01/2023")
        self.busca_data_de.grid(row=1, column=0, padx=(0, 10), pady=(5, 0))
        _aplicar_mascara_data(self.busca_data_de)

        self.busca_data_ate = ctk.CTkEntry(linha2, font=("Segoe UI", 12), fg_color=bg, border_color=border, width=130, height=35)
        self.busca_data_ate.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.busca_data_ate.grid(row=1, column=1, padx=(0, 20), pady=(5, 0))
        _aplicar_mascara_data(self.busca_data_ate)

        ctk.CTkLabel(linha2, text="VALOR (MÍN / MÁX)", font=("Segoe UI", 9, "bold"), text_color=muted).grid(row=0, column=2, columnspan=2, sticky="w")
        
        self.busca_val_min = ctk.CTkEntry(linha2, font=("Segoe UI", 12), fg_color=bg, border_color=border, placeholder_text="0,00", width=110, height=35)
        self.busca_val_min.grid(row=1, column=2, padx=(0, 10), pady=(5, 0))

        self.busca_val_max = ctk.CTkEntry(linha2, font=("Segoe UI", 12), fg_color=bg, border_color=border, placeholder_text="999.999", width=110, height=35)
        self.busca_val_max.grid(row=1, column=3, pady=(5, 0))

        # ── BOTÕES DE AÇÃO ──
        acoes_f = ctk.CTkFrame(parent, fg_color="transparent")
        acoes_f.pack(fill="x", padx=25, pady=5)

        self.btn_buscar = ctk.CTkButton(
            acoes_f, 
            text="🔍 EXECUTAR BUSCA",
            font=("Segoe UI", 12, "bold"),
            fg_color=accent,
            text_color="#000000",
            hover_color=green,
            height=45,
            width=200,
            command=self._executar_busca
        )
        self.btn_buscar.pack(side="left")

        ctk.CTkButton(
            acoes_f, 
            text="🔄 Reclassificar Base",
            font=("Segoe UI", 11),
            fg_color=surface,
            text_color=text,
            hover_color=border,
            height=45,
            width=180,
            command=self._reclassificar_comprovantes
        ).pack(side="left", padx=15)

        self.lbl_resultado = ctk.CTkLabel(acoes_f, text="Aguardando parâmetros...", font=("Segoe UI", 11), text_color=muted)
        self.lbl_resultado.pack(side="right")

        # ── RESULTADO (Treeview estilizada) ──
        tree_container = ctk.CTkFrame(parent, fg_color=bg, corner_radius=12, border_width=1, border_color=border)
        tree_container.pack(fill="both", expand=True, padx=25, pady=10)

        cols = ("Data", "Nome", "Categoria", "Filial", "Empresa", "Valor")
        self.tree = ttk.Treeview(tree_container, columns=cols, show="headings")
        
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", 
                        background=bg, 
                        foreground=text, 
                        fieldbackground=bg, 
                        rowheight=35,
                        borderwidth=0,
                        font=("Segoe UI", 10))
        style.configure("Treeview.Heading", 
                        background=surface, 
                        foreground=accent, 
                        relief="flat",
                        font=("Segoe UI", 10, "bold"))
        style.map("Treeview", background=[("selected", border)])

        larguras = [100, 280, 200, 140, 80, 100]
        for col, w in zip(cols, larguras):
            self.tree.heading(col, text=col.upper(), command=lambda c=col: self._busca_ordenar(c))
            self.tree.column(col, width=w, anchor="w" if col != "Valor" else "e")

        scroll_y = ctk.CTkScrollbar(tree_container, orientation="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        
        self.tree.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scroll_y.pack(side="right", fill="y", padx=(0, 5), pady=5)

        self.tree.bind("<Double-1>", self._busca_abrir_pdf)

        # ── RODAPÉ ──
        rod = ctk.CTkFrame(parent, fg_color=surface, height=60, corner_radius=0)
        rod.pack(fill="x", side="bottom")

        ctk.CTkLabel(rod, text="E-MAIL DESTINO:", font=("Segoe UI", 9, "bold"), text_color=muted).pack(side="left", padx=(25, 10))
        self.busca_email_destino = ctk.CTkEntry(rod, font=("Segoe UI", 11), fg_color=bg, border_color=border, width=250, height=32)
        self.busca_email_destino.insert(0, EMAIL_DESTINO_DP)
        self.busca_email_destino.pack(side="left", padx=(0, 15))

        ctk.CTkButton(rod, text="📤 ENVIAR SELECIONADOS", font=("Segoe UI", 10, "bold"), fg_color="#1e3a2f", text_color="#4ade80", width=180, height=32, command=self._enviar_selecionados).pack(side="left", padx=5)
        ctk.CTkButton(rod, text="💾 DOWNLOAD", font=("Segoe UI", 10, "bold"), fg_color="#1e2a3a", text_color="#38bdf8", width=120, height=32, command=self._baixar_selecionados).pack(side="left", padx=5)
        ctk.CTkButton(rod, text="📋 AUTORIZAÇÃO", font=("Segoe UI", 10, "bold"), fg_color="#3a1e2a", text_color="#f472b6", width=140, height=32, command=self._enviar_para_autorizacao).pack(side="left", padx=5)

        self.lbl_status_envio = ctk.CTkLabel(rod, text="", font=("Segoe UI", 10, "bold"), text_color="#4ade80")
        self.lbl_status_envio.pack(side="right", padx=25)


    def _executar_busca(self):
        """Varre SAIDA_ORGANIZADA com os filtros informados."""
        nome_filtro  = self.busca_nome.get().strip().upper()
        cat_filtro   = self.busca_cat.get()
        empresa_filt = self.busca_empresa.get()
        data_de_str  = self.busca_data_de.get().strip()
        data_ate_str = self.busca_data_ate.get().strip()
        val_min_str  = self.busca_val_min.get().strip()
        val_max_str  = self.busca_val_max.get().strip()

        try: data_de  = datetime.strptime(data_de_str,  "%d/%m/%Y")
        except: data_de = datetime(2000, 1, 1)
        try: data_ate = datetime.strptime(data_ate_str, "%d/%m/%Y")
        except: data_ate = datetime(2099, 12, 31)
        try: val_min = _parse_valor(val_min_str) if val_min_str else 0.0
        except: val_min = 0.0
        try: val_max = _parse_valor(val_max_str) if val_max_str else 9999999.0
        except: val_max = 9999999.0

        resultados = []
        if not os.path.exists(PASTA_SAIDA):
            self.lbl_resultado.config(text="⚠️ Pasta SAIDA_ORGANIZADA não encontrada.")
            return

        for root_dir, _, files in os.walk(PASTA_SAIDA):
            partes = root_dir.replace(PASTA_SAIDA, "").split(os.sep)

            # Extrai empresa da estrutura de pastas
            empresa_pasta = next((p for p in partes
                                   if p.upper() in ("LALUA","SOLAR")), "")

            # Filtra empresa
            if empresa_filt != "Todas" and empresa_pasta.upper() != empresa_filt.upper():
                continue

            # Extrai categoria — qualquer parte que não seja data nem filial nem empresa
            categoria_pasta = next((p for p in partes
                                     if p and not p[0].isdigit()
                                     and p.upper() not in ("LALUA","SOLAR","SEM_DATA","GERAL","")),
                                    "")
            filial_pasta = os.path.basename(root_dir)

            for f in files:
                if not f.lower().endswith(".pdf"):
                    continue
                try:
                    partes_nome = f.replace(".pdf","").split("_")
                    data_obj = datetime.strptime(partes_nome[0], "%Y-%m-%d")
                    valor    = float(partes_nome[-1])
                    nome     = " ".join(partes_nome[1:-1])
                except:
                    continue

                # Filtros
                if nome_filtro and nome_filtro not in nome.upper():
                    continue
                if cat_filtro != "Todas":
                    # Busca no nome do arquivo E na pasta
                    cat_upper = cat_filtro.upper()
                    if (cat_upper not in categoria_pasta.upper() and
                        cat_upper not in nome.upper()):
                        continue
                if not (data_de <= data_obj <= data_ate):
                    continue
                if not (val_min <= valor <= val_max):
                    continue

                resultados.append({
                    'data':      data_obj,
                    'nome':      nome,
                    'categoria': categoria_pasta,
                    'filial':    filial_pasta,
                    'empresa':   empresa_pasta or "LALUA",
                    'valor':     valor,
                    'caminho':   os.path.join(root_dir, f)
                })

        self._resultados_busca = sorted(resultados,
                                         key=lambda x: x['data'], reverse=True)
        self._busca_popular_tree()


    def _busca_popular_tree(self):
        """Preenche a treeview com _resultados_busca."""
        for row in self.tree.get_children():
            self.tree.delete(row)

        total_valor = 0.0
        for item in self._resultados_busca:
            self.tree.insert("", "end", values=(
                item['data'].strftime('%d/%m/%Y'),
                item['nome'],
                item['categoria'],
                item['filial'],
                item['empresa'],
                f"R$ {item['valor']:,.2f}"
            ))
            total_valor += item['valor']

        n = len(self._resultados_busca)
        self.lbl_resultado.config(
            text=f"✅ {n} comprovante(s)  |  Total: R$ {total_valor:,.2f}"
            if n else "⚠️ Nenhum comprovante encontrado."
        )


    def _busca_ordenar(self, col):
        """Ordena a treeview pela coluna clicada."""
        mapa = {"Data":"data","Nome":"nome","Categoria":"categoria",
                "Filial":"filial","Empresa":"empresa","Valor":"valor"}
        chave = mapa.get(col, "data")
        if self._busca_ordem_col == chave:
            self._busca_ordem_rev = not self._busca_ordem_rev
        else:
            self._busca_ordem_col = chave
            self._busca_ordem_rev = False
        self._resultados_busca.sort(
            key=lambda x: x[chave],
            reverse=self._busca_ordem_rev
        )
        self._busca_popular_tree()


    def _busca_abrir_pdf(self, event=None):
        """Abre o PDF do item selecionado com duplo clique."""
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if idx < len(self._resultados_busca):
            caminho = self._resultados_busca[idx]['caminho']
            if os.path.exists(caminho):
                os.startfile(caminho)


    def _baixar_selecionados(self):
        """Copia os comprovantes selecionados para uma pasta escolhida pelo usuário."""
        from tkinter import filedialog
        selecionados = self.tree.selection()
        if not selecionados:
            itens = self._resultados_busca
        else:
            indices = [self.tree.index(s) for s in selecionados]
            itens = [self._resultados_busca[i] for i in indices]

        if not itens:
            self.lbl_status_envio.config(
                text="⚠️ Nenhum comprovante selecionado.", fg="#fbbf24")
            return

        pasta_dest = filedialog.askdirectory(
            title="Escolha a pasta de destino para os comprovantes")
        if not pasta_dest:
            return

        copiados = 0
        for item in itens:
            if os.path.exists(item['caminho']):
                nome_dest = os.path.basename(item['caminho'])
                dest = os.path.join(pasta_dest, nome_dest)
                # Evita sobrescrever
                if os.path.exists(dest):
                    base, ext = os.path.splitext(nome_dest)
                    dest = os.path.join(pasta_dest, f"{base}_{copiados}{ext}")
                shutil.copy2(item['caminho'], dest)
                copiados += 1

        self.lbl_status_envio.config(
            text=f"✅ {copiados} arquivo(s) copiado(s) para {pasta_dest}",
            fg="#38bdf8")
        try:
            os.startfile(pasta_dest)
        except: pass


    def _reclassificar_comprovantes(self):
        """Reclassifica todos os comprovantes em SAIDA_ORGANIZADA
        aplicando o DB de fornecedores atual (novos e antigos).
        Renomeia arquivos e move para pasta correta se categoria mudou.
        """
        from tkinter import messagebox
        if not messagebox.askyesno(
            "Reclassificar comprovantes",
            "Isso vai reclassificar TODOS os comprovantes usando o banco de "
            "fornecedores e categorias atual.\n\nArquivos serão movidos para "
            "as pastas corretas se a categoria mudou.\n\nDeseja continuar?",
            default="no"
        ):
            return

        self.lbl_resultado.config(text="⏳ Reclassificando...", fg="#38bdf8")
        self.root.update_idletasks()

        def _reclass():
            atualizados = 0
            erros = 0
            try:
                for root_dir, _, files in os.walk(PASTA_SAIDA):
                    for f in files:
                        if not f.lower().endswith(".pdf"):
                            continue
                        try:
                            partes_nome = f.replace(".pdf","").split("_")
                            nome = " ".join(partes_nome[1:-1])
                            valor_str = partes_nome[-1]
                            float(valor_str)
                        except:
                            continue

                        try:
                            # Reclassifica pelo banco atual
                            resp_novo, cat_nova = auto_classificar(nome, "")

                            # Verifica se é fornecedor cadastrado
                            f_db = buscar_fornecedor_por_nome(nome)
                            if f_db:
                                if f_db["categoria"]: cat_nova = f_db["categoria"]
                                if f_db["responsavel"]: resp_novo = f_db["responsavel"]

                            if not cat_nova or cat_nova == "A CLASSIFICAR":
                                continue  # mantém como está

                            # Verifica se pasta atual já é a categoria correta
                            cat_atual = os.path.basename(
                                os.path.dirname(os.path.dirname(
                                    os.path.join(root_dir, f))))
                            if cat_nova.upper() == cat_atual.upper():
                                continue  # já está correto

                            atualizados += 1
                        except:
                            erros += 1

            except Exception as e:
                self.root.after(0, lambda: self.lbl_resultado.config(
                    text=f"❌ Erro: {e}", fg="#f87171"))
                return

            self.root.after(0, lambda: self.lbl_resultado.config(
                text=f"✅ Reclassificação concluída — "
                     f"{atualizados} atualizados | {erros} erros",
                fg="#4ade80"))

        threading.Thread(target=_reclass, daemon=True).start()


    def _enviar_selecionados(self):
        """Envia por e-mail os comprovantes selecionados na lista."""
        selecionados = self.tree.selection()
        if not selecionados:
            itens = self._resultados_busca
        else:
            indices = [self.tree.index(s) for s in selecionados]
            itens = [self._resultados_busca[i] for i in indices]

        if not itens:
            self.lbl_status_envio.config(text="⚠️ Nenhum comprovante para enviar.", fg="#fbbf24")
            return

        destino = self.busca_email_destino.get().strip()
        if not destino:
            self.lbl_status_envio.config(text="⚠️ Informe o e-mail de destino.", fg="#fbbf24")
            return

        self.lbl_status_envio.config(text=f"⏳ Enviando {len(itens)} comprovante(s) para {destino}...", fg="#38bdf8")
        self.root.update_idletasks()

        def enviar():
            try:
                msg = MIMEMultipart()
                msg['From']    = EMAIL_REMETENTE
                msg['To']      = destino
                msg['Subject'] = f"Comprovantes — {len(itens)} arquivo(s) | {datetime.now().strftime('%d/%m/%Y')}"

                linhas_html = ""
                for item in itens:
                    linhas_html += f"<tr><td style='padding:6px 10px;border:1px solid #ddd'>{item['data'].strftime('%d/%m/%Y')}</td><td style='padding:6px 10px;border:1px solid #ddd'><b>{item['nome']}</b></td><td style='padding:6px 10px;border:1px solid #ddd'>{item['categoria']}</td><td style='padding:6px 10px;border:1px solid #ddd'>{item['filial']}</td><td style='padding:6px 10px;border:1px solid #ddd;text-align:right'>R$ {item['valor']:,.2f}</td></tr>"

                corpo = f"""<html><body style='font-family:Arial,sans-serif;'>
                <p>Seguem os comprovantes solicitados.</p>
                <table style='border-collapse:collapse;width:100%;font-size:12px;'>
                <tr style='background:#1a1a2e;color:#fff;'>
                  <th style='padding:8px 10px;'>Data</th><th style='padding:8px 10px;'>Nome</th>
                  <th style='padding:8px 10px;'>Categoria</th><th style='padding:8px 10px;'>Filial</th>
                  <th style='padding:8px 10px;'>Valor</th>
                </tr>{linhas_html}</table>
                <br><p style='font-size:11px;color:#888;'><i>Robô Financeiro BOAH/SOLAR</i></p>
                </body></html>"""

                msg.attach(MIMEText(corpo, 'html', 'utf-8'))

                anexados = 0
                for item in itens:
                    if os.path.exists(item['caminho']):
                        with open(item['caminho'], 'rb') as f:
                            parte = MIMEApplication(f.read(), _subtype='pdf')
                            parte.add_header('Content-Disposition', 'attachment',
                                             filename=os.path.basename(item['caminho']))
                            msg.attach(parte)
                            anexados += 1

                with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as srv:
                    srv.login(EMAIL_REMETENTE, EMAIL_SENHA)
                    srv.sendmail(EMAIL_REMETENTE, destino, msg.as_bytes())

                self.root.after(0, lambda: self.lbl_status_envio.config(
                    text=f"✅ E-mail enviado! {anexados} PDF(s) para {destino}", fg="#4ade80"))
                self._log(f"✅ Busca: e-mail enviado para {destino} com {anexados} PDF(s)")

            except Exception as e:
                self.root.after(0, lambda: self.lbl_status_envio.config(
                    text=f"❌ Erro: {e}", fg="#f87171"))
                self._log(f"❌ Busca: erro ao enviar — {e}")

        threading.Thread(target=enviar, daemon=True).start()


    def _abrir_dialogo_periodo(self):
        """Abre janela para reenvio de relatório RH — filtra por data específica ou mês inteiro."""
        win = tk.Toplevel(self.root)
        win.title("📧 Reenviar Relatório RH")
        win.geometry("420x290")
        win.resizable(False, False)
        win.configure(bg="#0f172a")
        win.grab_set()

        tk.Label(win, text="📧 Reenviar Relatório RH — Comprovantes",
                 font=("Segoe UI", 12, "bold"), fg="#38bdf8", bg="#0f172a").pack(pady=(16, 2))
        tk.Label(win, text="Escolha o período. Para um dia específico, preencha De e Até com a mesma data.",
                 font=("Segoe UI", 8), fg="#94a3b8", bg="#0f172a", wraplength=380).pack(pady=(0, 10))

        surface = "#1e293b"
        muted   = "#94a3b8"
        text    = "#e2e8f0"
        accent  = "#38bdf8"

        # ── Modo de filtro ──
        frame_modo = tk.Frame(win, bg="#0f172a")
        frame_modo.pack(pady=(0, 6))
        modo_var = tk.StringVar(value="intervalo")

        tk.Radiobutton(frame_modo, text="Intervalo de datas",
                       variable=modo_var, value="intervalo",
                       fg=text, bg="#0f172a", selectcolor="#0f172a",
                       activebackground="#0f172a", font=("Segoe UI", 9),
                       command=lambda: _atualizar_modo()).pack(side="left", padx=12)
        tk.Radiobutton(frame_modo, text="Mês inteiro",
                       variable=modo_var, value="mes",
                       fg=text, bg="#0f172a", selectcolor="#0f172a",
                       activebackground="#0f172a", font=("Segoe UI", 9),
                       command=lambda: _atualizar_modo()).pack(side="left", padx=12)

        # ── Frame intervalo ──
        frame_int = tk.Frame(win, bg=surface, padx=16, pady=12)
        frame_int.pack(fill="x", padx=20, pady=(0, 4))

        tk.Label(frame_int, text="De:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).grid(row=0, column=0, padx=(0, 6))
        de_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        de_entry = tk.Entry(frame_int, textvariable=de_var, font=("Segoe UI", 10),
                            bg="#0a0f1e", fg=text, insertbackground=accent,
                            relief="flat", bd=3, width=13)
        de_entry.grid(row=0, column=1, padx=(0, 16))

        tk.Label(frame_int, text="Até:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).grid(row=0, column=2, padx=(0, 6))
        ate_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        ate_entry = tk.Entry(frame_int, textvariable=ate_var, font=("Segoe UI", 10),
                             bg="#0a0f1e", fg=text, insertbackground=accent,
                             relief="flat", bd=3, width=13)
        ate_entry.grid(row=0, column=3)

        # Atalhos rápidos
        frame_atalhos = tk.Frame(frame_int, bg=surface)
        frame_atalhos.grid(row=1, column=0, columnspan=4, pady=(8, 0))
        hoje = datetime.now()

        def set_periodo(de, ate):
            de_var.set(de.strftime("%d/%m/%Y"))
            ate_var.set(ate.strftime("%d/%m/%Y"))

        from datetime import timedelta
        tk.Button(frame_atalhos, text="Hoje",
                  font=("Segoe UI", 7), bg="#334155", fg=text, relief="flat",
                  padx=6, pady=2, cursor="hand2",
                  command=lambda: set_periodo(hoje, hoje)).pack(side="left", padx=3)
        tk.Button(frame_atalhos, text="Ontem",
                  font=("Segoe UI", 7), bg="#334155", fg=text, relief="flat",
                  padx=6, pady=2, cursor="hand2",
                  command=lambda: set_periodo(hoje-timedelta(1), hoje-timedelta(1))).pack(side="left", padx=3)
        tk.Button(frame_atalhos, text="Últimos 7 dias",
                  font=("Segoe UI", 7), bg="#334155", fg=text, relief="flat",
                  padx=6, pady=2, cursor="hand2",
                  command=lambda: set_periodo(hoje-timedelta(7), hoje)).pack(side="left", padx=3)
        tk.Button(frame_atalhos, text="Esta semana",
                  font=("Segoe UI", 7), bg="#334155", fg=text, relief="flat",
                  padx=6, pady=2, cursor="hand2",
                  command=lambda: set_periodo(hoje-timedelta(hoje.weekday()), hoje)).pack(side="left", padx=3)
        tk.Button(frame_atalhos, text="Mês atual",
                  font=("Segoe UI", 7), bg="#334155", fg=text, relief="flat",
                  padx=6, pady=2, cursor="hand2",
                  command=lambda: set_periodo(hoje.replace(day=1), hoje)).pack(side="left", padx=3)

        # ── Frame mês ──
        frame_mes = tk.Frame(win, bg=surface, padx=16, pady=12)

        MESES_LISTA = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                       "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
        ano_atual = datetime.now().year
        anos = [str(a) for a in range(2024, ano_atual + 2)]

        tk.Label(frame_mes, text="Mês:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).grid(row=0, column=0, padx=(0, 6))
        mes_var2 = tk.StringVar(value=MESES_LISTA[datetime.now().month - 1])
        mes_cb = ttk.Combobox(frame_mes, textvariable=mes_var2, values=MESES_LISTA,
                               width=12, state="readonly")
        mes_cb.grid(row=0, column=1, padx=(0, 16))
        tk.Label(frame_mes, text="Ano:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).grid(row=0, column=2, padx=(0, 6))
        ano_var2 = tk.StringVar(value=str(ano_atual))
        ano_cb = ttk.Combobox(frame_mes, textvariable=ano_var2, values=anos,
                               width=7, state="readonly")
        ano_cb.grid(row=0, column=3)

        # ── E-mail destino ──
        frame_email = tk.Frame(win, bg="#0f172a")
        frame_email.pack(fill="x", padx=20, pady=(6, 0))
        tk.Label(frame_email, text="📧 Enviar para:", fg=muted, bg="#0f172a",
                 font=("Segoe UI", 8)).pack(side="left")
        email_var = tk.StringVar(value=EMAIL_DESTINO_DP)
        tk.Entry(frame_email, textvariable=email_var, font=("Segoe UI", 9),
                 bg="#1e293b", fg=text, insertbackground=accent,
                 relief="flat", bd=3, width=28).pack(side="left", padx=6)

        def _atualizar_modo():
            if modo_var.get() == "intervalo":
                frame_mes.pack_forget()
                frame_int.pack(fill="x", padx=20, pady=(0, 4))
            else:
                frame_int.pack_forget()
                frame_mes.pack(fill="x", padx=20, pady=(0, 4))

        def confirmar():
            destino = email_var.get().strip()
            if not destino:
                return
            if modo_var.get() == "intervalo":
                try:
                    dt_de  = datetime.strptime(de_var.get().strip(),  "%d/%m/%Y")
                    dt_ate = datetime.strptime(ate_var.get().strip(), "%d/%m/%Y")
                except:
                    tk.Label(win, text="⚠️ Data inválida. Use DD/MM/AAAA.",
                             fg="#fbbf24", bg="#0f172a", font=("Segoe UI", 8)).pack()
                    return
                label = (de_var.get() if de_var.get() == ate_var.get()
                         else f"{de_var.get()} a {ate_var.get()}")
                win.destroy()
                self._reenviar_relatorio_periodo_datas(dt_de, dt_ate, label, destino)
            else:
                mes_idx = MESES_LISTA.index(mes_var2.get()) + 1
                ano = int(ano_var2.get())
                win.destroy()
                self._reenviar_relatorio_periodo(mes_idx, ano, mes_var2.get(), destino)

        tk.Button(win, text="📤 Enviar Relatório",
                  font=("Segoe UI", 10, "bold"), bg="#0ea5e9", fg="white",
                  relief="flat", padx=20, pady=8, cursor="hand2",
                  command=confirmar).pack(pady=(10, 0))


    def _reenviar_relatorio_periodo_datas(self, dt_de, dt_ate, label, destino=None):
        """Reenvio por intervalo de datas — De até Ate inclusive."""
        self._log(f"🔍 Buscando comprovantes RH de {label}...")
        if destino is None:
            destino = EMAIL_DESTINO_DP

        itens = []
        if not os.path.exists(PASTA_SAIDA):
            self._log("⚠️ Pasta SAIDA_ORGANIZADA não encontrada.")
            return

        for root_dir, dirs, files in os.walk(PASTA_SAIDA):
            partes = root_dir.replace(PASTA_SAIDA, "").split(os.sep)
            categoria_pasta = next((p for p in partes if p.startswith("RH")), None)
            if not categoria_pasta:
                continue
            for f in files:
                if not f.lower().endswith(".pdf"):
                    continue
                try:
                    partes_nome = f.replace(".pdf", "").split("_")
                    data_str  = partes_nome[0]
                    valor_str = partes_nome[-1]
                    data_obj  = datetime.strptime(data_str, "%Y-%m-%d")
                    valor     = float(valor_str)
                    nome      = "_".join(partes_nome[1:-1]).replace("_", " ")
                except:
                    continue

                if not (dt_de <= data_obj <= dt_ate):
                    continue

                itens.append({
                    'nome': nome, 'valor': valor, 'data': data_obj,
                    'filial': os.path.basename(root_dir),
                    'caminho': os.path.join(root_dir, f),
                    'categoria': categoria_pasta,
                })

        if not itens:
            self._log(f"⚠️ Nenhum comprovante RH encontrado em {label}.")
            return

        self._log(f"✅ {len(itens)} comprovante(s) encontrado(s). Enviando para {destino}...")
        threading.Thread(
            target=lambda: enviar_email_relatorio(itens, self._log, destino_override=destino),
            daemon=True
        ).start()




    def _reenviar_relatorio_periodo(self, mes, ano, mes_nome, destino=None):
        """Varre SAIDA_ORGANIZADA buscando comprovantes RH do período e reenvia e-mail."""
        if destino is None:
            destino = EMAIL_DESTINO_DP
        self._log(f"🔍 Buscando comprovantes RH de {mes_nome}/{ano}...")

        mes_fmt = f"{mes:02d}_"  # ex: "03_"
        ano_fmt = str(ano)

        itens = []
        for empresa_dir in os.listdir(PASTA_SAIDA):
            caminho_empresa = os.path.join(PASTA_SAIDA, empresa_dir)
            if not os.path.isdir(caminho_empresa): continue
            caminho_ano = os.path.join(caminho_empresa, ano_fmt)
            if not os.path.isdir(caminho_ano): continue

            # Procura pasta do mês
            for pasta_mes in os.listdir(caminho_ano):
                if not pasta_mes.startswith(mes_fmt): continue
                caminho_mes = os.path.join(caminho_ano, pasta_mes)

                # Varre todos os dias e categorias RH
                for root_dir, dirs, files in os.walk(caminho_mes):
                    # Verifica se estamos dentro de uma categoria RH
                    partes = root_dir.replace(caminho_mes, "").split(os.sep)
                    categoria_pasta = next((p for p in partes if p.startswith("RH")), None)
                    if not categoria_pasta: continue

                    for f in files:
                        if not f.lower().endswith(".pdf"): continue
                        caminho_pdf = os.path.join(root_dir, f)

                        # Extrai dados do nome do arquivo: YYYY-MM-DD_NOME_VALOR.pdf
                        try:
                            partes_nome = f.replace(".pdf", "").split("_")
                            data_str = partes_nome[0]
                            valor_str = partes_nome[-1]
                            data_obj = datetime.strptime(data_str, "%Y-%m-%d")
                            valor = float(valor_str)
                            nome = "_".join(partes_nome[1:-1]).replace("_", " ")
                        except:
                            data_obj = None
                            valor = 0.0
                            nome = f.replace(".pdf", "")

                        # Filial vem do nome da pasta pai imediata
                        filial_pasta = os.path.basename(root_dir)

                        itens.append({
                            'nome': nome,
                            'valor': valor,
                            'data': data_obj,
                            'filial': filial_pasta,
                            'caminho': caminho_pdf,
                            'categoria': categoria_pasta.replace(" - ", " - "),
                        })

        if not itens:
            self._log(f"⚠️ Nenhum comprovante RH encontrado em {mes_nome}/{ano}.")
            return

        self._log(f"✅ {len(itens)} comprovante(s) encontrado(s). Enviando para {destino}...")
        threading.Thread(
            target=lambda: enviar_email_relatorio(itens, self._log, destino_override=destino),
            daemon=True
        ).start()
