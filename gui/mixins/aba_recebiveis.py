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


class AbaRecebiveisMixin:

    # ==========================================================================
    # ── ABA 6: RECEBÍVEIS / INTEGRAÇÕES ───────────────────────────────────
    # ==========================================================================

    def _build_aba_recebiveis(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        """Aba com integrações: PagarMe, Rede, Getnet e Linx Microvix."""

        # ── TÍTULO ──
        hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.pack(fill="x", padx=25, pady=(20, 10))
        
        ctk.CTkLabel(hdr, text="💳 RECEBÍVEIS & INTEGRAÇÕES", font=("Segoe UI", 16, "bold"), text_color=accent).pack(side="left")
        ctk.CTkLabel(parent, text="Consulte agenda de recebíveis e vendas em tempo real. Configure as credenciais nas CONFIGURAÇÕES.", font=("Segoe UI", 11), text_color=muted).pack(anchor="w", padx=25, pady=(0, 15))

        # ── TABVIEW INTERNO ──
        self.receb_tabs = ctk.CTkTabview(
            parent, 
            fg_color="transparent",
            segmented_button_selected_color=accent,
            segmented_button_selected_hover_color=green,
            segmented_button_unselected_color=surface,
            text_color="#000000" if ctk.get_appearance_mode() == "Dark" else "#FFFFFF"
        )
        self.receb_tabs.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        aba_pm = self.receb_tabs.add("  💰 PAGARME  ")
        aba_rede = self.receb_tabs.add("  🔴 REDE  ")
        aba_getnet = self.receb_tabs.add("  🟢 GETNET  ")
        aba_mv = self.receb_tabs.add("  🏪 MICROVIX  ")

        self._build_sub_pagarme(aba_pm, bg, surface, accent, green, yellow, text, muted)
        self._build_sub_rede(aba_rede, bg, surface, accent, green, yellow, text, muted)
        self._build_sub_getnet(aba_getnet, bg, surface, accent, green, yellow, text, muted)
        self._build_sub_microvix(aba_mv, bg, surface, accent, green, yellow, text, muted)


    # ── HELPERS COMUNS ──────────────────────────────────────────────────────

    def _card_credencial(self, parent, bg, surface, muted, text, accent,
                         titulo, status_cfg, campos, var_dict):
        """Monta um bloco de configuração de credenciais reutilizável."""
        frame = tk.Frame(parent, bg=surface, padx=14, pady=10)
        frame.pack(fill="x", padx=0, pady=(0, 8))

        tk.Label(frame, text=titulo, font=("Segoe UI", 9, "bold"),
                 fg=accent, bg=surface).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))

        ok = all(v for v in status_cfg)
        status_txt = "✅ Credenciais configuradas" if ok else "⚠️ Credenciais não configuradas — edite as CONFIGURAÇÕES no topo do arquivo"
        status_cor = "#4ade80" if ok else "#fbbf24"
        tk.Label(frame, text=status_txt, font=("Consolas", 8),
                 fg=status_cor, bg=surface).grid(row=1, column=0, columnspan=4, sticky="w", pady=(0, 8))

        for i, (label, var_key, default, show) in enumerate(campos):
            row = 2 + i
            tk.Label(frame, text=label, fg=muted, bg=surface,
                     font=("Segoe UI", 8)).grid(row=row, column=0, sticky="w", padx=(0, 6))
            var = tk.StringVar(value=default)
            var_dict[var_key] = var
            e = tk.Entry(frame, textvariable=var, font=("Consolas", 9),
                         bg="#0a0f1e", fg=text, insertbackground=accent,
                         relief="flat", bd=3, width=38,
                         show="*" if show else "")
            e.grid(row=row, column=1, sticky="w", pady=2)

        return frame


    def _resultado_box(self, parent, bg, height=10):
        """Cria um ScrolledText para exibir resultados."""
        box = scrolledtext.ScrolledText(
            parent, height=height, font=("Consolas", 9),
            bg="#0a0f1e", fg="#a0c4ff", relief="flat", bd=0,
            padx=8, pady=8, state="disabled"
        )
        box.pack(fill="both", expand=True, pady=(4, 0))
        return box


    def _log_resultado(self, box, msg, cor="#a0c4ff"):
        """Escreve no box de resultado com timestamp."""
        box.config(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        box.insert("end", f"[{ts}] {msg}\n")
        box.see("end")
        box.config(state="disabled")


    def _limpar_resultado(self, box):
        box.config(state="normal")
        box.delete("1.0", "end")
        box.config(state="disabled")


    def _checar_requests(self):
        """Verifica se requests está instalado."""
        try:
            import requests
            return True
        except ImportError:
            return False


    # ── SUB-ABA PAGARME ─────────────────────────────────────────────────────

    def _build_sub_pagarme(self, parent, bg, surface, accent, green, yellow, text, muted):
        self._pm_vars = {}

        # Info
        info = tk.Frame(parent, bg=surface, padx=14, pady=10)
        info.pack(fill="x", padx=0, pady=(10, 8))
        tk.Label(info, text="💰  PagarMe — Agenda de Recebíveis",
                 font=("Segoe UI", 10, "bold"), fg=accent, bg=surface).pack(anchor="w")
        tk.Label(info,
                 text="Consulta recebíveis futuros (agenda financeira) e saldo disponível via API PagarMe v5.\n"
                      "Para obter a API Key: acesse dashboard.pagar.me → Configurações → API Keys → Criar chave de produção.",
                 font=("Segoe UI", 8), fg=muted, bg=surface, justify="left").pack(anchor="w", pady=(4, 0))

        # Credencial
        self._card_credencial(
            parent, bg, surface, muted, text, accent,
            titulo="🔑 Credencial PagarMe",
            status_cfg=[PAGARME_API_KEY],
            campos=[
                ("API Key:", "api_key", PAGARME_API_KEY, True),
            ],
            var_dict=self._pm_vars
        )

        # Filtros
        filtros = tk.Frame(parent, bg=surface, padx=14, pady=8)
        filtros.pack(fill="x", pady=(0, 8))
        tk.Label(filtros, text="Período — De:", fg=muted, bg=surface,
                 font=("Segoe UI", 8)).grid(row=0, column=0, padx=(0, 4))
        self._pm_de = tk.Entry(filtros, font=("Segoe UI", 9), bg="#0a0f1e", fg=text,
                                insertbackground=accent, relief="flat", bd=3, width=12)
        self._pm_de.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self._pm_de.grid(row=0, column=1, padx=(0, 12))

        tk.Label(filtros, text="Até:", fg=muted, bg=surface,
                 font=("Segoe UI", 8)).grid(row=0, column=2, padx=(0, 4))
        self._pm_ate = tk.Entry(filtros, font=("Segoe UI", 9), bg="#0a0f1e", fg=text,
                                 insertbackground=accent, relief="flat", bd=3, width=12)
        from datetime import timedelta
        self._pm_ate.insert(0, (datetime.now() + timedelta(days=30)).strftime("%d/%m/%Y"))
        self._pm_ate.grid(row=0, column=3, padx=(0, 16))

        tk.Button(filtros, text="🔍 Consultar Agenda",
                  font=("Segoe UI", 9, "bold"), bg="#7c3aed", fg="white",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=self._consultar_pagarme).grid(row=0, column=4, padx=(0, 8))

        tk.Button(filtros, text="💰 Ver Saldo Disponível",
                  font=("Segoe UI", 9, "bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=self._saldo_pagarme).grid(row=0, column=5)

        self._pm_resultado = self._resultado_box(parent, bg, height=12)


    def _consultar_pagarme(self):
        if not self._checar_requests():
            self._log_resultado(self._pm_resultado, "❌ Biblioteca 'requests' não instalada. Execute: pip install requests", "#f87171")
            return

        api_key = self._pm_vars.get("api_key", tk.StringVar()).get().strip()
        if not api_key:
            self._log_resultado(self._pm_resultado, "⚠️ Informe a API Key do PagarMe.", "#fbbf24")
            return

        try:
            de  = datetime.strptime(self._pm_de.get().strip(),  "%d/%m/%Y")
            ate = datetime.strptime(self._pm_ate.get().strip(), "%d/%m/%Y")
        except:
            self._log_resultado(self._pm_resultado, "⚠️ Data inválida. Use DD/MM/AAAA.", "#fbbf24")
            return

        self._limpar_resultado(self._pm_resultado)
        self._log_resultado(self._pm_resultado, f"⏳ Consultando agenda PagarMe de {self._pm_de.get()} a {self._pm_ate.get()}...")

        def consultar():
            try:
                import requests, base64
                token = base64.b64encode(f"{api_key}:".encode()).decode()
                headers = {"Authorization": f"Basic {token}", "Content-Type": "application/json"}

                # Consulta agenda de recebíveis
                params = {
                    "payment_date[gte]": de.strftime("%Y-%m-%dT00:00:00Z"),
                    "payment_date[lte]": ate.strftime("%Y-%m-%dT23:59:59Z"),
                    "count": 100,
                    "page": 1
                }
                resp = requests.get(f"{PAGARME_API_URL}/payables", headers=headers,
                                    params=params, timeout=15)

                if resp.status_code == 401:
                    self.root.after(0, lambda: self._log_resultado(
                        self._pm_resultado, "❌ API Key inválida ou sem permissão.", "#f87171"))
                    return
                if resp.status_code != 200:
                    self.root.after(0, lambda: self._log_resultado(
                        self._pm_resultado, f"❌ Erro {resp.status_code}: {resp.text[:200]}", "#f87171"))
                    return

                data = resp.json()
                itens = data.get("data", [])

                if not itens:
                    self.root.after(0, lambda: self._log_resultado(
                        self._pm_resultado, "ℹ️ Nenhum recebível encontrado no período.", "#fbbf24"))
                    return

                total = sum(i.get("amount", 0) for i in itens) / 100
                total_liquido = sum(i.get("net_amount", i.get("amount", 0)) for i in itens) / 100

                def atualizar():
                    self._log_resultado(self._pm_resultado,
                        f"✅ {len(itens)} recebível(is) encontrado(s)\n"
                        f"   Valor bruto:   R$ {total:>12,.2f}\n"
                        f"   Valor líquido: R$ {total_liquido:>12,.2f}\n"
                        f"{'─'*55}", "#4ade80")

                    # Agrupa por data de pagamento
                    por_data = {}
                    for i in itens:
                        dt = i.get("payment_date", "")[:10]
                        por_data.setdefault(dt, []).append(i)

                    for dt in sorted(por_data.keys()):
                        grupo = por_data[dt]
                        val_dia = sum(x.get("net_amount", x.get("amount", 0)) for x in grupo) / 100
                        try:
                            dt_fmt = datetime.strptime(dt, "%Y-%m-%d").strftime("%d/%m/%Y")
                        except:
                            dt_fmt = dt
                        self._log_resultado(self._pm_resultado,
                            f"  📅 {dt_fmt}  —  {len(grupo)} transação(ões)  —  R$ {val_dia:,.2f}")

                self.root.after(0, atualizar)

            except Exception as e:
                self.root.after(0, lambda: self._log_resultado(
                    self._pm_resultado, f"❌ Erro: {e}", "#f87171"))

        threading.Thread(target=consultar, daemon=True).start()


    def _saldo_pagarme(self):
        if not self._checar_requests():
            self._log_resultado(self._pm_resultado, "❌ 'requests' não instalado.", "#f87171")
            return

        api_key = self._pm_vars.get("api_key", tk.StringVar()).get().strip()
        if not api_key:
            self._log_resultado(self._pm_resultado, "⚠️ Informe a API Key.", "#fbbf24")
            return

        self._log_resultado(self._pm_resultado, "⏳ Consultando saldo PagarMe...")

        def consultar():
            try:
                import requests, base64
                token = base64.b64encode(f"{api_key}:".encode()).decode()
                headers = {"Authorization": f"Basic {token}"}
                resp = requests.get(f"{PAGARME_API_URL}/balance", headers=headers, timeout=15)

                if resp.status_code != 200:
                    self.root.after(0, lambda: self._log_resultado(
                        self._pm_resultado, f"❌ Erro {resp.status_code}: {resp.text[:200]}", "#f87171"))
                    return

                d = resp.json()
                disp    = d.get("available",   {}).get("amount", 0) / 100
                espera  = d.get("waiting_funds", {}).get("amount", 0) / 100
                transf  = d.get("transferred",  {}).get("amount", 0) / 100

                def atualizar():
                    self._log_resultado(self._pm_resultado,
                        f"💰 SALDO PAGARME\n"
                        f"   Disponível para saque: R$ {disp:>12,.2f}\n"
                        f"   Aguardando liquidação: R$ {espera:>12,.2f}\n"
                        f"   Total transferido:     R$ {transf:>12,.2f}", "#4ade80")

                self.root.after(0, atualizar)

            except Exception as e:
                self.root.after(0, lambda: self._log_resultado(
                    self._pm_resultado, f"❌ Erro: {e}", "#f87171"))

        threading.Thread(target=consultar, daemon=True).start()


    # ── SUB-ABA REDE ────────────────────────────────────────────────────────

    def _build_sub_rede(self, parent, bg, surface, accent, green, yellow, text, muted):
        self._rede_vars = {}

        info = tk.Frame(parent, bg=surface, padx=14, pady=10)
        info.pack(fill="x", padx=0, pady=(10, 8))
        tk.Label(info, text="🔴  Rede — Vendas & Recebíveis por Filial",
                 font=("Segoe UI", 10, "bold"), fg="#f87171", bg=surface).pack(anchor="w")
        tk.Label(info,
                 text="Consulta vendas do dia e agenda de recebíveis via API Rede (userede.com.br).\n"
                      "Credenciais: solicite em userede.com.br → Developers → Criar App.",
                 font=("Segoe UI", 8), fg=muted, bg=surface, justify="left").pack(anchor="w", pady=(4, 0))

        self._card_credencial(
            parent, bg, surface, muted, text, accent,
            titulo="🔑 Credenciais Rede",
            status_cfg=[REDE_CLIENT_ID, REDE_CLIENT_SECRET, REDE_PV],
            campos=[
                ("Client ID:", "client_id", REDE_CLIENT_ID, False),
                ("Client Secret:", "client_secret", REDE_CLIENT_SECRET, True),
                ("PV (estabelecimento):", "pv", REDE_PV, False),
            ],
            var_dict=self._rede_vars
        )

        # Filtros
        filtros = tk.Frame(parent, bg=surface, padx=14, pady=8)
        filtros.pack(fill="x", pady=(0, 8))
        tk.Label(filtros, text="Data:", fg=muted, bg=surface, font=("Segoe UI", 8)).grid(row=0, column=0, padx=(0, 4))
        self._rede_data = tk.Entry(filtros, font=("Segoe UI", 9), bg="#0a0f1e", fg=text,
                                    insertbackground=accent, relief="flat", bd=3, width=12)
        self._rede_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self._rede_data.grid(row=0, column=1, padx=(0, 12))

        tk.Button(filtros, text="📊 Consultar Vendas do Dia",
                  font=("Segoe UI", 9, "bold"), bg="#dc2626", fg="white",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=self._consultar_rede_vendas).grid(row=0, column=2, padx=(0, 8))

        tk.Button(filtros, text="📅 Agenda de Recebíveis",
                  font=("Segoe UI", 9, "bold"), bg="#7c2d12", fg="white",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=self._consultar_rede_recebiveis).grid(row=0, column=3)

        self._rede_resultado = self._resultado_box(parent, bg, height=12)


    def _rede_obter_token(self):
        """Obtém token OAuth2 da Rede."""
        import requests, base64
        client_id     = self._rede_vars.get("client_id",     tk.StringVar()).get().strip()
        client_secret = self._rede_vars.get("client_secret", tk.StringVar()).get().strip()
        if not client_id or not client_secret:
            raise ValueError("Client ID e Client Secret são obrigatórios.")
        cred = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
        resp = requests.post(
            f"{REDE_API_URL}/v1/token",
            headers={"Authorization": f"Basic {cred}",
                     "Content-Type": "application/x-www-form-urlencoded"},
            data={"grant_type": "client_credentials"},
            timeout=15
        )
        if resp.status_code != 200:
            raise ConnectionError(f"Erro ao obter token Rede ({resp.status_code}): {resp.text[:200]}")
        return resp.json().get("access_token", "")


    def _consultar_rede_vendas(self):
        if not self._checar_requests():
            self._log_resultado(self._rede_resultado, "❌ 'requests' não instalado. Execute: pip install requests", "#f87171")
            return
        self._limpar_resultado(self._rede_resultado)
        self._log_resultado(self._rede_resultado, "⏳ Consultando vendas Rede...")

        def consultar():
            try:
                token = self._rede_obter_token()
                pv    = self._rede_vars.get("pv", tk.StringVar()).get().strip()
                data_str = self._rede_data.get().strip()
                try:
                    dt = datetime.strptime(data_str, "%d/%m/%Y")
                    dt_fmt = dt.strftime("%Y%m%d")
                except:
                    raise ValueError("Data inválida. Use DD/MM/AAAA.")

                import requests
                resp = requests.get(
                    f"{REDE_API_URL}/v1/transactions",
                    headers={"Authorization": f"Bearer {token}"},
                    params={"pv": pv, "transactionDate": dt_fmt, "pageSize": 100},
                    timeout=15
                )
                if resp.status_code != 200:
                    self.root.after(0, lambda: self._log_resultado(
                        self._rede_resultado, f"❌ Erro {resp.status_code}: {resp.text[:300]}", "#f87171"))
                    return

                data = resp.json()
                transacoes = data.get("transactions", data.get("returnedTransactions", []))

                if not transacoes:
                    self.root.after(0, lambda: self._log_resultado(
                        self._rede_resultado, "ℹ️ Nenhuma venda encontrada.", "#fbbf24"))
                    return

                total = sum(float(t.get("amount", 0)) for t in transacoes)
                aprovadas = [t for t in transacoes if str(t.get("returnCode", "")) in ("00", "0")]

                def atualizar():
                    self._log_resultado(self._rede_resultado,
                        f"✅ Rede — {data_str}\n"
                        f"   Total transações: {len(transacoes)}\n"
                        f"   Aprovadas:        {len(aprovadas)}\n"
                        f"   Valor total:      R$ {total:,.2f}\n"
                        f"{'─'*55}", "#4ade80")

                    # Agrupa por tipo de cartão
                    por_tipo = {}
                    for t in aprovadas:
                        kind = t.get("kind", t.get("cardType", "Outros"))
                        por_tipo.setdefault(kind, {"qtd": 0, "total": 0.0})
                        por_tipo[kind]["qtd"] += 1
                        por_tipo[kind]["total"] += float(t.get("amount", 0))

                    for tipo, vals in sorted(por_tipo.items()):
                        self._log_resultado(self._rede_resultado,
                            f"  💳 {tipo:<20} {vals['qtd']:>4} transação(ões)   R$ {vals['total']:>10,.2f}")

                self.root.after(0, atualizar)

            except Exception as e:
                self.root.after(0, lambda: self._log_resultado(
                    self._rede_resultado, f"❌ Erro: {e}", "#f87171"))

        threading.Thread(target=consultar, daemon=True).start()


    def _consultar_rede_recebiveis(self):
        if not self._checar_requests():
            self._log_resultado(self._rede_resultado, "❌ 'requests' não instalado.", "#f87171")
            return
        self._limpar_resultado(self._rede_resultado)
        self._log_resultado(self._rede_resultado, "⏳ Consultando agenda de recebíveis Rede...")

        def consultar():
            try:
                token = self._rede_obter_token()
                pv    = self._rede_vars.get("pv", tk.StringVar()).get().strip()

                import requests
                hoje = datetime.now()
                params = {
                    "pv": pv,
                    "startDate": hoje.strftime("%Y%m%d"),
                    "endDate": (hoje.replace(day=28) if hoje.day < 28
                                else hoje.replace(month=hoje.month % 12 + 1, day=1)).strftime("%Y%m%d"),
                    "pageSize": 100
                }
                resp = requests.get(
                    f"{REDE_API_URL}/v1/receivables",
                    headers={"Authorization": f"Bearer {token}"},
                    params=params, timeout=15
                )
                if resp.status_code != 200:
                    self.root.after(0, lambda: self._log_resultado(
                        self._rede_resultado, f"❌ Erro {resp.status_code}: {resp.text[:300]}", "#f87171"))
                    return

                data = resp.json()
                recebiveis = data.get("receivables", data.get("returnedReceivables", []))

                if not recebiveis:
                    self.root.after(0, lambda: self._log_resultado(
                        self._rede_resultado, "ℹ️ Nenhum recebível encontrado.", "#fbbf24"))
                    return

                total = sum(float(r.get("grossAmount", r.get("amount", 0))) for r in recebiveis)
                total_liq = sum(float(r.get("netAmount", r.get("amount", 0))) for r in recebiveis)

                def atualizar():
                    self._log_resultado(self._rede_resultado,
                        f"✅ Recebíveis Rede — próximos 30 dias\n"
                        f"   {len(recebiveis)} lançamento(s)\n"
                        f"   Valor bruto:   R$ {total:>12,.2f}\n"
                        f"   Valor líquido: R$ {total_liq:>12,.2f}\n"
                        f"{'─'*55}", "#4ade80")

                    por_data = {}
                    for r in recebiveis:
                        dt = r.get("expectedPaymentDate", r.get("paymentDate", ""))[:10]
                        por_data.setdefault(dt, 0.0)
                        por_data[dt] += float(r.get("netAmount", r.get("amount", 0)))

                    for dt in sorted(por_data.keys()):
                        try:
                            dt_fmt = datetime.strptime(dt, "%Y-%m-%d").strftime("%d/%m/%Y")
                        except:
                            dt_fmt = dt
                        self._log_resultado(self._rede_resultado,
                            f"  📅 {dt_fmt}   R$ {por_data[dt]:>12,.2f}")

                self.root.after(0, atualizar)

            except Exception as e:
                self.root.after(0, lambda: self._log_resultado(
                    self._rede_resultado, f"❌ Erro: {e}", "#f87171"))

        threading.Thread(target=consultar, daemon=True).start()


    # ── SUB-ABA GETNET ──────────────────────────────────────────────────────

    def _build_sub_getnet(self, parent, bg, surface, accent, green, yellow, text, muted):
        self._getnet_vars = {}

        info = tk.Frame(parent, bg=surface, padx=14, pady=10)
        info.pack(fill="x", padx=0, pady=(10, 8))
        tk.Label(info, text="🟢  Getnet — Vendas & Recebíveis",
                 font=("Segoe UI", 10, "bold"), fg="#4ade80", bg=surface).pack(anchor="w")
        tk.Label(info,
                 text="Consulta vendas e agenda de recebíveis via API Getnet (Santander).\n"
                      "Credenciais: acesse developers.getnet.com.br → Minha Conta → Credenciais.",
                 font=("Segoe UI", 8), fg=muted, bg=surface, justify="left").pack(anchor="w", pady=(4, 0))

        self._card_credencial(
            parent, bg, surface, muted, text, accent,
            titulo="🔑 Credenciais Getnet",
            status_cfg=[GETNET_CLIENT_ID, GETNET_CLIENT_SECRET, GETNET_SELLER_ID],
            campos=[
                ("Client ID:", "client_id", GETNET_CLIENT_ID, False),
                ("Client Secret:", "client_secret", GETNET_CLIENT_SECRET, True),
                ("Seller ID:", "seller_id", GETNET_SELLER_ID, False),
            ],
            var_dict=self._getnet_vars
        )

        filtros = tk.Frame(parent, bg=surface, padx=14, pady=8)
        filtros.pack(fill="x", pady=(0, 8))
        tk.Label(filtros, text="Data:", fg=muted, bg=surface, font=("Segoe UI", 8)).grid(row=0, column=0, padx=(0, 4))
        self._getnet_data = tk.Entry(filtros, font=("Segoe UI", 9), bg="#0a0f1e", fg=text,
                                      insertbackground=accent, relief="flat", bd=3, width=12)
        self._getnet_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self._getnet_data.grid(row=0, column=1, padx=(0, 12))

        tk.Button(filtros, text="📊 Consultar Vendas",
                  font=("Segoe UI", 9, "bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=self._consultar_getnet_vendas).grid(row=0, column=2, padx=(0, 8))

        tk.Button(filtros, text="📅 Agenda de Recebíveis",
                  font=("Segoe UI", 9, "bold"), bg="#065f46", fg="white",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=self._consultar_getnet_recebiveis).grid(row=0, column=3)

        self._getnet_resultado = self._resultado_box(parent, bg, height=12)


    def _getnet_obter_token(self):
        """Obtém token OAuth2 da Getnet."""
        import requests, base64
        client_id     = self._getnet_vars.get("client_id",     tk.StringVar()).get().strip()
        client_secret = self._getnet_vars.get("client_secret", tk.StringVar()).get().strip()
        seller_id     = self._getnet_vars.get("seller_id",     tk.StringVar()).get().strip()
        if not client_id or not client_secret:
            raise ValueError("Client ID e Client Secret são obrigatórios.")
        cred = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
        resp = requests.post(
            f"{GETNET_API_URL}/auth/oauth/v2/token",
            headers={"Authorization": f"Basic {cred}",
                     "Content-Type": "application/x-www-form-urlencoded"},
            data={"grant_type": "client_credentials", "scope": "oob"},
            timeout=15
        )
        if resp.status_code != 200:
            raise ConnectionError(f"Erro ao obter token Getnet ({resp.status_code}): {resp.text[:200]}")
        return resp.json().get("access_token", ""), seller_id


    def _consultar_getnet_vendas(self):
        if not self._checar_requests():
            self._log_resultado(self._getnet_resultado, "❌ 'requests' não instalado.", "#f87171")
            return
        self._limpar_resultado(self._getnet_resultado)
        self._log_resultado(self._getnet_resultado, "⏳ Consultando vendas Getnet...")

        def consultar():
            try:
                token, seller_id = self._getnet_obter_token()
                data_str = self._getnet_data.get().strip()
                try:
                    dt = datetime.strptime(data_str, "%d/%m/%Y")
                except:
                    raise ValueError("Data inválida.")

                import requests
                resp = requests.get(
                    f"{GETNET_API_URL}/v1/reports/sales/transactions",
                    headers={"Authorization": f"Bearer {token}",
                             "seller_id": seller_id},
                    params={
                        "transaction_date_initial": dt.strftime("%Y-%m-%dT00:00:00.000Z"),
                        "transaction_date_final":   dt.strftime("%Y-%m-%dT23:59:59.999Z"),
                        "page": 1, "limit": 100
                    },
                    timeout=15
                )
                if resp.status_code != 200:
                    self.root.after(0, lambda: self._log_resultado(
                        self._getnet_resultado, f"❌ Erro {resp.status_code}: {resp.text[:300]}", "#f87171"))
                    return

                data = resp.json()
                vendas = data.get("transactions", data.get("sales", []))

                if not vendas:
                    self.root.after(0, lambda: self._log_resultado(
                        self._getnet_resultado, "ℹ️ Nenhuma venda encontrada.", "#fbbf24"))
                    return

                total = sum(float(v.get("amount", 0)) / 100 for v in vendas)
                aprovadas = [v for v in vendas if v.get("status", "").upper() in ("APPROVED", "CONFIRMED", "PAID")]

                def atualizar():
                    self._log_resultado(self._getnet_resultado,
                        f"✅ Getnet — {data_str}\n"
                        f"   Total transações: {len(vendas)}\n"
                        f"   Aprovadas:        {len(aprovadas)}\n"
                        f"   Valor total:      R$ {total:,.2f}\n"
                        f"{'─'*55}", "#4ade80")

                    por_tipo = {}
                    for v in aprovadas:
                        kind = v.get("payment_type", v.get("card_type", "Outros"))
                        por_tipo.setdefault(kind, {"qtd": 0, "total": 0.0})
                        por_tipo[kind]["qtd"] += 1
                        por_tipo[kind]["total"] += float(v.get("amount", 0)) / 100

                    for tipo, vals in sorted(por_tipo.items()):
                        self._log_resultado(self._getnet_resultado,
                            f"  💳 {tipo:<20} {vals['qtd']:>4} venda(s)   R$ {vals['total']:>10,.2f}")

                self.root.after(0, atualizar)

            except Exception as e:
                self.root.after(0, lambda: self._log_resultado(
                    self._getnet_resultado, f"❌ Erro: {e}", "#f87171"))

        threading.Thread(target=consultar, daemon=True).start()


    def _consultar_getnet_recebiveis(self):
        if not self._checar_requests():
            self._log_resultado(self._getnet_resultado, "❌ 'requests' não instalado.", "#f87171")
            return
        self._limpar_resultado(self._getnet_resultado)
        self._log_resultado(self._getnet_resultado, "⏳ Consultando agenda Getnet...")

        def consultar():
            try:
                token, seller_id = self._getnet_obter_token()
                import requests
                hoje = datetime.now()
                from datetime import timedelta
                fim  = hoje + timedelta(days=30)
                resp = requests.get(
                    f"{GETNET_API_URL}/v1/reports/receivables",
                    headers={"Authorization": f"Bearer {token}", "seller_id": seller_id},
                    params={
                        "payment_date_initial": hoje.strftime("%Y-%m-%d"),
                        "payment_date_final":   fim.strftime("%Y-%m-%d"),
                        "page": 1, "limit": 100
                    },
                    timeout=15
                )
                if resp.status_code != 200:
                    self.root.after(0, lambda: self._log_resultado(
                        self._getnet_resultado, f"❌ Erro {resp.status_code}: {resp.text[:300]}", "#f87171"))
                    return

                data = resp.json()
                recebiveis = data.get("receivables", [])

                if not recebiveis:
                    self.root.after(0, lambda: self._log_resultado(
                        self._getnet_resultado, "ℹ️ Nenhum recebível encontrado.", "#fbbf24"))
                    return

                total_bruto = sum(float(r.get("gross_amount", 0)) / 100 for r in recebiveis)
                total_liq   = sum(float(r.get("net_amount",   0)) / 100 for r in recebiveis)

                def atualizar():
                    self._log_resultado(self._getnet_resultado,
                        f"✅ Agenda Getnet — próximos 30 dias\n"
                        f"   {len(recebiveis)} lançamento(s)\n"
                        f"   Valor bruto:   R$ {total_bruto:>12,.2f}\n"
                        f"   Valor líquido: R$ {total_liq:>12,.2f}\n"
                        f"{'─'*55}", "#4ade80")

                    por_data = {}
                    for r in recebiveis:
                        dt = r.get("payment_date", "")[:10]
                        por_data.setdefault(dt, 0.0)
                        por_data[dt] += float(r.get("net_amount", 0)) / 100

                    for dt in sorted(por_data.keys()):
                        try:
                            dt_fmt = datetime.strptime(dt, "%Y-%m-%d").strftime("%d/%m/%Y")
                        except:
                            dt_fmt = dt
                        self._log_resultado(self._getnet_resultado,
                            f"  📅 {dt_fmt}   R$ {por_data[dt]:>12,.2f}")

                self.root.after(0, atualizar)

            except Exception as e:
                self.root.after(0, lambda: self._log_resultado(
                    self._getnet_resultado, f"❌ Erro: {e}", "#f87171"))

        threading.Thread(target=consultar, daemon=True).start()


    # ── SUB-ABA MICROVIX ────────────────────────────────────────────────────

    def _build_sub_microvix(self, parent, bg, surface, accent, green, yellow, text, muted):
        self._mv_vars = {}

        info = tk.Frame(parent, bg=surface, padx=14, pady=10)
        info.pack(fill="x", padx=0, pady=(10, 8))
        tk.Label(info, text="🏪  Linx Microvix — Vendas por Filial",
                 font=("Segoe UI", 10, "bold"), fg=yellow, bg=surface).pack(anchor="w")
        tk.Label(info,
                 text="Consulta vendas do dia por filial via API Linx Microvix.\n"
                      "Chave da empresa: acesse Microvix → Configurações → Integrações → Chave de Acesso.",
                 font=("Segoe UI", 8), fg=muted, bg=surface, justify="left").pack(anchor="w", pady=(4, 0))

        self._card_credencial(
            parent, bg, surface, muted, text, accent,
            titulo="🔑 Credenciais Microvix",
            status_cfg=[MICROVIX_CNPJ, MICROVIX_USUARIO, MICROVIX_CHAVE],
            campos=[
                ("CNPJ (só números):", "cnpj",    MICROVIX_CNPJ,    False),
                ("Usuário:",           "usuario",  MICROVIX_USUARIO, False),
                ("Senha:",             "senha",    MICROVIX_SENHA,   True),
                ("Chave da Empresa:",  "chave",    MICROVIX_CHAVE,   True),
            ],
            var_dict=self._mv_vars
        )

        filtros = tk.Frame(parent, bg=surface, padx=14, pady=8)
        filtros.pack(fill="x", pady=(0, 8))
        tk.Label(filtros, text="Data:", fg=muted, bg=surface, font=("Segoe UI", 8)).grid(row=0, column=0, padx=(0, 4))
        self._mv_data = tk.Entry(filtros, font=("Segoe UI", 9), bg="#0a0f1e", fg=text,
                                  insertbackground=accent, relief="flat", bd=3, width=12)
        self._mv_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self._mv_data.grid(row=0, column=1, padx=(0, 12))

        tk.Label(filtros, text="Filial:", fg=muted, bg=surface, font=("Segoe UI", 8)).grid(row=0, column=2, padx=(0, 4))
        self._mv_filial = tk.Entry(filtros, font=("Segoe UI", 9), bg="#0a0f1e", fg=text,
                                    insertbackground=accent, relief="flat", bd=3, width=8)
        self._mv_filial.insert(0, "0")  # 0 = todas
        self._mv_filial.grid(row=0, column=3, padx=(0, 12))
        tk.Label(filtros, text="(0=todas)", fg=muted, bg=surface,
                 font=("Consolas", 7)).grid(row=0, column=4, padx=(0, 12))

        tk.Button(filtros, text="📊 Consultar Vendas",
                  font=("Segoe UI", 9, "bold"), bg="#b45309", fg="white",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=self._consultar_microvix_vendas).grid(row=0, column=5, padx=(0, 8))

        tk.Button(filtros, text="📦 Consultar Estoque",
                  font=("Segoe UI", 9, "bold"), bg="#78350f", fg="white",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=self._consultar_microvix_estoque).grid(row=0, column=6)

        self._mv_resultado = self._resultado_box(parent, bg, height=12)


    def _microvix_xml_request(self, metodo, params_extras=None):
        """Monta e envia requisição SOAP para o Microvix."""
        import requests
        cnpj    = self._mv_vars.get("cnpj",    tk.StringVar()).get().strip()
        usuario = self._mv_vars.get("usuario", tk.StringVar()).get().strip()
        senha   = self._mv_vars.get("senha",   tk.StringVar()).get().strip()
        chave   = self._mv_vars.get("chave",   tk.StringVar()).get().strip()

        if not cnpj or not usuario or not chave:
            raise ValueError("CNPJ, Usuário e Chave são obrigatórios.")

        params_xml = f"""<param>
            <cnpj_emp>{cnpj}</cnpj_emp>
            <nome_sistema>RoboFinanceiro</nome_sistema>
            <chave>{chave}</chave>
            <usuario>{usuario}</usuario>
            <senha>{senha}</senha>"""
        if params_extras:
            for k, v in params_extras.items():
                params_xml += f"\n            <{k}>{v}</{k}>"
        params_xml += "\n        </param>"

        soap = f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <{metodo} xmlns="http://linx.com.br/microvix/integracao">
        {params_xml}
    </{metodo}>
  </soap:Body>
</soap:Envelope>"""

        resp = requests.post(
            MICROVIX_API_URL,
            data=soap.encode("utf-8"),
            headers={"Content-Type": "text/xml; charset=utf-8",
                     "SOAPAction": f"http://linx.com.br/microvix/integracao/{metodo}"},
            timeout=20
        )
        if resp.status_code != 200:
            raise ConnectionError(f"Erro HTTP {resp.status_code}: {resp.text[:300]}")
        return resp.text


    def _consultar_microvix_vendas(self):
        if not self._checar_requests():
            self._log_resultado(self._mv_resultado, "❌ 'requests' não instalado.", "#f87171")
            return
        self._limpar_resultado(self._mv_resultado)
        self._log_resultado(self._mv_resultado, "⏳ Consultando vendas Microvix...")

        def consultar():
            try:
                import xml.etree.ElementTree as ET
                data_str = self._mv_data.get().strip()
                filial   = self._mv_filial.get().strip() or "0"
                try:
                    dt = datetime.strptime(data_str, "%d/%m/%Y")
                    dt_mv = dt.strftime("%Y-%m-%d")
                except:
                    raise ValueError("Data inválida.")

                xml_resp = self._microvix_xml_request(
                    "LinxMovimentos",
                    {"dtini": dt_mv, "dtfim": dt_mv, "cod_filial": filial}
                )

                root = ET_SAFE.fromstring(xml_resp)  # nosec B314
                ns   = {"s": "http://schemas.xmlsoap.org/soap/envelope/"}

                # Encontra todos os elementos de venda
                movs = root.findall(".//{http://linx.com.br/microvix/integracao}LinxMovimentosResult//{http://linx.com.br/microvix/integracao}Table")
                if not movs:
                    # Tenta sem namespace
                    movs = root.findall(".//Table")

                if not movs:
                    self.root.after(0, lambda: self._log_resultado(
                        self._mv_resultado,
                        f"ℹ️ Nenhuma venda encontrada.\nResposta: {xml_resp[:400]}", "#fbbf24"))
                    return

                total = 0.0
                por_filial = {}
                por_forma  = {}
                for m in movs:
                    def get(tag):
                        el = m.find(tag)
                        return el.text.strip() if el is not None and el.text else ""
                    try:
                        valor = float(get("vl_total") or get("valor") or "0")
                    except:
                        valor = 0.0
                    fil   = get("cod_filial") or get("filial") or "?"
                    forma = get("descricao_forma_pagamento") or get("forma_pagamento") or "Outros"
                    total += valor
                    por_filial.setdefault(fil, 0.0)
                    por_filial[fil] += valor
                    por_forma.setdefault(forma, {"qtd": 0, "total": 0.0})
                    por_forma[forma]["qtd"] += 1
                    por_forma[forma]["total"] += valor

                def atualizar():
                    self._log_resultado(self._mv_resultado,
                        f"✅ Microvix — {data_str}\n"
                        f"   {len(movs)} movimento(s) | Total: R$ {total:,.2f}\n"
                        f"{'─'*55}", "#4ade80")
                    self._log_resultado(self._mv_resultado, "  📍 Por Filial:")
                    for fil, val in sorted(por_filial.items()):
                        nome_fil = DE_PARA_FILIAL.get(f"Filial_{fil}", fil)
                        self._log_resultado(self._mv_resultado,
                            f"     {nome_fil:<25}  R$ {val:>12,.2f}")
                    self._log_resultado(self._mv_resultado, "\n  💳 Por Forma de Pagamento:")
                    for forma, vals in sorted(por_forma.items(), key=lambda x: -x[1]["total"]):
                        self._log_resultado(self._mv_resultado,
                            f"     {forma:<25} {vals['qtd']:>4}x  R$ {vals['total']:>12,.2f}")

                self.root.after(0, atualizar)

            except Exception as e:
                self.root.after(0, lambda: self._log_resultado(
                    self._mv_resultado, f"❌ Erro: {e}", "#f87171"))

        threading.Thread(target=consultar, daemon=True).start()


    def _consultar_microvix_estoque(self):
        if not self._checar_requests():
            self._log_resultado(self._mv_resultado, "❌ 'requests' não instalado.", "#f87171")
            return
        self._limpar_resultado(self._mv_resultado)
        self._log_resultado(self._mv_resultado, "⏳ Consultando estoque Microvix...")

        def consultar():
            try:
                import xml.etree.ElementTree as ET
                filial = self._mv_filial.get().strip() or "0"

                xml_resp = self._microvix_xml_request(
                    "LinxProdutos",
                    {"cod_filial": filial, "apenas_ativos": "1"}
                )

                root = ET_SAFE.fromstring(xml_resp)  # nosec B314
                prods = root.findall(".//Table")

                if not prods:
                    self.root.after(0, lambda: self._log_resultado(
                        self._mv_resultado,
                        f"ℹ️ Nenhum produto encontrado.\nResposta: {xml_resp[:400]}", "#fbbf24"))
                    return

                total_itens = len(prods)
                valor_estoque = 0.0
                sem_estoque   = 0
                for p in prods:
                    def get(tag):
                        el = p.find(tag)
                        return el.text.strip() if el is not None and el.text else "0"
                    try:
                        qtd  = float(get("qt_estoque") or get("estoque") or "0")
                        custo = float(get("vl_custo") or get("preco_custo") or "0")
                        valor_estoque += qtd * custo
                        if qtd <= 0:
                            sem_estoque += 1
                    except:
                        pass

                def atualizar():
                    self._log_resultado(self._mv_resultado,
                        f"✅ Estoque Microvix\n"
                        f"   Total SKUs:         {total_itens}\n"
                        f"   Sem estoque:        {sem_estoque}\n"
                        f"   Valor em estoque:   R$ {valor_estoque:>12,.2f}", "#4ade80")

                self.root.after(0, atualizar)

            except Exception as e:
                self.root.after(0, lambda: self._log_resultado(
                    self._mv_resultado, f"❌ Erro: {e}", "#f87171"))

        threading.Thread(target=consultar, daemon=True).start()
