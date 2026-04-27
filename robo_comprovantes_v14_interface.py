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
import xml.etree.ElementTree as ET_SAFE


from config import *
from utils.data_processing import *
from utils.ocr_utils import *
from utils.email_service import *
from utils.file_manager import *
from utils.integrations_utils import *
from utils.extratos_processor import *
from utils.dashboard_service import *
from utils.formatters import *
import database
import cnab_generator

from database import *
from cnab_generator import *

def _aplicar_mascara_data(widget):
    """Aplica máscara de data (DD/MM/AAAA) a um widget Entry do Tkinter."""
    def format_date(event):
        if event.keysym in ("BackSpace", "Delete", "Left", "Right", "Tab"):
            return
        text = widget.get().replace('/', '')
        text = ''.join(c for c in text if c.isdigit())[:8]
        new_text = ""
        if len(text) > 0: new_text += text[:2]
        if len(text) > 2: new_text += '/' + text[2:4]
        if len(text) > 4: new_text += '/' + text[4:8]
        widget.delete(0, tk.END)
        widget.insert(0, new_text)
    widget.bind("<KeyRelease>", format_date)



# ==============================================================================
# --- INTERFACE GRÁFICA ---
# ==============================================================================

class RoboApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ComAgente ERP")
        self.root.geometry("900x680")
        self.root.resizable(True, True)
        self.root.state('zoomed')  # Maximiza no Windows
        self.root.configure(bg="#0f172a")
        self.rodando = False
        self._log_historico = []
        self._resultados_busca = []
        self._build_ui()

    def _build_ui(self):
        bg = "#0f172a"
        surface = "#1e293b"
        border = "#334155"
        accent = "#38bdf8"
        green = "#4ade80"
        yellow = "#fbbf24"
        text = "#e2e8f0"
        muted = "#94a3b8"

        # ── HEADER ──
        header = tk.Frame(self.root, bg="#0c1628", pady=14)
        header.pack(fill="x")
        tk.Label(header, text="🤖", font=("Segoe UI Emoji", 28), bg="#0c1628").pack(side="left", padx=(20, 10))
        title_frame = tk.Frame(header, bg="#0c1628")
        title_frame.pack(side="left")
        tk.Label(title_frame, text="Robô Organizador de Comprovantes", font=("Segoe UI", 15, "bold"), fg=accent, bg="#0c1628").pack(anchor="w")
        tk.Label(title_frame, text="BOAH Modas  •  Solar  •  v17.0", font=("Consolas", 10), fg=muted, bg="#0c1628").pack(anchor="w")

        # ── ABAS ──
        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook", background=bg, borderwidth=0)
        style.configure("TNotebook.Tab", background=surface, foreground=muted,
                        padding=[16, 8], font=("Segoe UI", 10, "bold"))
        style.map("TNotebook.Tab",
                  background=[("selected", accent)],
                  foreground=[("selected", "#0f172a")])
        style.configure("TProgressbar", troughcolor=surface, background=accent, thickness=6)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=0, pady=0)

        # ── ABA 1: ROBÔ ──
        aba_robo = tk.Frame(self.notebook, bg=bg)
        self.notebook.add(aba_robo, text="  🤖 Robô  ")
        self._build_aba_robo(aba_robo, bg, surface, border, accent, green, yellow, text, muted)

        # ── ABA 2: BUSCA ──
        aba_busca = tk.Frame(self.notebook, bg=bg)
        self.notebook.add(aba_busca, text="  🔍 Buscar Comprovante  ")
        self._build_aba_busca(aba_busca, bg, surface, border, accent, green, text, muted)

        # ── ABA 3: EXTRATO ──
        aba_extrato = tk.Frame(self.notebook, bg=bg)
        self.notebook.add(aba_extrato, text="  📊 Extrato / Fluxo de Caixa  ")
        self._build_aba_extrato(aba_extrato, bg, surface, border, accent, green, yellow, text, muted)

        # ── ABA 4: HISTÓRICO DE STATUS ──
        aba_hist_status = tk.Frame(self.notebook, bg=bg)
        self.notebook.add(aba_hist_status, text="  📜 Histórico de Status  ")
        self._build_aba_historico_status(aba_hist_status, bg, surface, border, accent, green, yellow, text, muted)

        # ── ABA 5: HISTÓRICO DE FLUXO ──
        aba_hist_fluxo = tk.Frame(self.notebook, bg=bg)
        self.notebook.add(aba_hist_fluxo, text="  📈 Histórico de Fluxo  ")
        self._build_aba_historico_fluxo(aba_hist_fluxo, bg, surface, border, accent, green, yellow, text, muted)

        # ── ABA 6: RECEBÍVEIS / INTEGRAÇÕES ──
        aba_recebiveis = tk.Frame(self.notebook, bg=bg)
        self.notebook.add(aba_recebiveis, text="  💳 Recebíveis  ")
        self._build_aba_recebiveis(aba_recebiveis, bg, surface, border, accent, green, yellow, text, muted)

        # ── ABA 7: AUTORIZAÇÃO DE PAGAMENTOS ──
        aba_autorizacao = tk.Frame(self.notebook, bg=bg)
        self.notebook.add(aba_autorizacao, text="  📋 Autorização de Pagamentos  ")
        self._build_aba_autorizacao(aba_autorizacao, bg, surface, border, accent, green, yellow, text, muted)

        # ── ABA 8: CONTAS A PAGAR / LANÇAMENTO DE NOTAS ──
        aba_cap = tk.Frame(self.notebook, bg=bg)
        self.notebook.add(aba_cap, text="  🧾 Contas a Pagar  ")
        self._build_aba_contas_pagar(aba_cap, bg, surface, border, accent, green, yellow, text, muted)

        # ── ABA 9: RESTITUIÇÕES SEFAZ (NOVA) ──
        aba_restituicoes = tk.Frame(self.notebook, bg=bg)
        self.notebook.add(aba_restituicoes, text="  🏛️ Restituições SEFAZ  ")
        self._build_aba_restituicoes(aba_restituicoes, bg, surface, border, accent, green, yellow, text, muted)

    # ══════════════════════════════════════════════════════════════════════════
    # ABA CONTAS A PAGAR
    # ══════════════════════════════════════════════════════════════════════════

    def _build_aba_contas_pagar(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        _db_init()

        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True, padx=0, pady=0)

        sub_lista = tk.Frame(nb, bg=bg)
        sub_nota  = tk.Frame(nb, bg=bg)
        sub_adian = tk.Frame(nb, bg=bg)
        sub_prest = tk.Frame(nb, bg=bg)
        sub_imp   = tk.Frame(nb, bg=bg)
        sub_forn  = tk.Frame(nb, bg=bg)
        sub_cnab  = tk.Frame(nb, bg=bg)

        nb.add(sub_lista, text="  📋 Notas Lançadas  ")
        nb.add(sub_nota,  text="  ➕ Nova Nota  ")
        nb.add(sub_adian, text="  💰 Adiantamentos  ")
        nb.add(sub_prest, text="  ✅ Prestação de Contas  ")
        nb.add(sub_imp,   text="  🏛️ Impostos Retidos  ")
        nb.add(sub_forn,  text="  🏢 Fornecedores  ")
        nb.add(sub_cnab,  text="  🏦 CNAB 240  ")

        self._cap_build_lista(sub_lista, bg, surface, accent, green, yellow, text, muted, nb)
        self._cap_build_nova_nota(sub_nota, bg, surface, accent, green, yellow, text, muted, nb)
        self._cap_build_adiantamentos(sub_adian, bg, surface, accent, green, yellow, text, muted)
        self._cap_build_prestacao(sub_prest, bg, surface, accent, green, yellow, text, muted)
        self._cap_build_impostos(sub_imp, bg, surface, accent, green, yellow, text, muted)
        self._cap_build_fornecedores(sub_forn, bg, surface, accent, green, yellow, text, muted)
        self._cap_build_cnab(sub_cnab, bg, surface, accent, green, yellow, text, muted)

    # ── SUB-ABA: LISTA DE NOTAS ───────────────────────────────────────────────
    def _cap_consultar_sefaz(self):
        """Consulta NF-e pela chave de acesso.
        Camada 1: extrai dados da própria chave (sempre funciona)
        Camada 2: consulta webservice SEFAZ com certificado A1
        """
        chave = re.sub(r'\D', '', self._cap_chave_var.get().strip())

        if len(chave) != 44:
            self._cap_lbl_xml.config(
                text=f"⚠️ Chave inválida — {len(chave)} dígitos (precisa de 44)",
                fg="#fbbf24")
            return

        # ── CAMADA 1: dados embutidos na chave ──────────────────────────────
        # Estrutura: cUF(2) AAMM(4) CNPJ(14) mod(2) serie(3) nNF(9) tpEmis(1) cNF(8) cDV(1)
        try:
            c_uf    = chave[0:2]
            c_aamm  = chave[2:6]
            c_cnpj  = chave[6:20]
            c_mod   = chave[20:22]
            c_serie = chave[22:25]
            c_nnf   = chave[25:34]
            c_cdv   = chave[43]

            ano  = "20" + c_aamm[0:2]
            mes  = c_aamm[2:4]
            cnpj_fmt = f"{c_cnpj[:2]}.{c_cnpj[2:5]}.{c_cnpj[5:8]}/{c_cnpj[8:12]}-{c_cnpj[12:14]}"
            num_nf = str(int(c_nnf))

            UF_NOMES = {
                "11":"RO","12":"AC","13":"AM","14":"RR","15":"PA","16":"AP","17":"TO",
                "21":"MA","22":"PI","23":"CE","24":"RN","25":"PB","26":"PE","27":"AL",
                "28":"SE","29":"BA","31":"MG","32":"ES","33":"RJ","35":"SP","41":"PR",
                "42":"SC","43":"RS","50":"MS","51":"MT","52":"GO","53":"DF",
            }
            uf_sigla = UF_NOMES.get(c_uf, c_uf)

            self._cap_lbl_xml.config(
                text=f"✅ Chave válida — NF {num_nf} | {uf_sigla} | {mes}/{ano} | CNPJ {cnpj_fmt}",
                fg="#4ade80")

            # Preenche campos que já dá para saber pela chave
            self._cap_cnpj.delete(0,"end")
            self._cap_cnpj.insert(0, cnpj_fmt)
            self._cap_desc.delete(0,"end")
            self._cap_desc.insert(0, f"NF {num_nf}")

            # Tenta auto-completar pelo CNPJ no banco
            f = buscar_fornecedor_por_cnpj(cnpj_fmt)
            if f:
                self._cap_forn.delete(0,"end")
                self._cap_forn.insert(0, f["razao_social"])
                if f["categoria"]: self._cap_cat_var.set(f["categoria"])
                if f["responsavel"]: self._cap_resp.set(f["responsavel"])
                if hasattr(self,"_cap_preencher_banco"):
                    self._cap_preencher_banco(dict(f))

        except Exception as e:
            self._cap_lbl_xml.config(text=f"❌ Erro ao ler chave: {e}", fg="#f87171")
            return

        # ── CAMADA 2: consulta SEFAZ com certificado A1 ─────────────────────
        # Procura certificados .pfx na pasta FINANCEIRO
        pfx_paths = []
        for fname in os.listdir(PASTA_ATUAL):
            if fname.lower().endswith(".pfx"):
                pfx_paths.append(os.path.join(PASTA_ATUAL, fname))

        if not pfx_paths:
            self._cap_lbl_xml.config(
                text="✅ Dados da chave extraídos. "
                     "Para consultar XML completo, exporte o certificado .pfx para a pasta FINANCEIRO.",
                fg="#fbbf24")
            return

        # Tenta consultar em background
        def _consultar_bg():
            try:
                import requests
                from cryptography.hazmat.primitives.serialization import pkcs12
                from cryptography.hazmat.primitives.serialization import (
                    Encoding, PrivateFormat, NoEncryption)
                import tempfile

                # Tenta cada .pfx até funcionar
                for pfx_path in pfx_paths:
                    try:
                        # Tenta sem senha primeiro, depois pede
                        senha = None
                        with open(pfx_path,"rb") as f:
                            pfx_data = f.read()

                        # Tenta carregar sem senha
                        try:
                            priv, cert, chain = pkcs12.load_key_and_certificates(
                                pfx_data, None)
                        except:
                            # Pede senha
                            from tkinter import simpledialog
                            senha_str = simpledialog.askstring(
                                "Certificado Digital",
                                f"Senha do certificado:\n{os.path.basename(pfx_path)}",
                                show="*", parent=self.root)
                            if not senha_str: continue
                            priv, cert, chain = pkcs12.load_key_and_certificates(
                                pfx_data, senha_str.encode())

                        # Gera PEM temporários para requests
                        with tempfile.NamedTemporaryFile(suffix=".pem",delete=False,mode="wb") as fc:
                            fc.write(cert.public_bytes(Encoding.PEM))
                            cert_pem = fc.name
                        with tempfile.NamedTemporaryFile(suffix=".pem",delete=False,mode="wb") as fk:
                            fk.write(priv.private_bytes(Encoding.PEM,PrivateFormat.TraditionalOpenSSL,NoEncryption()))
                            key_pem = fk.name

                        # Webservice SEFAZ por UF
                        WS_URL = {
                            "29": "https://nfe.sefaz.ba.gov.br/webservices/nfeconsultaprotocolo4/nfeconsultaprotocolo4.asmx",
                        }.get(c_uf,
                              "https://nfe.fazenda.gov.br/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx")

                        # Envelope SOAP NFeConsultaNF
                        soap_body = f"""<?xml version="1.0" encoding="UTF-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Body>
    <nfeDadosMsg xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4">
      <consSitNFe xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00">
        <tpAmb>1</tpAmb>
        <xServ>CONSULTAR</xServ>
        <chNFe>{chave}</chNFe>
      </consSitNFe>
    </nfeDadosMsg>
  </soap12:Body>
</soap12:Envelope>"""

                        resp = requests.post(
                            WS_URL,
                            data=soap_body.encode("utf-8"),
                            headers={"Content-Type":"application/soap+xml;charset=UTF-8"},
                            cert=(cert_pem, key_pem),
                            verify=False, timeout=15)  # nosec B501 — SEFAZ usa cert ICP-Brasil

                        os.unlink(cert_pem)
                        os.unlink(key_pem)

                        if resp.status_code == 200:
                            # Extrai XML da NF-e da resposta SOAP
                            root_soap = ET_SAFE.fromstring(resp.text)  # nosec B314
                            # Procura o nfeProc ou NFe na resposta
                            xml_nfe = None
                            for el in root_soap.iter():
                                tag = el.tag.split('}')[1] if '}' in el.tag else el.tag
                                if tag in ("nfeProc","NFe","retConsSitNFe"):
                                    xml_nfe = ET.tostring(el, encoding="unicode")
                                    break

                            if xml_nfe:
                                # Salva XML temporário e importa
                                tmp = os.path.join(PASTA_ATUAL,f"NF_{chave}.xml")
                                with open(tmp,"w",encoding="utf-8") as f:
                                    f.write(xml_nfe)
                                self.root.after(0, lambda: (
                                    self._cap_chave_var.set(""),
                                    self._cap_importar_xml_path(tmp),
                                ))
                                return
                    except Exception as e2:
                        continue

                self.root.after(0, lambda: self._cap_lbl_xml.config(
                    text="⚠️ Consulta SEFAZ falhou — campos extraídos da chave apenas.",
                    fg="#fbbf24"))

            except Exception as e:
                self.root.after(0, lambda: self._cap_lbl_xml.config(
                    text=f"⚠️ Certificado: {e}", fg="#fbbf24"))

        import threading
        threading.Thread(target=_consultar_bg, daemon=True).start()
        self._cap_lbl_xml.config(text="⏳ Consultando SEFAZ...", fg="#38bdf8")

    def _cap_importar_xml(self):
        """Abre diálogo para selecionar XML e importa."""
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="Selecionar XML da NF-e",
            filetypes=[("XML NF-e", "*.xml"), ("Todos", "*.*")]
        )
        if path:
            self._cap_importar_xml_path(path)

    def _cap_importar_xml_path(self, path):
        """Lê XML de NF-e e preenche o formulário de nova nota."""
        import xml.etree.ElementTree as ET
        try:
            tree = ET.parse(path)
            root = tree.getroot()

            # Remove namespace para facilitar busca
            def _strip_ns(tag):
                return tag.split('}')[1] if '}' in tag else tag

            def _find(root, *tags):
                """Busca elemento ignorando namespace."""
                for tag in tags:
                    found = None
                    for el in root.iter():
                        if _strip_ns(el.tag) == tag:
                            found = el
                            break
                    if found is None:
                        return None
                    root = found
                return root

            def _val(root, *tags):
                el = _find(root, *tags)
                return el.text.strip() if el is not None and el.text else ""

            # ── Dados do emitente (fornecedor) ──
            cnpj     = _val(root, "emit", "CNPJ")
            razao    = _val(root, "emit", "xNome")
            fantasia = _val(root, "emit", "xFant")
            num_nf   = _val(root, "ide", "nNF")

            # Data emissão — formato 2026-03-20T10:00:00 ou 2026-03-20
            dt_raw = _val(root, "ide", "dhEmi") or _val(root, "ide", "dEmi")
            dt_emissao = ""
            if dt_raw:
                try:
                    from datetime import datetime as dt_cls
                    dt_obj = dt_cls.fromisoformat(dt_raw[:10])
                    dt_emissao = dt_obj.strftime("%d/%m/%Y")
                except: dt_emissao = dt_raw[:10]

            # ── Valores ──
            vNF     = float(_val(root, "total", "ICMSTot", "vNF")    or 0)
            vPIS    = float(_val(root, "total", "ICMSTot", "vPIS")   or 0)
            vCOFINS = float(_val(root, "total", "ICMSTot", "vCOFINS") or 0)
            vIR     = float(_val(root, "total", "ICMSTot", "vIR")    or 0)
            vINSS   = float(_val(root, "total", "ICMSTot", "vINSS")  or 0)
            vISS    = float(_val(root, "total", "ICMSTot", "vISS")   or 0)

            # ── Parcelas (cobr/dup) ──
            parcelas_xml = []
            for dup in root.iter():
                if _strip_ns(dup.tag) == "dup":
                    n_dup  = ""
                    d_venc = ""
                    v_dup  = 0.0
                    for child in dup:
                        t = _strip_ns(child.tag)
                        if t == "nDup":  n_dup  = child.text or ""
                        if t == "dVenc":
                            try:
                                from datetime import datetime as dt_cls
                                d_venc = dt_cls.fromisoformat(child.text[:10]).strftime("%d/%m/%Y")
                            except: d_venc = child.text or ""
                        if t == "vDup":
                            try: v_dup = float(child.text or 0)
                            except: v_dup = 0.0
                    parcelas_xml.append((n_dup, d_venc, v_dup))

            # ── Preenche os campos ──
            # Fornecedor
            self._cap_forn.delete(0,"end")
            self._cap_forn.insert(0, razao)

            # CNPJ — formata XX.XXX.XXX/XXXX-XX
            if cnpj:
                cnpj_fmt = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}"
                self._cap_cnpj.delete(0,"end")
                self._cap_cnpj.insert(0, cnpj_fmt)

            # Descrição com número da NF
            self._cap_desc.delete(0,"end")
            self._cap_desc.insert(0, f"NF {num_nf}" if num_nf else "")

            # Data emissão
            if dt_emissao:
                self._cap_dt_emissao.delete(0,"end")
                self._cap_dt_emissao.insert(0, dt_emissao)

            # Valor bruto
            self._cap_valor.delete(0,"end")
            self._cap_valor.insert(0, f"{vNF:,.2f}".replace(",","X").replace(".",",").replace("X","."))

            # Vencimento — usa primeira parcela se existir
            if parcelas_xml:
                self._cap_dt_venc.delete(0,"end")
                self._cap_dt_venc.insert(0, parcelas_xml[0][1])

            # Impostos — preenche linhas existentes
            impostos_nf = [
                ("PIS",        ALIQUOTAS_PADRAO.get("PIS",0),    vPIS),
                ("COFINS",     ALIQUOTAS_PADRAO.get("COFINS",0), vCOFINS),
                ("IRRF",       ALIQUOTAS_PADRAO.get("IRRF",0),   vIR),
                ("INSS retido",ALIQUOTAS_PADRAO.get("INSS retido",0), vINSS),
                ("ISS",        ALIQUOTAS_PADRAO.get("ISS",0),    vISS),
            ]
            imp_validos = [(t,a,v) for t,a,v in impostos_nf if v > 0]

            for i, row in enumerate(self._cap_imp_rows):
                if i < len(imp_validos):
                    tp, aliq, val = imp_validos[i]
                    row["tipo"].set(tp)
                    row["aliq"].set(str(aliq))
                    row["val"].set(f"{val:.2f}")
                else:
                    row["tipo"].set("")
                    row["aliq"].set("")
                    row["val"].set("")

            # Parcelas — reconstrói se tiver mais de 1
            if len(parcelas_xml) > 1:
                # Limpa parcelas anteriores
                for w in self._cap_parc_widgets:
                    w.destroy()
                self._cap_parc_widgets = []
                self._cap_parc_data    = []

                # Configura spinbox
                self._cap_nparc.delete(0,"end")
                self._cap_nparc.insert(0, str(len(parcelas_xml)))

                # Gera linhas manualmente
                from tkinter import ttk as _ttk
                bg_s = "#0a0f1e"

                hdr = tk.Frame(self._cap_frm_parcelas, bg="#1e293b")
                hdr.pack(fill="x", pady=(0,2))
                for h, w in [("Parcela",60),("Vencimento",110),("Valor",100)]:
                    tk.Label(hdr, text=h, fg="#94a3b8", bg="#1e293b",
                             font=("Segoe UI",7,"bold"), width=w//7
                             ).pack(side="left", padx=4)
                self._cap_parc_widgets.append(hdr)

                for i, (n_dup, d_venc, v_dup) in enumerate(parcelas_xml):
                    row_f = tk.Frame(self._cap_frm_parcelas, bg="#111827")
                    row_f.pack(fill="x", pady=1)
                    tk.Label(row_f, text=f"{i+1}/{len(parcelas_xml)}",
                             fg="#94a3b8", bg="#111827",
                             font=("Segoe UI",8), width=6).pack(side="left", padx=4)
                    dt_var = tk.StringVar(value=d_venc)
                    tk.Entry(row_f, textvariable=dt_var, width=11,
                             font=("Segoe UI",8), bg=bg_s, fg="white",
                             insertbackground="#38bdf8", relief="flat", bd=2
                             ).pack(side="left", padx=4)
                    val_var = tk.StringVar(value=f"{v_dup:.2f}")
                    tk.Entry(row_f, textvariable=val_var, width=10,
                             font=("Segoe UI",8), bg=bg_s, fg="#4ade80",
                             insertbackground="#38bdf8", relief="flat", bd=2
                             ).pack(side="left", padx=4)
                    self._cap_parc_data.append({
                        "dt": dt_var, "val": val_var,
                        "num": i+1, "total": len(parcelas_xml)
                    })
                    self._cap_parc_widgets.append(row_f)

            # Tenta auto-completar pelo CNPJ no banco
            if cnpj:
                f = buscar_fornecedor_por_cnpj(
                    f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}")
                if f and hasattr(self, "_cap_preencher_banco"):
                    self._cap_preencher_banco(dict(f))
                    if f["categoria"]:
                        self._cap_cat_var.set(f["categoria"])
                    if f["responsavel"]:
                        self._cap_resp.set(f["responsavel"])

            # Cadastra fornecedor automaticamente se não existir
            if cnpj and razao:
                cnpj_fmt = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}"
                f_exist = buscar_fornecedor_por_cnpj(cnpj_fmt)
                if not f_exist:
                    salvar_fornecedor({
                        "razao_social": razao,
                        "nome_fantasia": fantasia,
                        "cnpj_cpf": cnpj_fmt,
                        "tipo": "FORNECEDOR",
                        "empresa": self._cap_emp_var.get(),
                        "ativo": 1,
                    })
                    self._cap_lbl_xml.config(
                        text=f"✅ NF {num_nf} importada — fornecedor cadastrado automaticamente")
                else:
                    self._cap_lbl_xml.config(
                        text=f"✅ NF {num_nf} importada")

        except Exception as e:
            self._cap_lbl_xml.config(
                text=f"❌ Erro ao ler XML: {e}", fg="#f87171")

    def _cap_build_lista(self, parent, bg, surface, accent, green, yellow, text, muted, nb_pai):

        # Toolbar
        tb = tk.Frame(parent, bg=surface, pady=6, padx=14)
        tb.pack(fill="x", padx=0, pady=0)

        tk.Label(tb, text="🧾 Contas a Pagar", font=("Segoe UI",12,"bold"),
                 fg=accent, bg=surface).pack(side="left", padx=(0,16))

        tk.Label(tb, text="Status:", fg=muted, bg=surface,
                 font=("Segoe UI",8)).pack(side="left")
        self._cap_filtro_status = ttk.Combobox(tb,
            values=["TODOS","PENDENTE","APROVADA","PAGA","VENCIDA","CANCELADA"],
            state="readonly", width=10, font=("Segoe UI",8))
        self._cap_filtro_status.set("TODOS")
        self._cap_filtro_status.pack(side="left", padx=(2,10))

        tk.Label(tb, text="Empresa:", fg=muted, bg=surface,
                 font=("Segoe UI",8)).pack(side="left")
        self._cap_filtro_emp = ttk.Combobox(tb,
            values=["TODOS","LALUA","SOLAR"],
            state="readonly", width=8, font=("Segoe UI",8))
        self._cap_filtro_emp.set("TODOS")
        self._cap_filtro_emp.pack(side="left", padx=(2,10))

        self._cap_busca_var = tk.StringVar()
        tk.Entry(tb, textvariable=self._cap_busca_var, font=("Segoe UI",8),
                 bg="#0a0f1e", fg=text, insertbackground=accent,
                 relief="flat", bd=2, width=20).pack(side="left", padx=(0,6))

        tk.Button(tb, text="🔍 Filtrar", font=("Segoe UI",8),
                  bg="#1e293b", fg=accent, relief="flat", bd=0,
                  padx=8, pady=3, cursor="hand2",
                  command=self._cap_carregar_lista).pack(side="left", padx=(0,6))

        tk.Button(tb, text="➕ Nova Nota", font=("Segoe UI",8,"bold"),
                  bg=accent, fg="#0f172a", relief="flat", bd=0,
                  padx=10, pady=3, cursor="hand2",
                  command=lambda: nb_pai.select(1)).pack(side="left", padx=(0,4))

        tk.Button(tb, text="💰 Adiantamento", font=("Segoe UI",8),
                  bg="#7c3aed", fg="white", relief="flat", bd=0,
                  padx=10, pady=3, cursor="hand2",
                  command=lambda: nb_pai.select(2)).pack(side="left")

        # Treeview
        cols = ("Nº TX","Tipo","Fornecedor","Empresa","Emissão",
                "Vencimento","Bruto","Líquido","Status","Categoria")
        self._cap_tree = ttk.Treeview(parent, columns=cols,
                                       show="headings", height=14)
        largs = [110,70,200,70,80,80,100,100,80,160]
        for c, w in zip(cols, largs):
            self._cap_tree.heading(c, text=c)
            self._cap_tree.column(c, width=w,
                anchor="e" if c in ("Bruto","Líquido") else "w")

        self._cap_tree.tag_configure("PENDENTE", foreground="#fbbf24")
        self._cap_tree.tag_configure("APROVADA", foreground="#38bdf8")
        self._cap_tree.tag_configure("PAGA",     foreground="#4ade80")
        self._cap_tree.tag_configure("VENCIDA",  foreground="#f87171")
        self._cap_tree.tag_configure("CANCELADA",foreground="#64748b")

        sb = ttk.Scrollbar(parent, orient="vertical", command=self._cap_tree.yview)
        self._cap_tree.configure(yscrollcommand=sb.set)

        # Rodapé com totais e ações
        rod = tk.Frame(parent, bg=surface, pady=5, padx=14)
        rod.pack(fill="x", side="bottom")

        self._cap_lbl_totais = tk.Label(rod, text="", font=("Consolas",8),
                                         fg=muted, bg=surface)
        self._cap_lbl_totais.pack(side="left")

        tk.Button(rod, text="✅ Marcar PAGA", font=("Segoe UI",8),
                  bg="#059669", fg="white", relief="flat", bd=0,
                  padx=10, pady=3, cursor="hand2",
                  command=lambda: self._cap_mudar_status("PAGA")).pack(side="right", padx=(4,0))
        tk.Button(rod, text="📋 Aprovar", font=("Segoe UI",8),
                  bg="#0ea5e9", fg="white", relief="flat", bd=0,
                  padx=10, pady=3, cursor="hand2",
                  command=lambda: self._cap_mudar_status("APROVADA")).pack(side="right", padx=(4,0))
        tk.Button(rod, text="❌ Cancelar", font=("Segoe UI",8),
                  bg="#7f1d1d", fg="#fca5a5", relief="flat", bd=0,
                  padx=10, pady=3, cursor="hand2",
                  command=lambda: self._cap_mudar_status("CANCELADA")).pack(side="right", padx=(4,0))

        self._cap_tree.pack(side="left", fill="both", expand=True, padx=(0,0))
        sb.pack(side="left", fill="y")

        self._cap_carregar_lista()

    def _cap_carregar_lista(self):
        for r in self._cap_tree.get_children():
            self._cap_tree.delete(r)

        st  = self._cap_filtro_status.get()
        emp = self._cap_filtro_emp.get()
        bsc = self._cap_busca_var.get().strip()

        notas = listar_notas(
            status=None if st=="TODOS" else st,
            empresa=None if emp=="TODOS" else emp,
            busca=bsc
        )

        total_bruto = 0.0
        total_liq   = 0.0
        for n in notas:
            tag = n["status"] if n["status"] in STATUS_NOTA else "PENDENTE"
            self._cap_tree.insert("", "end", iid=str(n["id"]), tags=(tag,),
                values=(
                    n["numero_tx"], n["tipo"], n["fornecedor"][:30],
                    n["empresa"], n["dt_emissao"], n["dt_vencimento"],
                    f"R$ {n['valor_bruto']:,.2f}",
                    f"R$ {n['valor_liquido']:,.2f}",
                    n["status"], n["categoria"][:25]
                ))
            total_bruto += n["valor_bruto"]
            total_liq   += n["valor_liquido"]

        self._cap_lbl_totais.config(
            text=f"{len(notas)} nota(s)  |  "
                 f"Bruto: R$ {total_bruto:,.2f}  |  "
                 f"Líquido: R$ {total_liq:,.2f}"
        )

    def _cap_mudar_status(self, novo_status):
        sel = self._cap_tree.selection()
        if not sel:
            return
        for iid in sel:
            atualizar_status_nota(int(iid), novo_status)
        self._cap_carregar_lista()

    # ── SUB-ABA: NOVA NOTA ────────────────────────────────────────────────────
    def _cap_build_nova_nota(self, parent, bg, surface, accent, green, yellow, text, muted, nb_pai):

        # Canvas + scrollbar para caber tudo
        canvas = tk.Canvas(parent, bg=bg, highlightthickness=0)
        sb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        frm = tk.Frame(canvas, bg=bg)
        win = canvas.create_window((0,0), window=frm, anchor="nw")

        def _resize(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(win, width=e.width)
        canvas.bind("<Configure>", _resize)
        frm.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))

        def _scroll(e):
            canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _scroll)

        pad = dict(padx=20, pady=4)

        # ── TÍTULO + BOTÃO XML ──
        frm_tit = tk.Frame(frm, bg=bg)
        frm_tit.pack(fill="x", padx=20, pady=4)
        tk.Label(frm_tit, text="➕  Nova Nota / Despesa",
                 font=("Segoe UI",12,"bold"), fg=accent, bg=bg
                 ).pack(side="left")

        tk.Button(frm_tit, text="📎 Importar XML (NF-e)",
                  font=("Segoe UI",9,"bold"), bg="#7c3aed", fg="white",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=lambda: self._cap_importar_xml()
                  ).pack(side="right")

        self._cap_lbl_xml = tk.Label(frm_tit, text="",
                                      font=("Segoe UI",8), fg=green, bg=bg)
        self._cap_lbl_xml.pack(side="right", padx=8)

        # ── CAMPO CHAVE DE ACESSO ──
        frm_chave = tk.Frame(frm, bg=surface, padx=14, pady=8)
        frm_chave.pack(fill="x", padx=20, pady=(0,6))

        tk.Label(frm_chave, text="🔑 Chave de Acesso NF-e:",
                 fg=accent, bg=surface, font=("Segoe UI",8,"bold")
                 ).pack(side="left", padx=(0,8))

        self._cap_chave_var = tk.StringVar()
        ent_chave = tk.Entry(frm_chave, textvariable=self._cap_chave_var,
                             font=("Consolas",9), bg="#0a0f1e", fg=yellow,
                             insertbackground=accent, relief="flat", bd=2, width=48)
        ent_chave.pack(side="left", padx=(0,8))
        tk.Label(frm_chave, text="(44 dígitos — pode bipar ou colar)",
                 fg="#475569", bg=surface, font=("Consolas",7)
                 ).pack(side="left", padx=(0,8))

        tk.Button(frm_chave, text="🔍 Consultar SEFAZ",
                  font=("Segoe UI",9,"bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=12, pady=4, cursor="hand2",
                  command=lambda: self._cap_consultar_sefaz()
                  ).pack(side="left")

        # ── BLOCO 1: DADOS BÁSICOS ──
        b1 = tk.LabelFrame(frm, text="  Dados da Nota  ", bg=surface,
                            fg=accent, font=("Segoe UI",9,"bold"), padx=12, pady=8)
        b1.pack(fill="x", padx=20, pady=(0,6))

        def lbl(parent, t): return tk.Label(parent, text=t, fg=muted, bg=surface,
                                             font=("Segoe UI",8), anchor="w")
        def ent(parent, w=25, **kw):
            e = tk.Entry(parent, font=("Segoe UI",9), bg="#0a0f1e", fg=text,
                         insertbackground=accent, relief="flat", bd=2, width=w, **kw)
            return e

        # Linha 0 — Tipo
        lbl(b1,"Tipo:").grid(row=0,column=0,sticky="w",pady=2)
        self._cap_tipo_var = tk.StringVar(value="NOTA")
        for i,(val,lbl_t) in enumerate([("NOTA","Nota Fiscal"),
                                         ("ADIANTAMENTO","Adiantamento")]):
            tk.Radiobutton(b1, text=lbl_t, variable=self._cap_tipo_var,
                           value=val, bg=surface, fg=text,
                           selectcolor="#0a0f1e",
                           font=("Segoe UI",8)).grid(row=0,column=i+1,sticky="w",padx=4)

        # Linha 1 — Fornecedor / CNPJ
        lbl(b1,"Fornecedor:").grid(row=1,column=0,sticky="w",pady=2)

        # Frame para campo + botão busca
        frm_forn = tk.Frame(b1, bg=surface)
        frm_forn.grid(row=1,column=1,columnspan=2,sticky="w",padx=(0,10))

        self._cap_forn = ent(frm_forn, 28)
        self._cap_forn.pack(side="left")

        tk.Button(frm_forn, text="🔍", font=("Segoe UI",8),
                  bg="#1e293b", fg=accent, relief="flat", bd=0,
                  padx=5, pady=2, cursor="hand2",
                  command=lambda: _abrir_busca_forn()
                  ).pack(side="left", padx=(3,0))

        lbl(b1,"CNPJ/CPF:").grid(row=1,column=3,sticky="w")
        self._cap_cnpj = ent(b1, 18)
        self._cap_cnpj.grid(row=1,column=4,sticky="w")

        def _aplicar_fornecedor(f):
            """Preenche todos os campos com dados do fornecedor."""
            self._cap_forn.delete(0,"end")
            self._cap_forn.insert(0, f["razao_social"])
            if f["cnpj_cpf"]:
                self._cap_cnpj.delete(0,"end")
                self._cap_cnpj.insert(0, f["cnpj_cpf"])
            if f["categoria"]:
                self._cap_cat_var.set(f["categoria"])
            if f["responsavel"]:
                self._cap_resp.set(f["responsavel"])
            if f["empresa"]:
                self._cap_emp_var.set(f["empresa"])
            if f.get("filial_padrao"):
                if self._cap_rat_rows and not self._cap_rat_rows[0]["filial"].get():
                    self._cap_rat_rows[0]["filial"].set(f["filial_padrao"])
                    if f["categoria"] and not self._cap_rat_rows[0]["cat"].get():
                        self._cap_rat_rows[0]["cat"].set(f["categoria"])
            # Preenche dados bancários se disponível
            if hasattr(self, "_cap_preencher_banco"):
                self._cap_preencher_banco(dict(f))

        def _abrir_busca_forn():
            """Abre popup de busca de fornecedor."""
            pop = tk.Toplevel(self.root)
            pop.title("Buscar Fornecedor")
            pop.geometry("600x400")
            pop.configure(bg="#0f172a")
            pop.grab_set()

            # Barra de busca
            tb_p = tk.Frame(pop, bg="#1e293b", pady=8, padx=12)
            tb_p.pack(fill="x")
            tk.Label(tb_p, text="Buscar:", fg="#94a3b8", bg="#1e293b",
                     font=("Segoe UI",9)).pack(side="left")
            bv = tk.StringVar()
            be = tk.Entry(tb_p, textvariable=bv, font=("Segoe UI",10),
                          bg="#0a0f1e", fg="#e2e8f0",
                          insertbackground="#38bdf8",
                          relief="flat", bd=2, width=30)
            be.pack(side="left", padx=(6,8))
            be.focus_set()

            # Treeview de resultados
            cols = ("Razão Social","CNPJ/CPF","Categoria","Resp.","Empresa")
            tree = ttk.Treeview(pop, columns=cols, show="headings", height=12)
            for c, w in zip(cols, [240,130,160,110,70]):
                tree.heading(c,text=c)
                tree.column(c,width=w,anchor="w")
            sb_p = ttk.Scrollbar(pop, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=sb_p.set)

            def _reload_pop(*args):
                for r in tree.get_children():
                    tree.delete(r)
                termo = bv.get().strip()
                for f in listar_fornecedores(busca=termo, ativo_only=True):
                    tree.insert("","end", iid=str(f["id"]), values=(
                        f["razao_social"], f["cnpj_cpf"] or "",
                        f["categoria"] or "", f["responsavel"] or "",
                        f["empresa"]
                    ))

            bv.trace_add("write", _reload_pop)
            _reload_pop()

            def _selecionar(event=None):
                sel = tree.selection()
                if not sel: return
                fid = int(sel[0])
                with _db_conn() as conn:
                    conn.row_factory = sqlite3.Row
                    f = conn.execute("SELECT * FROM fornecedores WHERE id=?",
                                     (fid,)).fetchone()
                if f:
                    _aplicar_fornecedor(dict(f))
                pop.destroy()

            tree.bind("<Double-1>", _selecionar)
            tree.bind("<Return>",   _selecionar)

            tk.Button(pop, text="✅ Selecionar",
                      font=("Segoe UI",9,"bold"), bg="#059669", fg="white",
                      relief="flat", bd=0, padx=14, pady=6, cursor="hand2",
                      command=_selecionar).pack(side="bottom", pady=8)

            tree.pack(side="left", fill="both", expand=True, padx=(12,0), pady=8)
            sb_p.pack(side="left", fill="y", pady=8, padx=(0,12))

        # Autocomplete por FocusOut como fallback (CNPJ exato)
        def _autocomplete_cnpj(e):
            cnpj = self._cap_cnpj.get().strip()
            if not cnpj: return
            f = buscar_fornecedor_por_cnpj(cnpj)
            if f:
                _aplicar_fornecedor(dict(f))

        self._cap_cnpj.bind("<FocusOut>", _autocomplete_cnpj)
        self._cap_cnpj.bind("<Return>",   _autocomplete_cnpj)

        # Linha 2 — Empresa / Categoria
        lbl(b1,"Empresa:").grid(row=2,column=0,sticky="w",pady=2)
        self._cap_emp_var = tk.StringVar(value="LALUA")
        ttk.Combobox(b1, textvariable=self._cap_emp_var,
                     values=["LALUA","SOLAR"], state="readonly",
                     width=8, font=("Segoe UI",8)).grid(row=2,column=1,sticky="w",padx=(0,10))

        lbl(b1,"Categoria:").grid(row=2,column=2,sticky="w")
        self._cap_cat_var = tk.StringVar()
        self._cap_cat_cb = ttk.Combobox(b1, textvariable=self._cap_cat_var,
                                          values=CATEGORIAS_LISTA,
                                          state="normal", width=32, font=("Segoe UI",8))
        self._cap_cat_cb.grid(row=2,column=3,columnspan=2,sticky="w")

        # Linha 3 — Responsável / Descrição
        lbl(b1,"Responsável:").grid(row=3,column=0,sticky="w",pady=2)
        self._cap_resp = ttk.Combobox(b1, values=RESPONSAVEIS_LISTA,
                                       state="readonly", width=15, font=("Segoe UI",8))
        self._cap_resp.grid(row=3,column=1,columnspan=2,sticky="w",padx=(0,10))

        lbl(b1,"Descrição:").grid(row=3,column=3,sticky="w")
        self._cap_desc = ent(b1, 32)
        self._cap_desc.grid(row=3,column=4,sticky="w")

        # Linha 4 — Datas / Valor
        lbl(b1,"Emissão:").grid(row=4,column=0,sticky="w",pady=2)
        self._cap_dt_emissao = ent(b1, 11)
        self._cap_dt_emissao.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self._cap_dt_emissao.grid(row=4,column=1,sticky="w",padx=(0,10))
        _aplicar_mascara_data(self._cap_dt_emissao)

        lbl(b1,"Vencimento:").grid(row=4,column=2,sticky="w")
        self._cap_dt_venc = ent(b1, 11)
        self._cap_dt_venc.grid(row=4,column=3,sticky="w",padx=(0,10))
        _aplicar_mascara_data(self._cap_dt_venc)

        lbl(b1,"Valor Bruto:").grid(row=4,column=4,sticky="w")
        self._cap_valor = ent(b1, 14)
        self._cap_valor.grid(row=4,column=5,sticky="w")
        _aplicar_mascara_valor(self._cap_valor)

        lbl(b1,"Observação:").grid(row=5,column=0,sticky="w",pady=2)
        self._cap_obs = ent(b1, 55)
        self._cap_obs.grid(row=5,column=1,columnspan=5,sticky="w")

        # Código de barras — campo largo, recebe leitura da pistola
        lbl(b1,"Cód. Barras:").grid(row=6,column=0,sticky="w",pady=2)
        frm_cb = tk.Frame(b1, bg=surface)
        frm_cb.grid(row=6,column=1,columnspan=5,sticky="w")
        self._cap_cod_barras = tk.Entry(
            frm_cb, font=("Consolas",9), bg="#0a0f1e", fg=yellow,
            insertbackground=accent, relief="flat", bd=2, width=52)
        self._cap_cod_barras.pack(side="left")
        tk.Label(frm_cb, text="← bipe aqui com a pistola",
                 fg="#475569", bg=surface,
                 font=("Consolas",7)).pack(side="left", padx=6)

        # Ao bipar, detecta automaticamente o tipo de pagamento
        def _detectar_tipo_barras(event=None):
            cb = self._cap_cod_barras.get().strip()
            if not cb: return
            # Identifica pelo primeiro dígito / banco
            if cb.startswith("858") or cb.startswith("859"):
                # DARF
                tipo_hint = "DARF"
            elif cb.startswith("856"):
                # GPS
                tipo_hint = "GPS"
            elif cb.startswith("858") and "DAS" in cb:
                tipo_hint = "DAS"
            elif len(cb) == 44 and cb[0] in "01234":
                tipo_hint = "Boleto"
            elif len(cb) in (44, 47, 48):
                tipo_hint = "Boleto"
            else:
                tipo_hint = ""
            if tipo_hint:
                # Atualiza descrição se estiver vazia
                if not self._cap_desc.get().strip():
                    self._cap_desc.delete(0,"end")
                    self._cap_desc.insert(0, tipo_hint)
        self._cap_cod_barras.bind("<Return>",   _detectar_tipo_barras)
        self._cap_cod_barras.bind("<FocusOut>", _detectar_tipo_barras)

        # ── BLOCO 1b: DADOS DE PAGAMENTO ──
        b1b = tk.LabelFrame(frm, text="  Dados de Pagamento  ", bg=surface,
                             fg=green, font=("Segoe UI",9,"bold"),
                             padx=12, pady=8)
        b1b.pack(fill="x", padx=20, pady=(0,6))

        # Forma de pagamento
        lbl(b1b,"Forma:").grid(row=0,column=0,sticky="w",pady=3)
        self._cap_forma_pgto = ttk.Combobox(b1b,
            values=["BOLETO","PIX","TED","DARF","GPS","DAS","CONCESSIONÁRIA"],
            state="readonly", width=16, font=("Segoe UI",9))
        self._cap_forma_pgto.set("BOLETO")
        self._cap_forma_pgto.grid(row=0,column=1,sticky="w",padx=(0,20))

        def _toggle_campos_pgto(event=None):
            forma = self._cap_forma_pgto.get()
            # Mostra/esconde campos conforme forma
            if forma in ("PIX","TED"):
                frm_boleto.grid_remove()
                frm_pix.grid()
                frm_banco.grid()
            elif forma in ("DARF","GPS","DAS","CONCESSIONÁRIA","BOLETO"):
                frm_pix.grid_remove()
                frm_banco.grid_remove()
                frm_boleto.grid()
            else:
                frm_pix.grid_remove()
                frm_boleto.grid_remove()
                frm_banco.grid_remove()
        self._cap_forma_pgto.bind("<<ComboboxSelected>>", _toggle_campos_pgto)

        # Linha PIX (escondida inicialmente)
        frm_pix = tk.Frame(b1b, bg=surface)
        frm_pix.grid(row=1,column=0,columnspan=6,sticky="w",pady=2)

        lbl(frm_pix,"Tipo chave:").pack(side="left",padx=(0,4))
        self._cap_pix_tipo = ttk.Combobox(frm_pix,
            values=["CPF","CNPJ","TELEFONE","EMAIL","ALEATÓRIA"],
            state="readonly", width=10, font=("Segoe UI",8))
        self._cap_pix_tipo.set("CNPJ")
        self._cap_pix_tipo.pack(side="left",padx=(0,10))

        lbl(frm_pix,"Chave PIX:").pack(side="left",padx=(0,4))
        self._cap_pix_chave = tk.Entry(frm_pix, font=("Consolas",9),
                                        bg="#0a0f1e", fg=green,
                                        insertbackground=accent,
                                        relief="flat", bd=2, width=35)
        self._cap_pix_chave.pack(side="left")
        frm_pix.grid_remove()

        # Linha dados bancários TED
        frm_banco = tk.Frame(b1b, bg=surface)
        frm_banco.grid(row=2,column=0,columnspan=6,sticky="w",pady=2)

        lbl(frm_banco,"Banco:").pack(side="left",padx=(0,4))
        self._cap_banco_dest = tk.Entry(frm_banco, font=("Segoe UI",8),
                                         bg="#0a0f1e", fg=text,
                                         insertbackground=accent,
                                         relief="flat", bd=2, width=8)
        self._cap_banco_dest.pack(side="left",padx=(0,10))

        lbl(frm_banco,"Agência:").pack(side="left",padx=(0,4))
        self._cap_ag_dest = tk.Entry(frm_banco, font=("Segoe UI",8),
                                      bg="#0a0f1e", fg=text,
                                      insertbackground=accent,
                                      relief="flat", bd=2, width=8)
        self._cap_ag_dest.pack(side="left",padx=(0,10))

        lbl(frm_banco,"Conta:").pack(side="left",padx=(0,4))
        self._cap_conta_dest = tk.Entry(frm_banco, font=("Segoe UI",8),
                                         bg="#0a0f1e", fg=text,
                                         insertbackground=accent,
                                         relief="flat", bd=2, width=12)
        self._cap_conta_dest.pack(side="left",padx=(0,10))

        lbl(frm_banco,"CPF/CNPJ dest:").pack(side="left",padx=(0,4))
        self._cap_cpf_cnpj_dest = tk.Entry(frm_banco, font=("Segoe UI",8),
                                             bg="#0a0f1e", fg=text,
                                             insertbackground=accent,
                                             relief="flat", bd=2, width=18)
        self._cap_cpf_cnpj_dest.pack(side="left")
        frm_banco.grid_remove()

        # Linha boleto (visível por padrão)
        frm_boleto = tk.Frame(b1b, bg=surface)
        frm_boleto.grid(row=3,column=0,columnspan=6,sticky="w",pady=2)
        tk.Label(frm_boleto,
                 text="ℹ️  Bipe o código de barras no campo acima. "
                      "Para PIX ou TED, selecione a forma e preencha os dados bancários.",
                 fg="#475569", bg=surface, font=("Consolas",7)
                 ).pack(side="left")

        # Auto-preenche dados bancários do fornecedor quando selecionado
        def _preencher_dados_banco(f_dict):
            """Chamado pelo autocomplete quando fornecedor é selecionado."""
            if f_dict.get("pix_chave"):
                self._cap_forma_pgto.set("PIX")
                self._cap_pix_chave.delete(0,"end")
                self._cap_pix_chave.insert(0, f_dict["pix_chave"])
                # Tipo da chave PIX baseado no formato
                chave = f_dict["pix_chave"]
                digits = "".join(c for c in chave if c.isdigit())
                if "@" in chave:
                    self._cap_pix_tipo.set("EMAIL")
                elif chave.startswith("+") or (chave.isdigit() and len(digits) == 11 and chave[0] in "67889"):
                    self._cap_pix_tipo.set("TELEFONE")
                elif len(digits) == 11:
                    self._cap_pix_tipo.set("CPF")
                elif len(digits) == 14:
                    self._cap_pix_tipo.set("CNPJ")
                else:
                    self._cap_pix_tipo.set("ALEATÓRIA")
                _toggle_campos_pgto()
            elif f_dict.get("conta"):
                self._cap_forma_pgto.set("TED")
                self._cap_banco_dest.delete(0,"end")
                self._cap_banco_dest.insert(0, f_dict.get("banco",""))
                self._cap_ag_dest.delete(0,"end")
                self._cap_ag_dest.insert(0, f_dict.get("agencia",""))
                self._cap_conta_dest.delete(0,"end")
                self._cap_conta_dest.insert(0, f_dict.get("conta",""))
                self._cap_cpf_cnpj_dest.delete(0,"end")
                self._cap_cpf_cnpj_dest.insert(0, f_dict.get("cnpj_cpf",""))
                _toggle_campos_pgto()

        # Guarda referência para o autocomplete usar
        self._cap_preencher_banco = _preencher_dados_banco

        # ── BLOCO 2: IMPOSTOS ──
        b2 = tk.LabelFrame(frm, text="  Impostos Retidos  ", bg=surface,
                            fg=yellow, font=("Segoe UI",9,"bold"), padx=12, pady=8)
        b2.pack(fill="x", padx=20, pady=(0,6))

        tk.Label(b2, text="Valor Bruto de referência é preenchido automaticamente ao digitar acima.",
                 fg="#475569", bg=surface, font=("Consolas",7)).pack(anchor="w")

        self._cap_imp_rows = []
        frm_imps = tk.Frame(b2, bg=surface)
        frm_imps.pack(fill="x")

        # Cabeçalho impostos
        for i, h in enumerate(["Imposto","Alíquota %","Valor R$","Venc. DARF/GPS"]):
            tk.Label(frm_imps, text=h, fg=accent, bg=surface,
                     font=("Segoe UI",7,"bold")).grid(row=0,column=i*2,sticky="w",padx=4)

        def _add_imposto(tipo="", aliq="", venc=""):
            row_i = len(self._cap_imp_rows) + 1
            tipo_var = tk.StringVar(value=tipo)
            aliq_var = tk.StringVar(value=str(aliq))
            val_var  = tk.StringVar(value="")
            venc_var = tk.StringVar(value=venc)

            cb = ttk.Combobox(frm_imps, textvariable=tipo_var,
                              values=TIPOS_IMPOSTO, state="readonly",
                              width=12, font=("Segoe UI",8))
            cb.grid(row=row_i, column=0, padx=4, pady=2)

            aliq_e = tk.Entry(frm_imps, textvariable=aliq_var, width=7,
                              font=("Segoe UI",8), bg="#0a0f1e", fg=text,
                              insertbackground=accent, relief="flat", bd=2)
            aliq_e.grid(row=row_i, column=2, padx=4)

            val_e = tk.Entry(frm_imps, textvariable=val_var, width=12,
                             font=("Segoe UI",8), bg="#0a0f1e", fg=yellow,
                             insertbackground=accent, relief="flat", bd=2)
            val_e.grid(row=row_i, column=4, padx=4)

            venc_e = tk.Entry(frm_imps, textvariable=venc_var, width=11,
                              font=("Segoe UI",8), bg="#0a0f1e", fg=text,
                              insertbackground=accent, relief="flat", bd=2)
            venc_e.grid(row=row_i, column=6, padx=4)

            # Auto-calcula valor quando alíquota ou tipo muda
            def _calc(*args):
                try:
                    vb = _parse_valor(self._cap_valor.get())
                    al = _parse_valor(aliq_var.get())
                    val_var.set(f"{vb * al / 100:.2f}")
                except: pass

            def _preenche_aliq(*args):
                tp = tipo_var.get()
                if tp in ALIQUOTAS_PADRAO and not aliq_var.get():
                    aliq_var.set(str(ALIQUOTAS_PADRAO[tp]))
                _calc()

            tipo_var.trace_add("write", _preenche_aliq)
            aliq_var.trace_add("write", _calc)
            self._cap_valor.bind("<FocusOut>", lambda e: [_calc() for _ in [1]])

            self._cap_imp_rows.append({
                "tipo": tipo_var, "aliq": aliq_var,
                "val": val_var,   "venc": venc_var
            })

        for tp in ["PIS","COFINS","IRRF"]:  # 3 linhas padrão
            _add_imposto(tp, ALIQUOTAS_PADRAO.get(tp,""))

        tk.Button(b2, text="+ imposto", font=("Segoe UI",7),
                  bg="#1e293b", fg=yellow, relief="flat", bd=0,
                  padx=6, pady=2, cursor="hand2",
                  command=_add_imposto).pack(anchor="w", pady=(4,0))

        # Label valor líquido
        self._cap_lbl_liq = tk.Label(b2, text="Valor líquido a pagar:  R$ 0,00",
                                      font=("Segoe UI",9,"bold"), fg=green, bg=surface)
        self._cap_lbl_liq.pack(anchor="e", pady=(4,0))

        def _atualiza_liquido(*args):
            try:
                vb = _parse_valor(self._cap_valor.get())
                ti = sum(float(r["val"].get() or 0) for r in self._cap_imp_rows)
                liq = vb - ti
                self._cap_lbl_liq.config(
                    text=f"Valor líquido a pagar:  R$ {liq:,.2f}")
            except: pass
        self._cap_valor.bind("<KeyRelease>", _atualiza_liquido)

        # ── BLOCO 3: PARCELAMENTO ──
        b3 = tk.LabelFrame(frm, text="  Parcelamento  ", bg=surface,
                            fg=accent, font=("Segoe UI",9,"bold"), padx=12, pady=8)
        b3.pack(fill="x", padx=20, pady=(0,6))

        frame_parc_ctrl = tk.Frame(b3, bg=surface)
        frame_parc_ctrl.pack(fill="x")

        tk.Label(frame_parc_ctrl, text="Nº parcelas:", fg=muted, bg=surface,
                 font=("Segoe UI",8)).pack(side="left")
        self._cap_nparc = tk.Spinbox(frame_parc_ctrl, from_=1, to=48, width=4,
                                      font=("Segoe UI",9), bg="#0a0f1e", fg=text,
                                      insertbackground=accent, relief="flat", bd=2)
        self._cap_nparc.pack(side="left", padx=(4,12))

        tk.Label(frame_parc_ctrl, text="1ª parcela:", fg=muted, bg=surface,
                 font=("Segoe UI",8)).pack(side="left")
        self._cap_dt_1parc = ent(frame_parc_ctrl, 11)
        self._cap_dt_1parc.pack(side="left", padx=(4,12))
        _aplicar_mascara_data(self._cap_dt_1parc)

        tk.Label(frame_parc_ctrl, text="Intervalo (dias):", fg=muted, bg=surface,
                 font=("Segoe UI",8)).pack(side="left")
        self._cap_intervalo = tk.Spinbox(frame_parc_ctrl, from_=1, to=90,
                                          width=4, font=("Segoe UI",9),
                                          bg="#0a0f1e", fg=text,
                                          insertbackground=accent, relief="flat", bd=2)
        self._cap_intervalo.delete(0,"end"); self._cap_intervalo.insert(0,"30")
        self._cap_intervalo.pack(side="left", padx=(4,12))

        tk.Button(frame_parc_ctrl, text="⚡ Calcular parcelas",
                  font=("Segoe UI",8), bg="#1e293b", fg=accent,
                  relief="flat", bd=0, padx=8, pady=3, cursor="hand2",
                  command=self._cap_calcular_parcelas).pack(side="left")

        self._cap_frm_parcelas = tk.Frame(b3, bg=surface)
        self._cap_frm_parcelas.pack(fill="x", pady=(6,0))
        self._cap_parc_widgets = []

        # ── BLOCO 4: RATEIO ──
        b4 = tk.LabelFrame(frm, text="  Rateio por Filial / Histórico Contábil  ",
                            bg=surface, fg=accent,
                            font=("Segoe UI",9,"bold"), padx=12, pady=8)
        b4.pack(fill="x", padx=20, pady=(0,6))

        tk.Label(b4, text="Soma dos percentuais deve ser 100%",
                 fg="#475569", bg=surface, font=("Consolas",7)).pack(anchor="w")

        self._cap_rat_rows = []
        frm_rats = tk.Frame(b4, bg=surface)
        frm_rats.pack(fill="x")

        for i,h in enumerate(["Filial","% Rateio","Categoria","Valor R$"]):
            tk.Label(frm_rats, text=h, fg=accent, bg=surface,
                     font=("Segoe UI",7,"bold")).grid(row=0,column=i*2,sticky="w",padx=4)

        def _add_rateio():
            ri = len(self._cap_rat_rows) + 1
            fil_var  = tk.StringVar()
            pct_var  = tk.StringVar()
            cat_var  = tk.StringVar()
            val_var  = tk.StringVar()

            ttk.Combobox(frm_rats, textvariable=fil_var,
                         values=FILIAIS_RATEIO, state="readonly",
                         width=16, font=("Segoe UI",8)
                         ).grid(row=ri, column=0, padx=4, pady=2)

            pct_e = tk.Entry(frm_rats, textvariable=pct_var, width=7,
                             font=("Segoe UI",8), bg="#0a0f1e", fg=text,
                             insertbackground=accent, relief="flat", bd=2)
            pct_e.grid(row=ri, column=2, padx=4)

            ttk.Combobox(frm_rats, textvariable=cat_var,
                         values=CATEGORIAS_LISTA, state="normal",
                         width=28, font=("Segoe UI",8)
                         ).grid(row=ri, column=4, padx=4)

            val_lbl = tk.Label(frm_rats, textvariable=val_var,
                               fg=yellow, bg=surface, font=("Consolas",8), width=12)
            val_lbl.grid(row=ri, column=6, padx=4)

            def _calc_rat(*args):
                try:
                    vb  = _parse_valor(self._cap_valor.get())
                    pct = _parse_valor(pct_var.get())
                    val_var.set(f"R$ {vb*pct/100:,.2f}")
                except: val_var.set("")
            pct_var.trace_add("write", _calc_rat)
            self._cap_valor.bind("<KeyRelease>",
                                  lambda e: [_calc_rat() for _ in [1]])

            self._cap_rat_rows.append({
                "filial":fil_var,"pct":pct_var,
                "cat":cat_var,"val":val_var
            })

        _add_rateio(); _add_rateio()  # 2 linhas iniciais

        tk.Button(b4, text="+ filial", font=("Segoe UI",7),
                  bg="#1e293b", fg=accent, relief="flat", bd=0,
                  padx=6, pady=2, cursor="hand2",
                  command=_add_rateio).pack(anchor="w", pady=(4,0))

        self._cap_lbl_pct_total = tk.Label(b4, text="Total rateio: 0%",
                                            font=("Segoe UI",8), fg=muted, bg=surface)
        self._cap_lbl_pct_total.pack(anchor="e")

        # ── BOTÃO SALVAR ──
        tk.Button(frm, text="💾  SALVAR NOTA",
                  font=("Segoe UI",11,"bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=24, pady=10, cursor="hand2",
                  command=lambda: self._cap_salvar_nota(nb_pai)
                  ).pack(padx=20, pady=(4,16), anchor="e")

        self._cap_lbl_save = tk.Label(frm, text="", font=("Segoe UI",9),
                                       fg=green, bg=bg)
        self._cap_lbl_save.pack(padx=20, anchor="e")

    def _cap_calcular_parcelas(self):
        """Gera as linhas de parcelas na tela."""
        for w in self._cap_parc_widgets:
            w.destroy()
        self._cap_parc_widgets = []

        try:
            n       = int(self._cap_nparc.get())
            vb      = _parse_valor(self._cap_valor.get())
            ti      = sum(float(r["val"].get() or 0) for r in self._cap_imp_rows)
            liq     = vb - ti
            dt_str  = self._cap_dt_1parc.get().strip()
            interv  = int(self._cap_intervalo.get())

            if not dt_str:
                dt_str = self._cap_dt_venc.get().strip()
            dt_base = datetime.strptime(dt_str, "%d/%m/%Y")
        except Exception as e:
            return

        valor_parc = round(liq / n, 2)
        resto      = round(liq - valor_parc * n, 2)

        self._cap_parc_data = []
        bg_s = "#0a0f1e"
        text_c = "white"

        hdr = tk.Frame(self._cap_frm_parcelas, bg="#1e293b")
        hdr.pack(fill="x", pady=(0,2))
        for h, w in [("Parcela",60),("Vencimento",110),("Valor",100)]:
            tk.Label(hdr, text=h, fg="#94a3b8", bg="#1e293b",
                     font=("Segoe UI",7,"bold"), width=w//7).pack(side="left", padx=4)
        self._cap_parc_widgets.append(hdr)

        from datetime import timedelta
        for i in range(n):
            dt_parc = dt_base + timedelta(days=interv*i)
            val_p   = valor_parc + (resto if i == n-1 else 0)

            row_f = tk.Frame(self._cap_frm_parcelas, bg="#111827")
            row_f.pack(fill="x", pady=1)

            tk.Label(row_f, text=f"{i+1}/{n}", fg="#94a3b8",
                     bg="#111827", font=("Segoe UI",8), width=6).pack(side="left", padx=4)

            dt_var = tk.StringVar(value=dt_parc.strftime("%d/%m/%Y"))
            tk.Entry(row_f, textvariable=dt_var, width=11,
                     font=("Segoe UI",8), bg=bg_s, fg=text_c,
                     insertbackground="#38bdf8", relief="flat", bd=2
                     ).pack(side="left", padx=4)

            val_var = tk.StringVar(value=f"{val_p:.2f}")
            tk.Entry(row_f, textvariable=val_var, width=10,
                     font=("Segoe UI",8), bg=bg_s, fg="#4ade80",
                     insertbackground="#38bdf8", relief="flat", bd=2
                     ).pack(side="left", padx=4)

            self._cap_parc_data.append({"dt": dt_var, "val": val_var,
                                         "num": i+1, "total": n})
            self._cap_parc_widgets.append(row_f)

    def _cap_salvar_nota(self, nb_pai):
        """Valida e salva a nota no banco SQLite."""
        try:
            forn = self._cap_forn.get().strip()
            if not forn:
                self._cap_lbl_save.config(text="⚠️ Informe o fornecedor.", fg="#fbbf24")
                return
            vb_str = self._cap_valor.get().strip()
            if not vb_str:
                self._cap_lbl_save.config(text="⚠️ Informe o valor.", fg="#fbbf24")
                return
            valor_bruto = _parse_valor(vb_str)

            impostos = []
            for r in self._cap_imp_rows:
                tp = r["tipo"].get().strip()
                if not tp: continue
                try: val = _parse_valor(r["val"].get())
                except: val = 0.0
                try: aliq = _parse_valor(r["aliq"].get())
                except: aliq = 0.0
                if val > 0:
                    impostos.append({
                        "tipo": tp, "aliquota": aliq,
                        "valor": val,
                        "dt_venc_imp": r["venc"].get().strip()
                    })

            parcelas = []
            if hasattr(self, "_cap_parc_data") and self._cap_parc_data:
                for p in self._cap_parc_data:
                    try: vp = _parse_valor(p["val"].get())
                    except: vp = 0.0
                    parcelas.append({
                        "numero": p["num"], "total": p["total"],
                        "dt_vencimento": p["dt"].get().strip(),
                        "valor": vp
                    })
            else:
                # Sem parcelamento — 1 parcela no vencimento
                ti = sum(i["valor"] for i in impostos)
                parcelas = [{
                    "numero": 1, "total": 1,
                    "dt_vencimento": self._cap_dt_venc.get().strip(),
                    "valor": valor_bruto - ti
                }]

            rateio = []
            total_pct = 0.0
            for r in self._cap_rat_rows:
                fil = r["filial"].get().strip()
                cat = r["cat"].get().strip()
                if not fil: continue
                try: pct = _parse_valor(r["pct"].get())
                except: pct = 0.0
                if pct > 0:
                    ti = sum(i["valor"] for i in impostos)
                    rateio.append({
                        "filial": fil, "categoria": cat,
                        "percentual": pct,
                        "valor": round((valor_bruto-ti)*pct/100, 2)
                    })
                    total_pct += pct

            dados = {
                "tipo":          self._cap_tipo_var.get(),
                "fornecedor":    forn,
                "cnpj":          self._cap_cnpj.get().strip(),
                "empresa":       self._cap_emp_var.get(),
                "descricao":     self._cap_desc.get().strip(),
                "dt_emissao":    self._cap_dt_emissao.get().strip(),
                "dt_vencimento": self._cap_dt_venc.get().strip(),
                "valor_bruto":   valor_bruto,
                "categoria":     self._cap_cat_var.get().strip(),
                "responsavel":   self._cap_resp.get().strip(),
                "observacao":    self._cap_obs.get().strip(),
                "cod_barras":    self._cap_cod_barras.get().strip(),
                "forma_pgto":    self._cap_forma_pgto.get().strip(),
                "banco_dest":    self._cap_banco_dest.get().strip(),
                "agencia_dest":  self._cap_ag_dest.get().strip(),
                "conta_dest":    self._cap_conta_dest.get().strip(),
                "pix_chave":     self._cap_pix_chave.get().strip(),
                "cpf_cnpj_dest": self._cap_cpf_cnpj_dest.get().strip(),
                "impostos":      impostos,
                "parcelas":      parcelas,
                "rateio":        rateio,
            }

            numero_tx = salvar_nota(dados)
            self._cap_lbl_save.config(
                text=f"✅ Nota salva com sucesso! Nº: {numero_tx}", fg="#4ade80")

            # Limpa campos
            for w in [self._cap_forn, self._cap_cnpj, self._cap_desc,
                       self._cap_valor, self._cap_obs, self._cap_dt_venc,
                       self._cap_cod_barras, self._cap_pix_chave,
                       self._cap_banco_dest, self._cap_ag_dest,
                       self._cap_conta_dest, self._cap_cpf_cnpj_dest]:
                w.delete(0,"end")
            self._cap_cat_var.set("")
            for r in self._cap_imp_rows:
                r["val"].set(""); r["aliq"].set("")
            for w in self._cap_parc_widgets:
                w.destroy()
            self._cap_parc_widgets = []
            if hasattr(self,"_cap_parc_data"):
                self._cap_parc_data = []

            # Volta para a lista e atualiza
            nb_pai.select(0)
            self._cap_carregar_lista()

        except Exception as e:
            self._cap_lbl_save.config(text=f"❌ Erro: {e}", fg="#f87171")

    # ── SUB-ABA: ADIANTAMENTOS ────────────────────────────────────────────────
    def _cap_build_adiantamentos(self, parent, bg, surface, accent, green, yellow, text, muted):

        tk.Label(parent, text="💰  Adiantamentos",
                 font=("Segoe UI",12,"bold"), fg=accent, bg=bg
                 ).pack(anchor="w", padx=20, pady=(12,4))

        # Formulário
        frm = tk.LabelFrame(parent, text="  Novo Adiantamento  ", bg=surface,
                             fg=accent, font=("Segoe UI",9,"bold"), padx=14, pady=10)
        frm.pack(fill="x", padx=20, pady=(0,8))

        def lbl(t): return tk.Label(frm, text=t, fg="#94a3b8", bg=surface,
                                     font=("Segoe UI",8))
        def ent(w=20):
            return tk.Entry(frm, font=("Segoe UI",9), bg="#0a0f1e", fg=text,
                            insertbackground=accent, relief="flat", bd=2, width=w)

        lbl("Beneficiário:").grid(row=0,column=0,sticky="w",pady=3)
        self._ad_benef = ent(28)
        self._ad_benef.grid(row=0,column=1,padx=(0,12),sticky="w")

        lbl("CPF/CNPJ:").grid(row=0,column=2,sticky="w")
        self._ad_cpf = ent(18)
        self._ad_cpf.grid(row=0,column=3,sticky="w")

        lbl("Tipo:").grid(row=1,column=0,sticky="w",pady=3)
        self._ad_tipo = ttk.Combobox(frm,
            values=["FUNCIONARIO","FORNECEDOR"],
            state="readonly", width=13, font=("Segoe UI",8))
        self._ad_tipo.set("FUNCIONARIO")
        self._ad_tipo.grid(row=1,column=1,sticky="w",padx=(0,12))

        lbl("Empresa:").grid(row=1,column=2,sticky="w")
        self._ad_emp = ttk.Combobox(frm,
            values=["LALUA","SOLAR"],
            state="readonly", width=8, font=("Segoe UI",8))
        self._ad_emp.set("LALUA")
        self._ad_emp.grid(row=1,column=3,sticky="w")

        lbl("Finalidade:").grid(row=2,column=0,sticky="w",pady=3)
        self._ad_final = ent(40)
        self._ad_final.grid(row=2,column=1,columnspan=3,sticky="w")

        lbl("Valor:").grid(row=3,column=0,sticky="w",pady=3)
        self._ad_valor = ent(14)
        self._ad_valor.grid(row=3,column=1,sticky="w",padx=(0,12))
        _aplicar_mascara_valor(self._ad_valor)

        lbl("Data:").grid(row=3,column=2,sticky="w")
        self._ad_dt = ent(11)
        self._ad_dt.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self._ad_dt.grid(row=3,column=3,sticky="w")
        _aplicar_mascara_data(self._ad_dt)

        self._ad_lbl = tk.Label(frm, text="", fg=green, bg=surface, font=("Segoe UI",9))
        self._ad_lbl.grid(row=4,column=0,columnspan=4,sticky="w",pady=4)

        def _salvar_ad():
            try:
                benef = self._ad_benef.get().strip()
                valor = _parse_valor(self._ad_valor.get())
                if not benef:
                    self._ad_lbl.config(text="⚠️ Informe o beneficiário.", fg="#fbbf24")
                    return
                tx = salvar_adiantamento({
                    "beneficiario": benef,
                    "cpf_cnpj":     self._ad_cpf.get().strip(),
                    "tipo":         self._ad_tipo.get(),
                    "finalidade":   self._ad_final.get().strip(),
                    "empresa":      self._ad_emp.get(),
                    "valor":        valor,
                    "dt_adiantamento": self._ad_dt.get().strip(),
                })
                self._ad_lbl.config(text=f"✅ Adiantamento salvo: {tx}", fg=green)
                self._ad_benef.delete(0,"end")
                self._ad_valor.delete(0,"end")
                self._ad_final.delete(0,"end")
                _reload_ad()
            except Exception as e:
                self._ad_lbl.config(text=f"❌ {e}", fg="#f87171")

        tk.Button(frm, text="💾 Salvar Adiantamento",
                  font=("Segoe UI",9,"bold"), bg="#7c3aed", fg="white",
                  relief="flat", bd=0, padx=14, pady=6, cursor="hand2",
                  command=_salvar_ad).grid(row=5,column=0,columnspan=4,
                                            sticky="e", pady=(4,0))

        # Lista de adiantamentos
        tk.Label(parent, text="📋 Adiantamentos em aberto",
                 font=("Segoe UI",9,"bold"), fg=text, bg=bg
                 ).pack(anchor="w", padx=20, pady=(4,2))

        cols = ("Nº TX","Beneficiário","Tipo","Empresa","Data","Valor","Saldo","Status","Finalidade")
        self._ad_tree = ttk.Treeview(parent, columns=cols, show="headings", height=8)
        largs = [100,180,100,70,90,100,100,80,200]
        for c,w in zip(cols,largs):
            self._ad_tree.heading(c,text=c)
            self._ad_tree.column(c,width=w,
                anchor="e" if c in ("Valor","Saldo") else "w")
        self._ad_tree.tag_configure("ABERTO",  foreground="#fbbf24")
        self._ad_tree.tag_configure("PARCIAL", foreground="#38bdf8")
        self._ad_tree.tag_configure("QUITADO", foreground="#4ade80")

        sb_ad = ttk.Scrollbar(parent, orient="vertical", command=self._ad_tree.yview)
        self._ad_tree.configure(yscrollcommand=sb_ad.set)
        self._ad_tree.pack(side="left", fill="both", expand=True, padx=(20,0), pady=(0,10))
        sb_ad.pack(side="left", fill="y", pady=(0,10))

        def _reload_ad():
            for r in self._ad_tree.get_children():
                self._ad_tree.delete(r)
            for a in listar_adiantamentos():
                self._ad_tree.insert("", "end", iid=str(a["id"]),
                    tags=(a["status"],),
                    values=(a["numero_tx"], a["beneficiario"],
                            a["tipo"], a["empresa"],
                            a["dt_adiantamento"],
                            f"R$ {a['valor']:,.2f}",
                            f"R$ {a['saldo_aberto']:,.2f}",
                            a["status"], a["finalidade"]))

        self._ad_reload = _reload_ad
        _reload_ad()

    # ── SUB-ABA: PRESTAÇÃO DE CONTAS ─────────────────────────────────────────
    def _cap_build_prestacao(self, parent, bg, surface, accent, green, yellow, text, muted):

        tk.Label(parent, text="✅  Prestação de Contas",
                 font=("Segoe UI",12,"bold"), fg=accent, bg=bg
                 ).pack(anchor="w", padx=20, pady=(12,4))
        tk.Label(parent, text="Selecione o adiantamento e vincule as despesas realizadas.",
                 font=("Segoe UI",9), fg="#64748b", bg=bg
                 ).pack(anchor="w", padx=20, pady=(0,8))

        frm = tk.LabelFrame(parent, text="  Registrar Prestação  ", bg=surface,
                             fg=accent, font=("Segoe UI",9,"bold"), padx=14, pady=10)
        frm.pack(fill="x", padx=20, pady=(0,8))

        def lbl(t): return tk.Label(frm, text=t, fg="#94a3b8", bg=surface,
                                     font=("Segoe UI",8))
        def ent(w=20):
            return tk.Entry(frm, font=("Segoe UI",9), bg="#0a0f1e", fg=text,
                            insertbackground=accent, relief="flat", bd=2, width=w)

        lbl("Adiantamento (Nº TX):").grid(row=0,column=0,sticky="w",pady=3)
        self._pr_ad_var = tk.StringVar()
        ads = listar_adiantamentos(status="ABERTO")
        opts = [f"{a['numero_tx']} — {a['beneficiario']} (R$ {a['saldo_aberto']:,.2f})"
                for a in ads] + \
               [f"{a['numero_tx']} — {a['beneficiario']} (R$ {a['saldo_aberto']:,.2f})"
                for a in listar_adiantamentos(status="PARCIAL")]
        self._pr_ad_ids = {f"{a['numero_tx']} — {a['beneficiario']} (R$ {a['saldo_aberto']:,.2f})":
                            a["id"] for a in ads}
        self._pr_ad_ids.update({
            f"{a['numero_tx']} — {a['beneficiario']} (R$ {a['saldo_aberto']:,.2f})":
            a["id"] for a in listar_adiantamentos(status="PARCIAL")
        })
        ttk.Combobox(frm, textvariable=self._pr_ad_var,
                     values=opts, state="readonly",
                     width=45, font=("Segoe UI",8)
                     ).grid(row=0,column=1,columnspan=3,sticky="w",padx=(0,8))

        lbl("Descrição despesa:").grid(row=1,column=0,sticky="w",pady=3)
        self._pr_desc = ent(35)
        self._pr_desc.grid(row=1,column=1,columnspan=2,sticky="w",padx=(0,8))

        lbl("Nº Nota (opcional):").grid(row=1,column=3,sticky="w")
        self._pr_nota_var = tk.StringVar()
        ttk.Combobox(frm, textvariable=self._pr_nota_var,
                     values=[f"{n['numero_tx']} - {n['fornecedor']}"
                             for n in listar_notas(status="PENDENTE")],
                     state="normal", width=22, font=("Segoe UI",8)
                     ).grid(row=1,column=4,sticky="w")

        lbl("Valor aplicado:").grid(row=2,column=0,sticky="w",pady=3)
        self._pr_val_ap = ent(12)
        self._pr_val_ap.grid(row=2,column=1,sticky="w",padx=(0,12))

        lbl("Valor devolvido:").grid(row=2,column=2,sticky="w")
        self._pr_val_dev = ent(12)
        self._pr_val_dev.insert(0,"0,00")
        self._pr_val_dev.grid(row=2,column=3,sticky="w",padx=(0,12))

        lbl("Observação:").grid(row=3,column=0,sticky="w",pady=3)
        self._pr_obs = ent(50)
        self._pr_obs.grid(row=3,column=1,columnspan=4,sticky="w")

        self._pr_lbl = tk.Label(frm, text="", fg=green, bg=surface, font=("Segoe UI",9))
        self._pr_lbl.grid(row=4,column=0,columnspan=5,sticky="w",pady=4)

        def _salvar_pr():
            try:
                sel = self._pr_ad_var.get()
                if not sel:
                    self._pr_lbl.config(text="⚠️ Selecione o adiantamento.", fg="#fbbf24")
                    return
                ad_id = self._pr_ad_ids.get(sel)
                val_ap  = _parse_valor(self._pr_val_ap.get())
                val_dev = _parse_valor(self._pr_val_dev.get() or "0")

                nota_id = None
                nota_str = self._pr_nota_var.get().strip()
                if nota_str:
                    tx = nota_str.split(" - ")[0].strip()
                    with _db_conn() as conn:
                        row = conn.execute(
                            "SELECT id FROM notas WHERE numero_tx=?", (tx,)
                        ).fetchone()
                        if row: nota_id = row[0]

                registrar_prestacao(ad_id, nota_id, val_ap, val_dev,
                                     self._pr_desc.get().strip(),
                                     self._pr_obs.get().strip())
                self._pr_lbl.config(text="✅ Prestação registrada!", fg=green)
                for w in [self._pr_desc, self._pr_val_ap, self._pr_obs]:
                    w.delete(0,"end")
                self._pr_val_dev.delete(0,"end")
                self._pr_val_dev.insert(0,"0,00")
            except Exception as e:
                self._pr_lbl.config(text=f"❌ {e}", fg="#f87171")

        tk.Button(frm, text="✅ Registrar Prestação",
                  font=("Segoe UI",9,"bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=14, pady=6, cursor="hand2",
                  command=_salvar_pr).grid(row=5,column=0,columnspan=5,
                                            sticky="e", pady=(4,0))

    # ── SUB-ABA: IMPOSTOS RETIDOS ─────────────────────────────────────────────
    def _cap_build_impostos(self, parent, bg, surface, accent, green, yellow, text, muted):

        tk.Label(parent, text="🏛️  Impostos Retidos — Conferência e Baixa",
                 font=("Segoe UI",12,"bold"), fg=accent, bg=bg
                 ).pack(anchor="w", padx=20, pady=(12,4))
        tk.Label(parent,
                 text="Todos os impostos destacados nas notas aparecem aqui para conferência e baixa.",
                 font=("Segoe UI",9), fg="#64748b", bg=bg
                 ).pack(anchor="w", padx=20, pady=(0,8))

        # Toolbar
        tb = tk.Frame(parent, bg=surface, pady=6, padx=14)
        tb.pack(fill="x", padx=0)

        tk.Label(tb, text="Status:", fg=muted, bg=surface,
                 font=("Segoe UI",8)).pack(side="left")
        self._imp_filtro = ttk.Combobox(tb,
            values=["TODOS","PENDENTE","PAGO"],
            state="readonly", width=9, font=("Segoe UI",8))
        self._imp_filtro.set("PENDENTE")
        self._imp_filtro.pack(side="left", padx=(2,10))

        tk.Button(tb, text="🔍 Atualizar",
                  font=("Segoe UI",8), bg="#1e293b", fg=accent,
                  relief="flat", bd=0, padx=8, pady=3, cursor="hand2",
                  command=self._imp_carregar).pack(side="left", padx=(0,8))

        tk.Button(tb, text="✅ Marcar PAGO",
                  font=("Segoe UI",8,"bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=10, pady=3, cursor="hand2",
                  command=self._imp_marcar_pago).pack(side="right")

        # Treeview
        cols = ("ID Imp","Nº TX","Fornecedor","Tipo Imp","Alíquota",
                "Valor","Venc. DARF","Status","Nº Doc")
        self._imp_tree = ttk.Treeview(parent, columns=cols,
                                       show="headings", height=14)
        largs = [60,110,200,100,80,100,100,80,120]
        for c,w in zip(cols,largs):
            self._imp_tree.heading(c,text=c)
            self._imp_tree.column(c,width=w,
                anchor="e" if c in ("Valor","Alíquota") else "w")
        self._imp_tree.tag_configure("PENDENTE", foreground="#fbbf24")
        self._imp_tree.tag_configure("PAGO",     foreground="#4ade80")

        sb_i = ttk.Scrollbar(parent, orient="vertical", command=self._imp_tree.yview)
        self._imp_tree.configure(yscrollcommand=sb_i.set)

        # Rodapé totais
        rod = tk.Frame(parent, bg=surface, pady=4, padx=14)
        rod.pack(fill="x", side="bottom")
        self._imp_lbl_total = tk.Label(rod, text="", font=("Consolas",8),
                                        fg=muted, bg=surface)
        self._imp_lbl_total.pack(side="left")

        self._imp_tree.pack(side="left", fill="both", expand=True,
                             padx=(20,0), pady=(0,0))
        sb_i.pack(side="left", fill="y")
        self._imp_carregar()

    def _imp_carregar(self):
        for r in self._imp_tree.get_children():
            self._imp_tree.delete(r)

        st = self._imp_filtro.get()
        sql = """
            SELECT ni.id, n.numero_tx, n.fornecedor,
                   ni.tipo, ni.aliquota, ni.valor,
                   ni.dt_venc_imp, ni.status, ni.numero_doc
            FROM nota_impostos ni
            JOIN notas n ON n.id = ni.nota_id
            WHERE 1=1
        """
        params = []
        if st != "TODOS":
            sql += " AND ni.status=?"; params.append(st)
        sql += " ORDER BY ni.dt_venc_imp ASC"

        total = 0.0
        with _db_conn() as conn:
            for row in conn.execute(sql, params).fetchall():
                tag = row[7] if row[7] in ("PENDENTE","PAGO") else "PENDENTE"
                self._imp_tree.insert("", "end", iid=str(row[0]), tags=(tag,),
                    values=(row[0], row[1], row[2][:28],
                            row[3], f"{row[4]:.2f}%",
                            f"R$ {row[5]:,.2f}",
                            row[6] or "", row[7], row[8] or ""))
                if tag == "PENDENTE":
                    total += row[5]

        self._imp_lbl_total.config(
            text=f"Total pendente: R$ {total:,.2f}")

    def _imp_marcar_pago(self):
        sel = self._imp_tree.selection()
        if not sel: return
        with _db_conn() as conn:
            for iid in sel:
                conn.execute(
                    "UPDATE nota_impostos SET status='PAGO' WHERE id=?",
                    (int(iid),))
        self._imp_carregar()

    # ── SUB-ABA: FORNECEDORES ────────────────────────────────────────────────
    def _cap_build_fornecedores(self, parent, bg, surface, accent, green, yellow, text, muted):

        # ── TOOLBAR ──
        tb = tk.Frame(parent, bg=surface, pady=6, padx=14)
        tb.pack(fill="x")

        tk.Label(tb, text="🏢 Cadastro de Fornecedores",
                 font=("Segoe UI",12,"bold"), fg=accent, bg=surface
                 ).pack(side="left", padx=(0,16))

        self._forn_busca_var = tk.StringVar()
        tk.Entry(tb, textvariable=self._forn_busca_var,
                 font=("Segoe UI",8), bg="#0a0f1e", fg=text,
                 insertbackground=accent, relief="flat", bd=2, width=22
                 ).pack(side="left", padx=(0,4))

        tk.Button(tb, text="🔍", font=("Segoe UI",9),
                  bg="#1e293b", fg=accent, relief="flat", bd=0,
                  padx=6, pady=3, cursor="hand2",
                  command=self._forn_carregar).pack(side="left", padx=(0,10))

        self._forn_filtro_ativo = tk.BooleanVar(value=True)
        tk.Checkbutton(tb, text="Só ativos", variable=self._forn_filtro_ativo,
                       bg=surface, fg=muted, font=("Segoe UI",8),
                       selectcolor="#0a0f1e",
                       command=self._forn_carregar).pack(side="left", padx=(0,10))

        tk.Button(tb, text="➕ Novo Fornecedor",
                  font=("Segoe UI",8,"bold"), bg=accent, fg="#0f172a",
                  relief="flat", bd=0, padx=10, pady=3, cursor="hand2",
                  command=self._forn_novo).pack(side="right")

        # ── TREEVIEW ──
        cols = ("ID","Razão Social","Nome Fantasia","CNPJ/CPF","Tipo",
                "Empresa","Categoria","Responsável","PIX","Contato")
        self._forn_tree = ttk.Treeview(parent, columns=cols,
                                        show="headings", height=12)
        largs = [40,220,140,130,90,70,170,120,160,130]
        for c, w in zip(cols, largs):
            self._forn_tree.heading(c, text=c)
            self._forn_tree.column(c, width=w, anchor="w")

        self._forn_tree.tag_configure("ativo",   foreground="#e2e8f0")
        self._forn_tree.tag_configure("inativo",  foreground="#475569")

        sb = ttk.Scrollbar(parent, orient="vertical", command=self._forn_tree.yview)
        sbx = ttk.Scrollbar(parent, orient="horizontal", command=self._forn_tree.xview)
        self._forn_tree.configure(yscrollcommand=sb.set, xscrollcommand=sbx.set)

        self._forn_tree.bind("<Double-1>", self._forn_editar)

        # ── RODAPÉ ──
        rod = tk.Frame(parent, bg=surface, pady=5, padx=14)
        rod.pack(fill="x", side="bottom")
        self._forn_lbl_count = tk.Label(rod, text="", font=("Consolas",8),
                                         fg=muted, bg=surface)
        self._forn_lbl_count.pack(side="left")
        tk.Button(rod, text="✏️ Editar", font=("Segoe UI",8),
                  bg="#1e293b", fg=accent, relief="flat", bd=0,
                  padx=8, pady=3, cursor="hand2",
                  command=self._forn_editar).pack(side="right", padx=(4,0))
        tk.Button(rod, text="🚫 Inativar", font=("Segoe UI",8),
                  bg="#7f1d1d", fg="#fca5a5", relief="flat", bd=0,
                  padx=8, pady=3, cursor="hand2",
                  command=self._forn_inativar).pack(side="right", padx=(4,0))
        tk.Button(rod, text="🗑 Excluir", font=("Segoe UI",8),
                  bg="#450a0a", fg="#fca5a5", relief="flat", bd=0,
                  padx=8, pady=3, cursor="hand2",
                  command=self._forn_excluir).pack(side="right", padx=(4,0))

        sbx.pack(side="bottom", fill="x", padx=20)
        self._forn_tree.pack(side="left", fill="both", expand=True, padx=(20,0))
        sb.pack(side="left", fill="y")

        self._forn_carregar()

    def _forn_carregar(self):
        for r in self._forn_tree.get_children():
            self._forn_tree.delete(r)

        busca   = self._forn_busca_var.get().strip()
        ativos  = self._forn_filtro_ativo.get()
        fornecs = listar_fornecedores(busca=busca, ativo_only=ativos)

        for f in fornecs:
            tag = "ativo" if f["ativo"] else "inativo"
            self._forn_tree.insert("", "end", iid=str(f["id"]), tags=(tag,),
                values=(f["id"], f["razao_social"], f["nome_fantasia"] or "",
                        f["cnpj_cpf"] or "", f["tipo"], f["empresa"],
                        f["categoria"] or "", f["responsavel"] or "",
                        f["pix_chave"] or "", f["contato"] or ""))

        self._forn_lbl_count.config(text=f"{len(fornecs)} fornecedor(es)")

    def _forn_novo(self):
        self._forn_abrir_form(None)

    def _forn_editar(self, event=None):
        sel = self._forn_tree.selection()
        if not sel: return
        forn_id = int(sel[0])
        with _db_conn() as conn:
            conn.row_factory = sqlite3.Row
            f = conn.execute("SELECT * FROM fornecedores WHERE id=?",
                             (forn_id,)).fetchone()
        if f:
            self._forn_abrir_form(dict(f))

    def _forn_inativar(self):
        sel = self._forn_tree.selection()
        if not sel: return
        with _db_conn() as conn:
            for iid in sel:
                conn.execute("UPDATE fornecedores SET ativo=0 WHERE id=?",
                             (int(iid),))
        self._forn_carregar()

    def _forn_excluir(self):
        sel = self._forn_tree.selection()
        if not sel: return

        # Pega nome(s) para mostrar na confirmação
        nomes = []
        for iid in sel:
            vals = self._forn_tree.item(iid, "values")
            nomes.append(vals[1] if vals else f"ID {iid}")

        from tkinter import messagebox
        msg = f"Excluir permanentemente:\n\n" + "\n".join(nomes)
        if len(nomes) > 1:
            msg += f"\n\n({len(nomes)} fornecedores)"
        msg += "\n\nEsta ação não pode ser desfeita."

        if not messagebox.askyesno("Confirmar exclusão", msg,
                                    icon="warning", default="no"):
            return

        with _db_conn() as conn:
            for iid in sel:
                conn.execute("DELETE FROM fornecedores WHERE id=?", (int(iid),))
        self._forn_carregar()

    def _forn_abrir_form(self, dados=None):
        """Abre janela modal para cadastro/edição de fornecedor."""
        win = tk.Toplevel(self.root)
        win.title("Fornecedor" if not dados else f"Editar — {dados.get('razao_social','')}")
        win.geometry("620x580")
        win.configure(bg="#0f172a")
        win.grab_set()

        bg    = "#0f172a"
        surf  = "#1e293b"
        acc   = "#38bdf8"
        txt   = "#e2e8f0"
        mut   = "#94a3b8"

        tk.Label(win, text="🏢  Cadastro de Fornecedor",
                 font=("Segoe UI",12,"bold"), fg=acc, bg=bg
                 ).pack(anchor="w", padx=20, pady=(14,8))

        frm = tk.Frame(win, bg=surf, padx=16, pady=12)
        frm.pack(fill="x", padx=20)

        def lbl(t, r, c):
            tk.Label(frm, text=t, fg=mut, bg=surf,
                     font=("Segoe UI",8), anchor="w"
                     ).grid(row=r, column=c, sticky="w", pady=3, padx=(0,4))

        def ent(r, c, w=25, val="", cs=1):
            e = tk.Entry(frm, font=("Segoe UI",9), bg="#0a0f1e", fg=txt,
                         insertbackground=acc, relief="flat", bd=2, width=w)
            e.grid(row=r, column=c, columnspan=cs, sticky="w", padx=(0,8))
            if val: e.insert(0, val)
            return e

        def cbx(r, c, values, val="", w=14):
            v = tk.StringVar(value=val)
            cb = ttk.Combobox(frm, textvariable=v, values=values,
                              state="readonly", width=w, font=("Segoe UI",8))
            cb.grid(row=r, column=c, sticky="w", padx=(0,8), pady=3)
            return v

        d = dados or {}

        lbl("Razão Social *", 0, 0)
        e_rs = ent(0, 1, 30, d.get("razao_social",""), cs=3)

        lbl("Nome Fantasia", 1, 0)
        e_nf = ent(1, 1, 30, d.get("nome_fantasia",""), cs=3)

        lbl("CNPJ/CPF *", 2, 0)
        e_cnpj = ent(2, 1, 18, d.get("cnpj_cpf",""))

        lbl("Tipo", 2, 2)
        v_tipo = cbx(2, 3, ["FORNECEDOR","PRESTADOR","FUNCIONARIO"],
                     d.get("tipo","FORNECEDOR"), w=13)

        lbl("Empresa", 3, 0)
        v_emp = cbx(3, 1, ["LALUA","SOLAR"], d.get("empresa","LALUA"), w=8)

        lbl("Filial Padrão", 3, 2)
        v_filial = tk.StringVar(value=d.get("filial_padrao",""))
        ttk.Combobox(frm, textvariable=v_filial, values=FILIAIS_RATEIO,
                     state="normal", width=16, font=("Segoe UI",8)
                     ).grid(row=3, column=3, sticky="w", padx=(0,8), pady=3)

        lbl("Categoria Padrão", 4, 0)
        v_cat = tk.StringVar(value=d.get("categoria",""))
        ttk.Combobox(frm, textvariable=v_cat, values=CATEGORIAS_LISTA,
                     state="normal", width=32, font=("Segoe UI",8)
                     ).grid(row=4, column=1, columnspan=3, sticky="w", pady=3)

        lbl("Responsável Padrão", 5, 0)
        v_resp = tk.StringVar(value=d.get("responsavel",""))
        ttk.Combobox(frm, textvariable=v_resp, values=RESPONSAVEIS_LISTA,
                     state="readonly", width=16, font=("Segoe UI",8)
                     ).grid(row=5, column=1, sticky="w", pady=3)

        lbl("Contato", 5, 2)
        e_contato = ent(5, 3, 20, d.get("contato",""))

        lbl("E-mail", 6, 0)
        e_email = ent(6, 1, 28, d.get("email",""), cs=3)

        # Dados bancários
        tk.Label(frm, text="── Dados Bancários ──", fg="#475569", bg=surf,
                 font=("Segoe UI",7)).grid(row=7, column=0, columnspan=4,
                                            sticky="w", pady=(8,2))

        lbl("Banco", 8, 0)
        e_banco = ent(8, 1, 20, d.get("banco",""))

        lbl("Agência", 8, 2)
        e_ag = ent(8, 3, 10, d.get("agencia",""))

        lbl("Conta", 9, 0)
        e_ct = ent(9, 1, 16, d.get("conta",""))

        lbl("Chave PIX", 9, 2)
        e_pix = ent(9, 3, 22, d.get("pix_chave",""))

        # Status
        v_ativo = tk.BooleanVar(value=bool(d.get("ativo", 1)))
        tk.Checkbutton(frm, text="Fornecedor ativo",
                       variable=v_ativo, bg=surf, fg=txt,
                       selectcolor="#0a0f1e", font=("Segoe UI",8)
                       ).grid(row=10, column=0, columnspan=2,
                               sticky="w", pady=(8,0))

        lbl_save = tk.Label(win, text="", font=("Segoe UI",9),
                            fg=acc, bg=bg)
        lbl_save.pack(anchor="w", padx=20, pady=4)

        def _salvar():
            rs = e_rs.get().strip()
            if not rs:
                lbl_save.config(text="⚠️ Razão Social obrigatória.", fg="#fbbf24")
                return
            payload = {
                "razao_social":  rs,
                "nome_fantasia": e_nf.get().strip(),
                "cnpj_cpf":      e_cnpj.get().strip(),
                "tipo":          v_tipo.get(),
                "empresa":       v_emp.get(),
                "filial_padrao": v_filial.get(),
                "categoria":     v_cat.get(),
                "responsavel":   v_resp.get(),
                "contato":       e_contato.get().strip(),
                "email":         e_email.get().strip(),
                "banco":         e_banco.get().strip(),
                "agencia":       e_ag.get().strip(),
                "conta":         e_ct.get().strip(),
                "pix_chave":     e_pix.get().strip(),
                "ativo":         int(v_ativo.get()),
            }
            forn_id = dados.get("id") if dados else None
            salvar_fornecedor(payload, forn_id)
            lbl_save.config(text="✅ Salvo com sucesso!", fg="#4ade80")
            self._forn_carregar()
            win.after(800, win.destroy)

        btn_f = tk.Frame(win, bg=bg)
        btn_f.pack(fill="x", padx=20, pady=(0,12))
        tk.Button(btn_f, text="💾 Salvar",
                  font=("Segoe UI",10,"bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=16, pady=7, cursor="hand2",
                  command=_salvar).pack(side="right", padx=(6,0))
        tk.Button(btn_f, text="Cancelar",
                  font=("Segoe UI",9), bg=surf, fg=mut,
                  relief="flat", bd=0, padx=12, pady=7, cursor="hand2",
                  command=win.destroy).pack(side="right")

    # ── SUB-ABA: CNAB 240 ────────────────────────────────────────────────────
    def _cap_build_cnab(self, parent, bg, surface, accent, green, yellow, text, muted):

        tk.Label(parent, text="🏦  Gerador CNAB 240 — Itaú",
                 font=("Segoe UI",12,"bold"), fg=accent, bg=bg
                 ).pack(anchor="w", padx=20, pady=(12,2))
        tk.Label(parent,
                 text="Gera arquivo de remessa para envio ao Itaú Empresas. "
                      "Um arquivo por empresa. Suba no Sispag → Cobrança/Pagamento → Remessa.",
                 font=("Segoe UI",8), fg=muted, bg=bg
                 ).pack(anchor="w", padx=20, pady=(0,8))

        # ── CONFIGURAÇÕES ──
        cfg = tk.LabelFrame(parent, text="  Configurações  ", bg=surface,
                             fg=accent, font=("Segoe UI",9,"bold"),
                             padx=14, pady=10)
        cfg.pack(fill="x", padx=20, pady=(0,8))

        def lbl(t): return tk.Label(cfg, text=t, fg=muted, bg=surface,
                                     font=("Segoe UI",8))

        lbl("Empresa:").grid(row=0, column=0, sticky="w", pady=3)
        self._cnab_empresa = ttk.Combobox(cfg,
            values=["LALUA","SOLAR","AS DUAS (arquivos separados)"],
            state="readonly", width=28, font=("Segoe UI",9))
        self._cnab_empresa.set("LALUA")
        self._cnab_empresa.grid(row=0, column=1, sticky="w", padx=(0,20))

        lbl("Fonte dos pagamentos:").grid(row=0, column=2, sticky="w")
        self._cnab_fonte = ttk.Combobox(cfg,
            values=["Notas lançadas (Contas a Pagar)",
                    "Autorização de Pagamentos (carregada)"],
            state="readonly", width=35, font=("Segoe UI",9))
        self._cnab_fonte.set("Notas lançadas (Contas a Pagar)")
        self._cnab_fonte.grid(row=0, column=3, sticky="w")

        lbl("Filtrar vencimento:").grid(row=1, column=0, sticky="w", pady=3)
        self._cnab_dt_de = tk.Entry(cfg, font=("Segoe UI",9),
                                     bg="#0a0f1e", fg=text,
                                     insertbackground=accent,
                                     relief="flat", bd=2, width=11)
        self._cnab_dt_de.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self._cnab_dt_de.grid(row=1, column=1, sticky="w", padx=(0,6))
        _aplicar_mascara_data(self._cnab_dt_de)

        lbl("até:").grid(row=1, column=2, sticky="w")
        self._cnab_dt_ate = tk.Entry(cfg, font=("Segoe UI",9),
                                      bg="#0a0f1e", fg=text,
                                      insertbackground=accent,
                                      relief="flat", bd=2, width=11)
        self._cnab_dt_ate.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self._cnab_dt_ate.grid(row=1, column=3, sticky="w")
        _aplicar_mascara_data(self._cnab_dt_ate)

        lbl("Nº Remessa:").grid(row=2, column=0, sticky="w", pady=3)
        self._cnab_num_rem = tk.Spinbox(cfg, from_=1, to=9999,
                                         width=6, font=("Segoe UI",9),
                                         bg="#0a0f1e", fg=text,
                                         insertbackground=accent,
                                         relief="flat", bd=2)
        self._cnab_num_rem.grid(row=2, column=1, sticky="w")

        lbl("Pasta de saída:").grid(row=2, column=2, sticky="w")
        self._cnab_pasta_var = tk.StringVar(value=PASTA_ATUAL)
        tk.Entry(cfg, textvariable=self._cnab_pasta_var,
                 font=("Consolas",8), bg="#0a0f1e", fg=text,
                 insertbackground=accent, relief="flat", bd=2, width=35
                 ).grid(row=2, column=3, sticky="w", padx=(0,6))
        tk.Button(cfg, text="📂", font=("Segoe UI",8),
                  bg=surface, fg=accent, relief="flat", bd=0, cursor="hand2",
                  command=self._cnab_selecionar_pasta
                  ).grid(row=2, column=4, padx=(0,4))

        # ── PREVIEW DE PAGAMENTOS ──
        tk.Label(parent, text="📋 Pagamentos que serão incluídos no arquivo:",
                 font=("Segoe UI",9,"bold"), fg=text, bg=bg
                 ).pack(anchor="w", padx=20, pady=(4,2))

        cols = ("Tipo","Fornecedor/Beneficiário","CNPJ","Vencimento",
                "Valor","Empresa","Código Barras / Chave")
        self._cnab_tree = ttk.Treeview(parent, columns=cols,
                                        show="headings", height=10)
        largs = [80,200,130,90,100,70,220]
        for c, w in zip(cols, largs):
            self._cnab_tree.heading(c, text=c)
            self._cnab_tree.column(c, width=w,
                anchor="e" if c=="Valor" else "w")

        self._cnab_tree.tag_configure("BOLETO", foreground="#38bdf8")
        self._cnab_tree.tag_configure("PIX",    foreground="#4ade80")
        self._cnab_tree.tag_configure("TED",    foreground="#a78bfa")
        self._cnab_tree.tag_configure("DARF",   foreground="#fbbf24")
        self._cnab_tree.tag_configure("GPS",    foreground="#fb923c")
        self._cnab_tree.tag_configure("DAS",    foreground="#f472b6")
        self._cnab_tree.tag_configure("CONC",   foreground="#94a3b8")

        sb = ttk.Scrollbar(parent, orient="vertical",
                            command=self._cnab_tree.yview)
        self._cnab_tree.configure(yscrollcommand=sb.set)

        # Rodapé
        rod = tk.Frame(parent, bg=surface, pady=6, padx=14)
        rod.pack(fill="x", side="bottom")

        self._cnab_lbl_status = tk.Label(rod, text="",
                                          font=("Segoe UI",9,"bold"),
                                          fg=muted, bg=surface)
        self._cnab_lbl_status.pack(side="left")

        tk.Button(rod, text="📄 GERAR CNAB 240",
                  font=("Segoe UI",10,"bold"), bg="#7c3aed", fg="white",
                  relief="flat", bd=0, padx=18, pady=7, cursor="hand2",
                  command=self._cnab_gerar).pack(side="right", padx=(6,0))

        tk.Button(rod, text="🔍 Carregar preview",
                  font=("Segoe UI",9), bg="#1e293b", fg=accent,
                  relief="flat", bd=0, padx=12, pady=7, cursor="hand2",
                  command=self._cnab_carregar_preview).pack(side="right")

        self._cnab_tree.pack(side="left", fill="both", expand=True,
                              padx=(20,0), pady=(0,0))
        sb.pack(side="left", fill="y")

        # Guarda pagamentos carregados
        self._cnab_pagamentos = []

    def _cnab_selecionar_pasta(self):
        from tkinter import filedialog
        p = filedialog.askdirectory(title="Pasta para o arquivo CNAB")
        if p:
            self._cnab_pasta_var.set(p)

    def _cnab_carregar_preview(self):
        """Carrega os pagamentos conforme filtros e mostra na treeview."""
        for r in self._cnab_tree.get_children():
            self._cnab_tree.delete(r)
        self._cnab_pagamentos = []

        empresa = self._cnab_empresa.get()
        fonte   = self._cnab_fonte.get()
        dt_de   = self._cnab_dt_de.get().strip()
        dt_ate  = self._cnab_dt_ate.get().strip()

        try:
            from datetime import datetime as dt_cls
            d_de  = dt_cls.strptime(dt_de,  "%d/%m/%Y") if dt_de  else None
            d_ate = dt_cls.strptime(dt_ate, "%d/%m/%Y") if dt_ate else None
        except:
            self._cnab_lbl_status.config(
                text="⚠️ Data inválida.", fg="#fbbf24")
            return

        emps = ["LALUA","SOLAR"] if "DUAS" in empresa else [empresa]
        pgtos = []

        if "Notas" in fonte:
            # Carrega das notas aprovadas/pendentes no banco
            for emp in emps:
                notas = listar_notas(status=None, empresa=emp)
                for n in notas:
                    if n["status"] in ("CANCELADA","PAGA"):
                        continue
                    # Filtra por vencimento
                    try:
                        dv = datetime.strptime(n["dt_vencimento"], "%d/%m/%Y")
                        if d_de  and dv < d_de:  continue
                        if d_ate and dv > d_ate: continue
                    except: pass

                    # Determina tipo de pagamento
                    cat = (n["categoria"] or "").upper()
                    desc = (n["descricao"] or "").upper()
                    if "DARF" in desc or "IRRF" in cat or "IRRF" in desc:
                        tp = "DARF"
                    elif "GPS" in desc or "INSS" in cat:
                        tp = "GPS"
                    elif "DAS" in desc or "SIMPLES" in desc:
                        tp = "DAS"
                    elif "ENERGIA" in cat or "ÁGUA" in cat or "CONCESS" in cat:
                        tp = "CONC"
                    elif "PIX" in desc:
                        tp = "PIX"
                    else:
                        tp = "BOLETO"

                    # Busca dados do fornecedor
                    forn = buscar_fornecedor_por_cnpj(n["cnpj"]) or {}
                    pix_chave = dict(forn).get("pix_chave","") if forn else ""
                    banco     = dict(forn).get("banco","") if forn else ""
                    agencia   = dict(forn).get("agencia","") if forn else ""
                    conta     = dict(forn).get("conta","") if forn else ""

                    pgto = {
                        "tipo_pgto":      tp,
                        "nome":           n["fornecedor"],
                        "cpf_cnpj_dest":  n["cpf_cnpj_dest"] or n["cnpj"],
                        "cnpj_pagador":   CNAB_CONTAS[emp]["cnpj"],
                        "nome_benef":     n["fornecedor"],
                        "valor":          n["valor_liquido"],
                        "dt_vencimento":  n["dt_vencimento"],
                        "dt_pagamento":   n["dt_vencimento"],
                        "empresa":        emp,
                        "nota_id":        n["id"],
                        "numero_tx":      n["numero_tx"],
                        "banco_dest":     n["banco_dest"] or banco or "341",
                        "agencia_dest":   n["agencia_dest"] or agencia or "0",
                        "conta_dest":     n["conta_dest"] or conta or "0",
                        "chave_pix":      n["pix_chave"] or pix_chave,
                        "cod_barras":     n["cod_barras"] or "",
                    }
                    pgtos.append(pgto)

        else:
            # Carrega da autorização de pagamentos já carregada na aba
            if hasattr(self, "_aut_pagamentos") and self._aut_pagamentos:
                for p in self._aut_pagamentos:
                    emp = p.get("empresa","LALUA")
                    if emp not in emps:
                        continue
                    try:
                        dv = datetime.strptime(p.get("data",""), "%d/%m/%Y")
                        if d_de  and dv < d_de:  continue
                        if d_ate and dv > d_ate: continue
                    except: pass

                    tipo_raw = p.get("tipo","").upper()
                    if "BOLETO" in tipo_raw:
                        tp = "BOLETO"
                    elif "PIX QR" in tipo_raw:
                        tp = "BOLETO"
                    elif "PIX" in tipo_raw:
                        tp = "PIX"
                    elif "TED" in tipo_raw:
                        tp = "TED"
                    elif "DARF" in tipo_raw:
                        tp = "DARF"
                    elif "DAS" in tipo_raw:
                        tp = "DAS"
                    elif "CONCESS" in tipo_raw or "TRIBUTO" in tipo_raw:
                        tp = "CONC"
                    else:
                        tp = "BOLETO"

                    forn = buscar_fornecedor_por_cnpj(p.get("cnpj","")) or {}
                    pgto = {
                        "tipo_pgto":      tp,
                        "nome":           p.get("nome",""),
                        "cpf_cnpj_dest":  p.get("cnpj",""),
                        "cnpj_pagador":   CNAB_CONTAS[emp]["cnpj"],
                        "nome_benef":     p.get("nome",""),
                        "valor":          p.get("valor",0),
                        "dt_vencimento":  p.get("data",""),
                        "dt_pagamento":   p.get("data",""),
                        "empresa":        emp,
                        "banco_dest":     dict(forn).get("banco","341") if forn else "341",
                        "agencia_dest":   dict(forn).get("agencia","0") if forn else "0",
                        "conta_dest":     dict(forn).get("conta","0") if forn else "0",
                        "chave_pix":      dict(forn).get("pix_chave","") if forn else "",
                        "cod_barras":     "",
                    }
                    pgtos.append(pgto)

        # Preenche treeview
        total = 0.0
        for p in pgtos:
            val = _parse_valor(p.get("valor",0))
            total += val
            self._cnab_tree.insert("", "end", tags=(p["tipo_pgto"],),
                values=(
                    p["tipo_pgto"],
                    p["nome"][:30],
                    p.get("cpf_cnpj_dest","")[:18],
                    p.get("dt_vencimento",""),
                    f"R$ {val:,.2f}",
                    p.get("empresa",""),
                    p.get("cod_barras","") or p.get("chave_pix",""),
                ))

        self._cnab_pagamentos = pgtos
        self._cnab_lbl_status.config(
            text=f"{len(pgtos)} pagamento(s)  |  "
                 f"Total: R$ {total:,.2f}",
            fg="#4ade80" if pgtos else "#fbbf24")

    def _cnab_gerar(self):
        if not self._cnab_pagamentos:
            self._cnab_lbl_status.config(
                text="⚠️ Carregue o preview primeiro.", fg="#fbbf24")
            return

        empresa  = self._cnab_empresa.get()
        pasta    = self._cnab_pasta_var.get().strip()
        num_rem  = int(self._cnab_num_rem.get())

        try:
            num_rem = int(self._cnab_num_rem.get())
        except:
            num_rem = 1

        os.makedirs(pasta, exist_ok=True)
        emps = ["LALUA","SOLAR"] if "DUAS" in empresa else [empresa]
        arquivos_gerados = []

        for emp in emps:
            pgtos_emp = [p for p in self._cnab_pagamentos
                         if p.get("empresa","LALUA") == emp]
            if not pgtos_emp:
                continue

            dt_str   = datetime.now().strftime("%Y%m%d_%H%M")
            # Nome máximo 8 chars + .rem (padrão Sispag Itaú)
            # Ex: LAL00001.rem / SOL00001.rem
            prefixo = "LAL" if emp == "LALUA" else "SOL"
            nome_arq = f"{prefixo}{str(num_rem).zfill(5)}.rem"
            caminho  = os.path.join(pasta, nome_arq)

            ok, resultado = gerar_cnab240(pgtos_emp, emp, caminho, num_rem)
            if ok:
                arquivos_gerados.append(nome_arq)
            else:
                self._cnab_lbl_status.config(
                    text=f"❌ Erro {emp}: {resultado}", fg="#f87171")
                return

        if arquivos_gerados:
            msg = f"✅ Gerado: {', '.join(arquivos_gerados)}"
            self._cnab_lbl_status.config(text=msg, fg="#4ade80")
            # Abre a pasta
            try:
                os.startfile(pasta)
            except: pass

    def _build_aba_robo(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        # ── STATS BAR ──
        stats_frame = tk.Frame(parent, bg=bg, pady=10)
        stats_frame.pack(fill="x", padx=20)
        self.lbl_processados = self._stat_card(stats_frame, "PROCESSADOS", "0", accent)
        self.lbl_duplicados   = self._stat_card(stats_frame, "DUPLICADOS",  "0", yellow)
        entrada_count = self._contar_pdfs()
        self.lbl_entrada = self._stat_card(stats_frame, "NA ENTRADA", str(entrada_count), green)

        # ── PASTA INFO ──
        path_frame = tk.Frame(parent, bg=surface, pady=8, padx=14)
        path_frame.pack(fill="x", padx=20, pady=(0, 8))
        tk.Label(path_frame, text="📁 ENTRADA:", font=("Consolas", 9), fg=muted, bg=surface).pack(side="left")
        tk.Label(path_frame, text=PASTA_ENTRADA, font=("Consolas", 9), fg=text, bg=surface).pack(side="left", padx=6)

        # ── LOG ──
        tk.Label(parent, text="📋 Log em tempo real", font=("Segoe UI", 10, "bold"), fg=text, bg=bg).pack(anchor="w", padx=20)
        self.log_area = scrolledtext.ScrolledText(
            parent, height=12, font=("Consolas", 9),
            bg="#0a0f1e", fg="#a0c4ff", insertbackground=accent,
            relief="flat", bd=0, padx=10, pady=10
        )
        self.log_area.pack(fill="both", padx=20, pady=(4, 0), expand=True)
        self.log_area.config(state="disabled")

        # ── BARRA DE PROGRESSO ──
        self.progress = ttk.Progressbar(parent, mode="indeterminate")
        self.progress.pack(fill="x", padx=20, pady=(8, 0))

        # ── BOTÕES ──
        btn_frame = tk.Frame(parent, bg=bg, pady=10)
        btn_frame.pack(fill="x", padx=20)
        self.btn_rodar = tk.Button(btn_frame, text="▶   RODAR ROBÔ AGORA",
            font=("Segoe UI", 13, "bold"), bg="#0ea5e9", fg="white",
            activebackground="#0284c7", activeforeground="white",
            relief="flat", bd=0, padx=30, pady=12, cursor="hand2",
            command=self._iniciar_robo)
        self.btn_rodar.pack(side="left", expand=True, fill="x", padx=(0, 8))
        tk.Button(btn_frame, text="📂 Abrir Saída",
            font=("Segoe UI", 10), bg=surface, fg=text,
            activebackground=border, relief="flat", bd=0,
            padx=14, pady=12, cursor="hand2",
            command=self._abrir_saida).pack(side="left")

        # ── BOTÃO REENVIAR ──
        btn_frame2 = tk.Frame(parent, bg=bg)
        btn_frame2.pack(fill="x", padx=20, pady=(0, 10))
        tk.Button(btn_frame2, text="📧 Reenviar Relatório RH por Período",
            font=("Segoe UI", 10), bg="#1e3a2f", fg="#4ade80",
            activebackground="#166534", activeforeground="white",
            relief="flat", bd=0, padx=16, pady=8, cursor="hand2",
            command=self._abrir_dialogo_periodo).pack(side="left", expand=True, fill="x")

        self._log("🤖 Robô pronto! Coloque PDFs na pasta ENTRADA e clique em RODAR.")
        self._log(f"📂 Pasta: {PASTA_ENTRADA}\n")

    def _build_aba_busca(self, parent, bg, surface, border, accent, green, text, muted):

        # Categorias reais — mesmas do plano de contas + pasta legacy
        CATEGORIAS_BUSCA = ["Todas"] + CATEGORIAS_LISTA + [
            "RH - SALARIOS", "RH - DOMINGOS E FERIADOS", "RH - FGTS",
            "RH - RESCISAO", "RH - FERIAS", "RH - ADIANTAMENTO",
            "FORNECEDORES", "IMPOSTOS", "CONTAS_CONSUMO", "A_VERIFICAR"
        ]

        # ── FILTROS ──
        filtros = tk.Frame(parent, bg=surface, pady=10, padx=14)
        filtros.pack(fill="x", padx=20, pady=(12, 6))

        # Linha 1: Nome + Categoria com busca rápida
        linha1 = tk.Frame(filtros, bg=surface)
        linha1.pack(fill="x", pady=(0, 6))

        tk.Label(linha1, text="Nome:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).pack(side="left")
        self.busca_nome = tk.Entry(linha1, font=("Segoe UI", 10),
                                   bg="#0f172a", fg=text,
                                   insertbackground=accent,
                                   relief="flat", bd=4, width=25)
        self.busca_nome.pack(side="left", padx=(4, 16))
        self.busca_nome.bind("<Return>", lambda e: self._executar_busca())

        tk.Label(linha1, text="Categoria:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).pack(side="left")

        frm_cat_busca = tk.Frame(linha1, bg=surface)
        frm_cat_busca.pack(side="left", padx=(4,0))

        self._busca_cat_filtro = tk.StringVar()
        tk.Entry(frm_cat_busca, textvariable=self._busca_cat_filtro,
                 font=("Segoe UI",8), bg="#0a0f1e", fg=text,
                 insertbackground=accent, relief="flat", bd=2, width=20
                 ).pack(fill="x")
        tk.Label(frm_cat_busca, text="🔍 filtre a categoria",
                 fg="#475569", bg=surface, font=("Consolas",6)).pack(anchor="w")

        self.busca_cat = ttk.Combobox(frm_cat_busca,
                                       values=CATEGORIAS_BUSCA,
                                       state="readonly", width=25,
                                       font=("Segoe UI", 8))
        self.busca_cat.set("Todas")
        self.busca_cat.pack(fill="x")

        def _filtrar_cat_busca(*args):
            termo = self._busca_cat_filtro.get().upper()
            if not termo:
                self.busca_cat["values"] = CATEGORIAS_BUSCA
            else:
                self.busca_cat["values"] = [c for c in CATEGORIAS_BUSCA
                                             if termo in c.upper()]
                if self.busca_cat["values"]:
                    self.busca_cat.set(self.busca_cat["values"][0])
        self._busca_cat_filtro.trace_add("write", _filtrar_cat_busca)

        # Linha 2: Datas + Valores + Empresa
        linha2 = tk.Frame(filtros, bg=surface)
        linha2.pack(fill="x")

        tk.Label(linha2, text="De:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).pack(side="left")
        self.busca_data_de = tk.Entry(linha2, font=("Segoe UI", 10),
                                       bg="#0f172a", fg=text,
                                       insertbackground=accent,
                                       relief="flat", bd=4, width=12)
        self.busca_data_de.insert(0, "01/01/2023")
        self.busca_data_de.pack(side="left", padx=(4, 12))
        _aplicar_mascara_data(self.busca_data_de)

        tk.Label(linha2, text="Até:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).pack(side="left")
        self.busca_data_ate = tk.Entry(linha2, font=("Segoe UI", 10),
                                        bg="#0f172a", fg=text,
                                        insertbackground=accent,
                                        relief="flat", bd=4, width=12)
        self.busca_data_ate.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.busca_data_ate.pack(side="left", padx=(4, 16))
        _aplicar_mascara_data(self.busca_data_ate)

        tk.Label(linha2, text="Valor mín:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).pack(side="left")
        self.busca_val_min = tk.Entry(linha2, font=("Segoe UI", 10),
                                       bg="#0f172a", fg=text,
                                       insertbackground=accent,
                                       relief="flat", bd=4, width=10)
        self.busca_val_min.pack(side="left", padx=(4, 12))

        tk.Label(linha2, text="Valor máx:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).pack(side="left")
        self.busca_val_max = tk.Entry(linha2, font=("Segoe UI", 10),
                                       bg="#0f172a", fg=text,
                                       insertbackground=accent,
                                       relief="flat", bd=4, width=10)
        self.busca_val_max.pack(side="left", padx=(4, 12))

        tk.Label(linha2, text="Empresa:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).pack(side="left")
        self.busca_empresa = ttk.Combobox(linha2,
                                           values=["Todas","LALUA","SOLAR"],
                                           state="readonly", width=8,
                                           font=("Segoe UI", 9))
        self.busca_empresa.set("Todas")
        self.busca_empresa.pack(side="left", padx=(4, 0))

        # ── BOTÃO BUSCAR ──
        frm_buscar = tk.Frame(parent, bg=bg)
        frm_buscar.pack(fill="x", padx=20, pady=(0,6))

        tk.Button(frm_buscar, text="🔍  BUSCAR",
                  font=("Segoe UI", 11, "bold"), bg=accent, fg="#0f172a",
                  activebackground="#0ea5e9", relief="flat", bd=0,
                  padx=20, pady=8, cursor="hand2",
                  command=self._executar_busca).pack(side="left")

        tk.Button(frm_buscar, text="🔄 Reclassificar todos",
                  font=("Segoe UI", 8), bg="#1e293b", fg=muted,
                  relief="flat", bd=0, padx=10, pady=8, cursor="hand2",
                  command=self._reclassificar_comprovantes).pack(side="left", padx=(8,0))

        self.lbl_resultado = tk.Label(frm_buscar, text="",
                                       fg=muted, bg=bg,
                                       font=("Segoe UI", 9))
        self.lbl_resultado.pack(side="left", padx=(16,0))

        # ── RESULTADO ──
        cols = ("Data", "Nome", "Categoria", "Filial", "Empresa", "Valor")
        self.tree = ttk.Treeview(parent, columns=cols,
                                  show="headings", height=12)
        style2 = ttk.Style()
        style2.configure("Treeview", background="#0f172a", foreground=text,
                         fieldbackground="#0f172a", rowheight=24,
                         font=("Segoe UI", 9))
        style2.configure("Treeview.Heading", background=surface,
                         foreground=accent, font=("Segoe UI", 9, "bold"))
        style2.map("Treeview", background=[("selected", "#1e3a5f")])

        larguras = [90, 240, 180, 120, 70, 90]
        for col, w in zip(cols, larguras):
            self.tree.heading(col, text=col,
                              command=lambda c=col: self._busca_ordenar(c))
            self.tree.column(col, width=w,
                             anchor="w" if col != "Valor" else "e")

        scroll_y = ttk.Scrollbar(parent, orient="vertical",
                                  command=self.tree.yview)
        scroll_x = ttk.Scrollbar(parent, orient="horizontal",
                                  command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set,
                             xscrollcommand=scroll_x.set)

        # Duplo clique abre o PDF
        self.tree.bind("<Double-1>", self._busca_abrir_pdf)

        # ── RODAPÉ: AÇÕES ──
        rod = tk.Frame(parent, bg=surface, pady=6, padx=14)
        rod.pack(fill="x", side="bottom")

        tk.Label(rod, text="📧 E-mail:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).pack(side="left")
        self.busca_email_destino = tk.Entry(rod, font=("Segoe UI", 9),
                                             bg="#0a0f1e", fg=text,
                                             insertbackground=accent,
                                             relief="flat", bd=3, width=28)
        self.busca_email_destino.insert(0, EMAIL_DESTINO_DP)
        self.busca_email_destino.pack(side="left", padx=(4, 8))

        tk.Button(rod, text="📤 Enviar por E-mail",
                  font=("Segoe UI", 9, "bold"), bg="#1e3a2f", fg="#4ade80",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=self._enviar_selecionados).pack(side="left", padx=(0,6))

        tk.Button(rod, text="💾 Baixar Selecionados",
                  font=("Segoe UI", 9, "bold"), bg="#1e2a3a", fg="#38bdf8",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=self._baixar_selecionados).pack(side="left", padx=(0,6))

        tk.Button(rod, text="📋 Enviar p/ Autorização",
                  font=("Segoe UI", 9, "bold"), bg="#3a1e2a", fg="#f472b6",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=self._enviar_para_autorizacao).pack(side="left")

        self.lbl_status_envio = tk.Label(rod, text="",
                                          fg="#4ade80", bg=surface,
                                          font=("Segoe UI", 9, "bold"))
        self.lbl_status_envio.pack(side="left", padx=(12,0))

        # Layout treeview
        scroll_x.pack(side="bottom", fill="x", padx=20)
        self.tree.pack(side="left", fill="both", expand=True,
                       padx=(20, 0), pady=(4, 0))
        scroll_y.pack(side="left", fill="y", pady=(4, 0))

        # Inicializa dados
        self._resultados_busca   = []
        self._busca_ordem_col    = None
        self._busca_ordem_rev    = False

    def _build_aba_extrato(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        # ── TÍTULO ──
        tk.Label(parent, text="📊 Extrato Bancário + Adquirentes → Fluxo de Caixa",
                 font=("Segoe UI", 13, "bold"), fg=accent, bg=bg).pack(anchor="w", padx=20, pady=(14, 2))
        tk.Label(parent, text="Combine o extrato Itaú com os extratos dos adquirentes para gerar o fluxo consolidado.",
                 font=("Segoe UI", 9), fg=muted, bg=bg).pack(anchor="w", padx=20, pady=(0, 8))

        # ── SEÇÃO ITAÚ ──
        frame_itau = tk.LabelFrame(parent, text="  🏦 Extrato Itaú (PDF)  ",
                                    bg=surface, fg=accent, font=("Segoe UI", 9, "bold"),
                                    padx=12, pady=8)
        frame_itau.pack(fill="x", padx=20, pady=(0, 6))

        tk.Label(frame_itau, text="📄 PDF do Extrato:", fg=muted, bg=surface,
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=(0,6))
        self.extrato_caminho = tk.StringVar()
        self.extrato_entry = tk.Entry(frame_itau, textvariable=self.extrato_caminho,
                                       font=("Consolas", 8), bg="#0f172a", fg=text,
                                       insertbackground=accent, relief="flat", bd=3, width=55)
        self.extrato_entry.grid(row=0, column=1, padx=(0,6), sticky="w")
        tk.Button(frame_itau, text="📂", font=("Segoe UI", 9), bg=surface, fg=accent,
                  relief="flat", bd=0, cursor="hand2",
                  command=self._selecionar_extrato).grid(row=0, column=2)
        tk.Label(frame_itau, text="(opcional se só usar adquirentes)",
                 fg="#475569", bg=surface, font=("Consolas", 7)).grid(row=0, column=3, padx=6)

        # ── SEÇÃO GETNET ──
        frame_getnet = tk.LabelFrame(parent, text="  💳 Getnet (Excel .xlsx)  ",
                                      bg=surface, fg=yellow, font=("Segoe UI", 9, "bold"),
                                      padx=12, pady=8)
        frame_getnet.pack(fill="x", padx=20, pady=(0, 6))

        self._getnet_paths = []
        self._getnet_frames = []

        def _add_getnet_linha(path=""):
            idx = len(self._getnet_paths)
            var = tk.StringVar(value=path)
            self._getnet_paths.append(var)
            row_f = tk.Frame(frame_getnet, bg=surface)
            row_f.pack(fill="x", pady=2)
            tk.Label(row_f, text=f"Filial {idx+1}:", fg=muted, bg=surface,
                     font=("Segoe UI", 8), width=7).pack(side="left")
            tk.Entry(row_f, textvariable=var, font=("Consolas", 8),
                     bg="#0f172a", fg=text, insertbackground=accent,
                     relief="flat", bd=3, width=55).pack(side="left", padx=(0,6))
            tk.Button(row_f, text="📂", font=("Segoe UI", 8), bg=surface, fg=yellow,
                      relief="flat", bd=0, cursor="hand2",
                      command=lambda v=var: _sel_getnet(v)).pack(side="left")
            self._getnet_frames.append(row_f)

        def _sel_getnet(var):
            from tkinter import filedialog
            p = filedialog.askopenfilename(
                title="Selecionar extrato Getnet",
                filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")]
            )
            if p: var.set(p)

        def _add_mais():
            _add_getnet_linha()

        _add_getnet_linha()  # começa com 1 campo

        tk.Button(frame_getnet, text="+ Adicionar filial",
                  font=("Segoe UI", 8), bg="#1e293b", fg=yellow,
                  relief="flat", bd=0, padx=8, pady=3, cursor="hand2",
                  command=_add_mais).pack(anchor="w", pady=(4,0))

        # ── BOTÕES ──
        frame_bts = tk.Frame(parent, bg=bg)
        frame_bts.pack(fill="x", padx=20, pady=(4, 4))

        tk.Button(frame_bts, text="⚡  GERAR FLUXO CONSOLIDADO",
                  font=("Segoe UI", 11, "bold"), bg="#7c3aed", fg="white",
                  activebackground="#6d28d9", relief="flat", bd=0,
                  padx=20, pady=9, cursor="hand2",
                  command=self._processar_fluxo_consolidado).pack(side="left", padx=(0,8))

        tk.Button(frame_bts, text="📊 Só Extrato Itaú",
                  font=("Segoe UI", 9), bg=surface, fg=accent,
                  relief="flat", bd=0, padx=12, pady=9, cursor="hand2",
                  command=self._processar_extrato).pack(side="left")

        # ── ENVIO POR EMAIL ──
        frame_email = tk.Frame(parent, bg=surface, pady=6, padx=14)
        frame_email.pack(fill="x", padx=20, pady=(0, 6))
        tk.Label(frame_email, text="📧 Enviar para:", fg="#94a3b8", bg=surface,
                 font=("Segoe UI", 9)).pack(side="left")
        self.extrato_email = tk.Entry(frame_email, font=("Segoe UI", 10),
                                       bg="#0f172a", fg="#e2e8f0",
                                       insertbackground=accent, relief="flat", bd=4, width=35)
        self.extrato_email.insert(0, EMAIL_DESTINO_EXTRATO)
        self.extrato_email.pack(side="left", padx=(8, 10))
        tk.Button(frame_email, text="📤 Enviar por E-mail",
                  font=("Segoe UI", 9, "bold"), bg="#1e3a2f", fg="#4ade80",
                  activebackground="#166534", relief="flat", bd=0,
                  padx=14, pady=4, cursor="hand2",
                  command=self._enviar_dashboard_email).pack(side="left")

        # ── STATUS ──
        self.extrato_status = tk.Label(parent, text="", fg=muted, bg=bg,
                                        font=("Segoe UI", 10, "bold"))
        self.extrato_status.pack(anchor="w", padx=20)

        # ── PREVIEW ──
        tk.Label(parent, text="📋 Preview",
                 font=("Segoe UI", 9, "bold"), fg=text, bg=bg).pack(anchor="w", padx=20, pady=(6, 2))

        cols = ("Origem", "Data", "Descrição / Favorecido", "Tipo", "Entrada", "Saída")
        self.extrato_tree = ttk.Treeview(parent, columns=cols, show="headings", height=10)
        larguras = [100, 80, 320, 160, 110, 110]
        for col, w in zip(cols, larguras):
            self.extrato_tree.heading(col, text=col)
            self.extrato_tree.column(col, width=w,
                                      anchor="e" if col in ("Entrada","Saída") else "w")

        self.extrato_tree.tag_configure("deb",  background="#2d1515", foreground="#f87171")
        self.extrato_tree.tag_configure("cred", background="#152d1f", foreground="#4ade80")
        self.extrato_tree.tag_configure("getnet",background="#1a1510", foreground="#fbbf24")

        scroll_ext = ttk.Scrollbar(parent, orient="vertical", command=self.extrato_tree.yview)
        self.extrato_tree.configure(yscrollcommand=scroll_ext.set)
        self.extrato_tree.pack(side="left", fill="both", expand=True, padx=(20,0), pady=(0,10))
        scroll_ext.pack(side="left", fill="y", pady=(0,10))

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

    # ==========================================================================
    # ── ABA 4: HISTÓRICO DE STATUS ─────────────────────────────────────────
    # ==========================================================================

    def _build_aba_historico_status(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        """Aba que lista todos os logs de execução anteriores do robô."""

        # ── TÍTULO ──
        tk.Label(parent, text="📜  Histórico de Execuções do Robô",
                 font=("Segoe UI", 13, "bold"), fg=accent, bg=bg).pack(anchor="w", padx=20, pady=(14, 4))
        tk.Label(parent, text="Todos os logs salvos automaticamente pelo robô estão listados abaixo.",
                 font=("Segoe UI", 9), fg=muted, bg=bg).pack(anchor="w", padx=20, pady=(0, 8))

        # ── FRAME PRINCIPAL: lista esquerda + detalhe direito ──
        main = tk.Frame(parent, bg=bg)
        main.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # ── LISTA DE LOGS (esquerda) ──
        frame_lista = tk.Frame(main, bg=surface, padx=8, pady=8)
        frame_lista.pack(side="left", fill="y", padx=(0, 10))

        tk.Label(frame_lista, text="Execuções salvas", font=("Segoe UI", 9, "bold"),
                 fg=text, bg=surface).pack(anchor="w", pady=(0, 4))

        self._hist_log_listbox = tk.Listbox(
            frame_lista, font=("Consolas", 9),
            bg="#0a0f1e", fg="#a0c4ff", selectbackground=accent,
            selectforeground="#0f172a", relief="flat", bd=0,
            width=30, activestyle="none"
        )
        scroll_lb = ttk.Scrollbar(frame_lista, orient="vertical",
                                   command=self._hist_log_listbox.yview)
        self._hist_log_listbox.configure(yscrollcommand=scroll_lb.set)
        self._hist_log_listbox.pack(side="left", fill="both", expand=True)
        scroll_lb.pack(side="left", fill="y")
        self._hist_log_listbox.bind("<<ListboxSelect>>", self._carregar_log_selecionado)

        # ── DETALHE DO LOG (direita) ──
        frame_detalhe = tk.Frame(main, bg=bg)
        frame_detalhe.pack(side="left", fill="both", expand=True)

        barra_det = tk.Frame(frame_detalhe, bg=surface, pady=6, padx=10)
        barra_det.pack(fill="x", pady=(0, 6))

        self._hist_lbl_resumo = tk.Label(barra_det, text="Selecione um log para visualizar",
                                          font=("Segoe UI", 9, "bold"), fg=muted, bg=surface)
        self._hist_lbl_resumo.pack(side="left")

        tk.Button(barra_det, text="🗑 Apagar log selecionado",
                  font=("Segoe UI", 8), bg="#7f1d1d", fg="#fca5a5",
                  relief="flat", bd=0, padx=10, pady=4, cursor="hand2",
                  command=self._apagar_log_selecionado).pack(side="right")

        self._hist_log_area = scrolledtext.ScrolledText(
            frame_detalhe, font=("Consolas", 9),
            bg="#0a0f1e", fg="#a0c4ff",
            relief="flat", bd=0, padx=10, pady=10,
            state="disabled"
        )
        self._hist_log_area.pack(fill="both", expand=True)

        # ── BARRA INFERIOR ──
        barra_inf = tk.Frame(parent, bg=bg, pady=6)
        barra_inf.pack(fill="x", padx=20)

        tk.Button(barra_inf, text="🔄 Atualizar lista",
                  font=("Segoe UI", 9, "bold"), bg=surface, fg=text,
                  relief="flat", bd=0, padx=12, pady=6, cursor="hand2",
                  command=self._carregar_lista_logs).pack(side="left", padx=(0, 8))

        tk.Button(barra_inf, text="🗑 Apagar todos os logs",
                  font=("Segoe UI", 9, "bold"), bg="#7f1d1d", fg="#fca5a5",
                  relief="flat", bd=0, padx=12, pady=6, cursor="hand2",
                  command=self._apagar_todos_logs).pack(side="left")

        self._hist_lbl_total = tk.Label(barra_inf, text="",
                                         font=("Segoe UI", 9), fg=muted, bg=bg)
        self._hist_lbl_total.pack(side="right")

        # Carrega lista ao iniciar
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

    # ==========================================================================
    # ── ABA 5: HISTÓRICO DE FLUXO ──────────────────────────────────────────
    # ==========================================================================

    def _build_aba_historico_fluxo(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        """Aba que lê todos os dashboards Excel gerados e monta histórico de saldo por conta."""

        tk.Label(parent, text="📈  Histórico de Fluxo de Caixa",
                 font=("Segoe UI", 13, "bold"), fg=accent, bg=bg).pack(anchor="w", padx=20, pady=(14, 2))
        tk.Label(parent,
                 text="Acumula os saldos de todos os extratos processados. Lê os Dashboards Excel em EXTRATOS_PROCESSADOS.",
                 font=("Segoe UI", 9), fg=muted, bg=bg).pack(anchor="w", padx=20, pady=(0, 8))

        # ── CONTROLES ──
        ctrl = tk.Frame(parent, bg=surface, pady=8, padx=14)
        ctrl.pack(fill="x", padx=20, pady=(0, 8))

        tk.Label(ctrl, text="Conta:", fg=muted, bg=surface, font=("Segoe UI", 9)).pack(side="left")
        self._fluxo_conta_var = tk.StringVar(value="Todas")
        self._fluxo_conta_cb = ttk.Combobox(ctrl, textvariable=self._fluxo_conta_var,
                                              values=["Todas"], state="readonly", width=22)
        self._fluxo_conta_cb.pack(side="left", padx=(4, 16))
        self._fluxo_conta_cb.bind("<<ComboboxSelected>>", lambda e: self._filtrar_fluxo())

        tk.Button(ctrl, text="🔄 Recarregar Extratos",
                  font=("Segoe UI", 9, "bold"), bg=accent, fg="#0f172a",
                  relief="flat", bd=0, padx=12, pady=4, cursor="hand2",
                  command=self._carregar_historico_fluxo).pack(side="left", padx=(0, 8))

        tk.Button(ctrl, text="📥 Exportar para Excel",
                  font=("Segoe UI", 9, "bold"), bg="#1e3a2f", fg=green,
                  relief="flat", bd=0, padx=12, pady=4, cursor="hand2",
                  command=self._exportar_fluxo_excel).pack(side="left")

        self._fluxo_lbl_status = tk.Label(ctrl, text="", font=("Segoe UI", 9),
                                           fg=muted, bg=surface)
        self._fluxo_lbl_status.pack(side="right")

        # ── CARDS DE RESUMO ──
        self._fluxo_cards_frame = tk.Frame(parent, bg=bg)
        self._fluxo_cards_frame.pack(fill="x", padx=20, pady=(0, 8))

        # ── TREEVIEW DE HISTÓRICO ──
        colunas = ("Período", "Conta", "Empresa", "Saldo Anterior", "Créditos", "Débitos", "Saldo Final", "Variação")
        self._fluxo_tree = ttk.Treeview(parent, columns=colunas, show="headings", height=14)
        larguras = [95, 160, 80, 115, 115, 115, 115, 100]
        ancors   = ["center", "w", "center", "e", "e", "e", "e", "e"]
        for col, w, anc in zip(colunas, larguras, ancors):
            self._fluxo_tree.heading(col, text=col,
                                      command=lambda c=col: self._ordenar_fluxo(c))
            self._fluxo_tree.column(col, width=w, anchor=anc)

        self._fluxo_tree.tag_configure("positivo", foreground="#4ade80")
        self._fluxo_tree.tag_configure("negativo", foreground="#f87171")
        self._fluxo_tree.tag_configure("neutro",   foreground="#94a3b8")
        self._fluxo_tree.tag_configure("matriz",   background="#1c1033")

        scroll_f = ttk.Scrollbar(parent, orient="vertical", command=self._fluxo_tree.yview)
        self._fluxo_tree.configure(yscrollcommand=scroll_f.set)

        frame_tree = tk.Frame(parent, bg=bg)
        frame_tree.pack(fill="both", expand=True, padx=20, pady=(0, 4))
        self._fluxo_tree.pack(in_=frame_tree, side="left", fill="both", expand=True)
        scroll_f.pack(in_=frame_tree, side="left", fill="y")

        # ── TOTAIS ──
        self._fluxo_lbl_totais = tk.Label(parent, text="",
                                           font=("Consolas", 9), fg=muted, bg=bg)
        self._fluxo_lbl_totais.pack(anchor="e", padx=24, pady=(0, 8))

        # Dados internos
        self._fluxo_dados = []          # lista de dicts com todas as linhas
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

    def _atualizar_cards_fluxo(self):
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

        bg = "#0f172a"
        for conta, d in sorted(ultimos.items()):
            card = tk.Frame(self._fluxo_cards_frame, bg="#1e293b", padx=14, pady=8, relief="flat")
            card.pack(side="left", expand=True, fill="x", padx=(0, 6), pady=4)
            nome_curto = conta.split("-")[-1].strip()[:14] if "-" in conta else conta[:14]
            tk.Label(card, text=nome_curto, font=("Consolas", 8),
                     fg="#64748b", bg="#1e293b").pack(anchor="w")
            cor_saldo = "#4ade80" if d["saldo_fin"] >= 0 else "#f87171"
            if d["saldo_fin"] < 100_000 and "Matriz" in conta:
                cor_saldo = "#fbbf24"  # alerta crítico Matriz
            tk.Label(card, text=f"R$ {d['saldo_fin']:,.0f}",
                     font=("Segoe UI", 13, "bold"), fg=cor_saldo, bg="#1e293b").pack(anchor="w")
            var = d["variacao"]
            sinal = "▲" if var >= 0 else "▼"
            cor_var = "#4ade80" if var >= 0 else "#f87171"
            tk.Label(card, text=f"{sinal} {abs(var):,.0f}  ({d['periodo']})",
                     font=("Consolas", 8), fg=cor_var, bg="#1e293b").pack(anchor="w")

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

    # ==========================================================================
    # ── ABA 6: RECEBÍVEIS / INTEGRAÇÕES ───────────────────────────────────
    # ==========================================================================

    def _build_aba_recebiveis(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        """Aba com integrações: PagarMe, Rede, Getnet e Linx Microvix."""

        # ── TÍTULO ──
        tk.Label(parent, text="💳  Recebíveis & Integrações",
                 font=("Segoe UI", 13, "bold"), fg=accent, bg=bg).pack(anchor="w", padx=20, pady=(12, 2))
        tk.Label(parent,
                 text="Consulte agenda de recebíveis e vendas em tempo real. Configure as credenciais nas CONFIGURAÇÕES no topo do arquivo.",
                 font=("Segoe UI", 9), fg=muted, bg=bg).pack(anchor="w", padx=20, pady=(0, 8))

        # ── NOTEBOOK INTERNO (sub-abas por operadora) ──
        sub_style = ttk.Style()
        sub_style.configure("Sub.TNotebook", background=bg, borderwidth=0)
        sub_style.configure("Sub.TNotebook.Tab", background="#1e293b", foreground=muted,
                            padding=[12, 6], font=("Segoe UI", 9, "bold"))
        sub_style.map("Sub.TNotebook.Tab",
                      background=[("selected", "#7c3aed")],
                      foreground=[("selected", "white")])

        sub_nb = ttk.Notebook(parent, style="Sub.TNotebook")
        sub_nb.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # Sub-aba PagarMe
        aba_pm = tk.Frame(sub_nb, bg=bg)
        sub_nb.add(aba_pm, text="  💰 PagarMe  ")
        self._build_sub_pagarme(aba_pm, bg, surface, accent, green, yellow, text, muted)

        # Sub-aba Rede
        aba_rede = tk.Frame(sub_nb, bg=bg)
        sub_nb.add(aba_rede, text="  🔴 Rede  ")
        self._build_sub_rede(aba_rede, bg, surface, accent, green, yellow, text, muted)

        # Sub-aba Getnet
        aba_getnet = tk.Frame(sub_nb, bg=bg)
        sub_nb.add(aba_getnet, text="  🟢 Getnet  ")
        self._build_sub_getnet(aba_getnet, bg, surface, accent, green, yellow, text, muted)

        # Sub-aba Microvix
        aba_mv = tk.Frame(sub_nb, bg=bg)
        sub_nb.add(aba_mv, text="  🏪 Microvix  ")
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

        # ── INSERÇÃO MANUAL DE PAGAMENTOS ──
        frame_manual = tk.LabelFrame(parent,
                                      text="  ➕ Inserção Manual de Pagamento  ",
                                      bg=surface, fg=yellow,
                                      font=("Segoe UI", 8, "bold"),
                                      padx=12, pady=6)
        frame_manual.pack(fill="x", padx=20, pady=(0, 6))

        def _lbl(t):
            return tk.Label(frame_manual, text=t, fg=muted, bg=surface,
                            font=("Segoe UI", 8))
        def _ent(w=16):
            return tk.Entry(frame_manual, font=("Segoe UI", 8),
                            bg="#0a0f1e", fg=text,
                            insertbackground=accent,
                            relief="flat", bd=2, width=w)

        # Linha 1 — Nome / Empresa / Tipo
        _lbl("Nome/Beneficiário:").grid(row=0, column=0, sticky="w", pady=2)
        _man_nome = _ent(28)
        _man_nome.grid(row=0, column=1, sticky="w", padx=(0,8))

        _lbl("Empresa:").grid(row=0, column=2, sticky="w")
        _man_emp_var = tk.StringVar(value="LALUA")
        ttk.Combobox(frame_manual, textvariable=_man_emp_var,
                     values=["LALUA","SOLAR"], state="readonly",
                     width=7, font=("Segoe UI",8)
                     ).grid(row=0, column=3, sticky="w", padx=(0,8))

        _lbl("Tipo:").grid(row=0, column=4, sticky="w")
        _man_tipo_var = tk.StringVar(value="PIX")
        ttk.Combobox(frame_manual, textvariable=_man_tipo_var,
                     values=["PIX","TED","BOLETO","DARF","GPS","DAS",
                             "CONCESSIONÁRIA","PIX FILIAL→MATRIZ"],
                     state="readonly", width=16, font=("Segoe UI",8)
                     ).grid(row=0, column=5, sticky="w")

        # Linha 2 — Data / Valor / Categoria / Responsável
        _lbl("Data:").grid(row=1, column=0, sticky="w", pady=2)
        _man_data = _ent(11)
        _man_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        _man_data.grid(row=1, column=1, sticky="w", padx=(0,8))
        _aplicar_mascara_data(_man_data)

        _lbl("Valor:").grid(row=1, column=2, sticky="w")
        _man_valor = _ent(12)
        _man_valor.grid(row=1, column=3, sticky="w", padx=(0,8))
        _aplicar_mascara_valor(_man_valor)

        _lbl("Categoria:").grid(row=1, column=4, sticky="w")
        _man_cat_var = tk.StringVar()
        ttk.Combobox(frame_manual, textvariable=_man_cat_var,
                     values=CATEGORIAS_LISTA, state="normal",
                     width=32, font=("Segoe UI",8)
                     ).grid(row=1, column=5, sticky="w")

        # Linha 3 — Responsável / Observação + Botão
        _lbl("Responsável:").grid(row=2, column=0, sticky="w", pady=2)
        _man_resp_var = tk.StringVar()
        ttk.Combobox(frame_manual, textvariable=_man_resp_var,
                     values=RESPONSAVEIS_LISTA, state="readonly",
                     width=16, font=("Segoe UI",8)
                     ).grid(row=2, column=1, sticky="w", padx=(0,8))

        _lbl("Observação:").grid(row=2, column=2, sticky="w")
        _man_obs = _ent(32)
        _man_obs.grid(row=2, column=3, columnspan=2, sticky="w", padx=(0,8))

        self._aut_lbl_manual = tk.Label(frame_manual, text="",
                                         font=("Segoe UI",8), fg=yellow,
                                         bg=surface)
        self._aut_lbl_manual.grid(row=3, column=0, columnspan=5,
                                   sticky="w", pady=(2,0))

        def _adicionar_manual():
            nome  = _man_nome.get().strip()
            valor_s = _man_valor.get().strip()
            data  = _man_data.get().strip()
            tipo  = _man_tipo_var.get()
            emp   = _man_emp_var.get()
            cat   = _man_cat_var.get().strip()
            resp  = _man_resp_var.get().strip()
            obs   = _man_obs.get().strip()

            if not nome:
                self._aut_lbl_manual.config(text="⚠️ Informe o nome.", fg="#fbbf24")
                return
            try:
                valor = _parse_valor(valor_s)
                if valor <= 0: raise ValueError
            except:
                self._aut_lbl_manual.config(text="⚠️ Valor inválido.", fg="#fbbf24")
                return
            try:
                data_obj = datetime.strptime(data, "%d/%m/%Y")
            except:
                self._aut_lbl_manual.config(text="⚠️ Data inválida.", fg="#fbbf24")
                return

            # PIX Filial→Matriz força categoria
            if tipo == "PIX FILIAL→MATRIZ":
                cat  = "TRANSFERENCIA MATRIZ - PIX Filial"
                resp = resp or "ADM/FINANCEIRO"

            self._aut_pagamentos.append({
                "nome":        nome,
                "cnpj":        "",
                "tipo":        tipo,
                "data":        data,
                "data_obj":    data_obj,
                "valor":       valor,
                "status":      "Aprovada",
                "responsavel": resp or "ADM/FINANCEIRO",
                "categoria":   cat or "A CLASSIFICAR",
                "empresa":     emp,
                "observacao":  obs,
                "manual":      True,
            })
            self._aut_popular_treeview()
            self._aut_lbl_manual.config(
                text=f"✅ {nome} — R$ {valor:,.2f} adicionado!",
                fg="#4ade80")

            # Limpa campos
            _man_nome.delete(0,"end")
            _man_valor.delete(0,"end")
            _man_obs.delete(0,"end")
            _man_data.delete(0,"end")
            _man_data.insert(0, datetime.now().strftime("%d/%m/%Y"))

        tk.Button(frame_manual, text="➕ Adicionar",
                  font=("Segoe UI", 9, "bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=14, pady=5, cursor="hand2",
                  command=_adicionar_manual
                  ).grid(row=2, column=5, sticky="e", pady=2)

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

        self._aut_lbl_resumo = tk.Label(frame_bot, text="",
                                         font=("Consolas", 8), fg=muted, bg=bg)
        self._aut_lbl_resumo.pack(side="left")

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

        # ── TREEVIEW — expande no espaço restante entre cabeçalho e rodapé ──
        container = tk.Frame(parent, bg=bg)
        container.pack(fill="both", expand=True, padx=20, pady=(0, 2))
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        cols = ("Favorecido", "CNPJ", "Tipo", "Data", "Valor", "Responsável", "Categoria", "Empresa")
        self._aut_tree = ttk.Treeview(container, columns=cols, show="headings", height=8)
        larguras = [200, 130, 128, 80, 90, 120, 230, 65]
        for col, w in zip(cols, larguras):
            self._aut_tree.heading(col, text=col,
                                   command=lambda c=col: self._aut_ordenar(c))
            self._aut_tree.column(col, width=w,
                                   anchor="e" if col == "Valor" else "w",
                                   minwidth=40)

        self._aut_tree.tag_configure("ok",           background="#0a1f12", foreground="#4ade80")
        self._aut_tree.tag_configure("pendente",      background="#1a150a", foreground="#fbbf24")
        self._aut_tree.tag_configure("a_classificar", background="#1f0a0a", foreground="#f87171")

        scroll_y = ttk.Scrollbar(container, orient="vertical",   command=self._aut_tree.yview)
        scroll_x = ttk.Scrollbar(container, orient="horizontal", command=self._aut_tree.xview)
        self._aut_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self._aut_tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        self._aut_tree.bind("<Double-1>",        self._aut_editar_linha)
        self._aut_tree.bind("<<TreeviewSelect>>", self._aut_editar_linha)

        # Sugere período ao iniciar
        self._aut_sugerir_periodo()

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

                # LALUA — arquivo principal
                if path_forn and os.path.exists(path_forn):
                    itens = ler_itau_pagamentos(path_forn)
                    for p in itens:
                        p["empresa"] = "LALUA"
                    pagamentos += itens

                # SOLAR — arquivo separado, força empresa = SOLAR
                if path_solar and os.path.exists(path_solar):
                    itens_solar = ler_itau_pagamentos(path_solar)
                    for p in itens_solar:
                        p["empresa"] = "SOLAR"
                    pagamentos += itens_solar

                # Folha de salários
                if path_folha and os.path.exists(path_folha):
                    pagamentos += ler_itau_folha(path_folha)

                # Filtra por período
                filtrados = []
                for p in pagamentos:
                    if p["data_obj"]:
                        if dt_de <= p["data_obj"] <= dt_ate:
                            filtrados.append(p)
                    else:
                        filtrados.append(p)

                self._aut_pagamentos = filtrados
                self.root.after(0, self._aut_popular_treeview)
            except Exception as e:
                self.root.after(0, lambda: self._aut_lbl_status.config(
                    text=f"❌ {e}", fg="#f87171"))

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
        self._aut_lbl_status.config(
            text=f"✅ {len(self._aut_itens_tree)} pagamento(s) | "
                 f"{ok_count} classificados | {sem_cat} pendentes",
            fg="#4ade80" if sem_cat == 0 else "#fbbf24"
        )
        self._aut_lbl_resumo.config(
            text=f"Total: R$ {total:,.2f}   |   "
                 f"{len(self._aut_itens_tree)} lançamento(s)   |   "
                 f"{sem_cat} sem classificação"
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

            self._aut_lbl_status.config(
                text=f"✅ PDF gerado: {nome_arq}", fg="#4ade80")
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

        bg_w    = "#0f172a"
        surf_w  = "#1e293b"
        acc_w   = "#38bdf8"
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
                # Coleta dados das linhas
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

                # ── BLOCO: RESUMO DAS TRANSFERÊNCIAS ──
                # Cabeçalho do bloco
                cb_bloco = [[
                    P("  RESUMO DAS TRANSFERÊNCIAS RECEBIDAS", bold=True,
                      fs=9, cor=BRANCO, align=TA_LEFT),
                    P("Filiais → Conta Matriz LALUA · Ag. 0334 · Cc. 98775-7",
                      fs=8, cor=CREME, align=TA_RIGHT),
                ]]
                t_cb = Table(cb_bloco, colWidths=[doc.width*0.55, doc.width*0.45])
                t_cb.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,-1), CINZA_HDR),
                    ("TOPPADDING",    (0,0), (-1,-1), 6),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 6),
                    ("LEFTPADDING",   (0,0), (0,-1),  8),
                    ("RIGHTPADDING",  (1,0), (1,-1),  8),
                    ("LINEBELOW",     (0,0), (-1,-1), 2, VINHO_CLR),
                    ("LINEBEFORE",    (0,0), (0,-1),  3, VINHO_CLR),
                ]))
                story.append(t_cb)

                # Tabela de transferências
                cw_transf = [doc.width*0.30, doc.width*0.12,
                              doc.width*0.15, doc.width*0.22, doc.width*0.21]
                hdrs_t = ["Filial Remetente", "Data", "Valor (R$)",
                          "Referência / Observação", "Status Conferência"]
                rows_t = [[P(h, fs=7, bold=True, cor=BRANCO, align=TA_CENTER)
                           for h in hdrs_t]]

                for tr in sorted(transferencias, key=lambda x: x["filial"]):
                    rows_t.append([
                        P(f"  {tr['filial']}", fs=8, bold=True, cor=AZUL_PIX),
                        P(tr["data"], fs=8, align=TA_CENTER),
                        P(f"R$ {tr['valor']:,.2f}", fs=9, bold=True,
                          cor=VERDE, align=TA_RIGHT),
                        P(tr["obs"] or "—", fs=7.5),
                        P("( )", fs=9, align=TA_CENTER),  # campo p/ conferência manual
                    ])

                # Linha total
                rows_t.append([
                    P("TOTAL TRANSFERIDO", bold=True, fs=9, cor=BRANCO),
                    "", "",
                    P(f"R$ {total:,.2f}", bold=True, fs=10,
                      cor=CREME, align=TA_RIGHT),
                    "",
                ])

                nr = len(rows_t)
                tbl_t = Table(rows_t, colWidths=cw_transf, repeatRows=1)
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
                    ("LINEBEFORE",    (0,0), (0,-1), 3, VINHO_CLR),
                    ("LINEBELOW",     (0,0), (-1,0), 0.5, BEGE),
                    ("LINEBELOW",     (0,1), (-1,nr-2), 0.3, BEGE),
                    ("TOPPADDING",    (0,0), (-1,-1), 4),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 4),
                    ("LEFTPADDING",   (0,0), (-1,-1), 5),
                    ("RIGHTPADDING",  (0,0), (-1,-1), 5),
                    ("FONTNAME",      (0,nr-1), (-1,nr-1), "Helvetica-Bold"),
                    ("TOPPADDING",    (0,nr-1), (-1,nr-1), 7),
                    ("BOTTOMPADDING", (0,nr-1), (-1,nr-1), 7),
                ]))
                story.append(tbl_t)
                story.append(Spacer(1, 14))

                # ── BLOCO: CONFERÊNCIA DE FORNECEDORES (20 linhas) ──
                story.append(HRFlowable(width="100%", thickness=1,
                                         color=BEGE, spaceBefore=2, spaceAfter=4))

                cb2 = [[
                    P("  CONFERÊNCIA DE FORNECEDORES / LANÇAMENTOS", bold=True,
                      fs=9, cor=BRANCO, align=TA_LEFT),
                    P("Para uso da Matriz — confira os comprovantes recebidos",
                      fs=8, cor=BEGE, align=TA_RIGHT),
                ]]
                t_cb2 = Table(cb2, colWidths=[doc.width*0.6, doc.width*0.4])
                t_cb2.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,-1), VINHO),
                    ("TOPPADDING",    (0,0), (-1,-1), 6),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 6),
                    ("LEFTPADDING",   (0,0), (0,-1),  8),
                    ("RIGHTPADDING",  (1,0), (1,-1),  8),
                    ("LINEBELOW",     (0,0), (-1,-1), 2, VINHO_CLR),
                    ("LINEBEFORE",    (0,0), (0,-1),  3, CREME),
                ]))
                story.append(t_cb2)

                # 20 linhas em branco para conferência
                cw_conf = [doc.width*0.04, doc.width*0.25, doc.width*0.13,
                            doc.width*0.14, doc.width*0.14, doc.width*0.30]
                hdrs_c  = ["#", "Fornecedor / Favorecido", "Data",
                            "Valor (R$)", "Categoria", "Observação / Conferência"]
                rows_c  = [[P(h, fs=7, bold=True, cor=BRANCO, align=TA_CENTER)
                             for h in hdrs_c]]

                for i in range(20):
                    bg_row = BRANCO if i % 2 == 0 else CINZA_CLR
                    rows_c.append([
                        P(str(i+1), fs=7.5, cor=colors.HexColor("#999999"),
                          align=TA_CENTER),
                        P("", fs=8),
                        P("", fs=8),
                        P("", fs=8),
                        P("", fs=8),
                        P("", fs=8),
                    ])

                nr_c = len(rows_c)
                tbl_c = Table(rows_c, colWidths=cw_conf,
                               rowHeights=[None] + [18]*20)
                tbl_c.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,0), CINZA_HDR),
                    ("TEXTCOLOR",     (0,0), (-1,0), BRANCO),
                    ("ROWBACKGROUNDS",(0,1), (-1,-1), [BRANCO, CINZA_CLR]),
                    ("LINEBEFORE",    (0,0), (0,-1), 3, VINHO_CLR),
                    ("LINEBELOW",     (0,0), (-1,0), 0.5, BEGE),
                    ("LINEBELOW",     (0,1), (-1,-1), 0.3, BEGE),
                    ("LINEAFTER",     (0,0), (-1,-1), 0.3, BEGE),
                    ("TOPPADDING",    (0,0), (-1,-1), 2),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 2),
                    ("LEFTPADDING",   (0,0), (-1,-1), 4),
                    ("RIGHTPADDING",  (0,0), (-1,-1), 4),
                    ("ALIGN",         (0,1), (0,-1), "CENTER"),
                    ("FONTSIZE",      (0,1), (-1,-1), 8),
                ]))
                story.append(tbl_c)
                story.append(Spacer(1, 12))

                # ── ASSINATURAS ──
                story.append(HRFlowable(width="100%", thickness=0.5, color=BEGE))
                story.append(Spacer(1, 8))

                cw_ass = [doc.width*0.25] * 4
                ass_data = [[
                    Table([[P("_"*28, fs=8), P("Conferido por / Matriz",
                              fs=7, cor=colors.HexColor("#777777"), align=TA_CENTER)]],
                           colWidths=[doc.width*0.25]),
                    Table([[P("_"*28, fs=8), P("Aprovado por / Financeiro",
                              fs=7, cor=colors.HexColor("#777777"), align=TA_CENTER)]],
                           colWidths=[doc.width*0.25]),
                    Table([[P("_"*28, fs=8), P("Responsável / Filial",
                              fs=7, cor=colors.HexColor("#777777"), align=TA_CENTER)]],
                           colWidths=[doc.width*0.25]),
                    Table([[P("_"*28, fs=8), P(f"Data: ____/____/______",
                              fs=7, cor=colors.HexColor("#777777"), align=TA_CENTER)]],
                           colWidths=[doc.width*0.25]),
                ]]
                tbl_ass = Table(ass_data, colWidths=cw_ass)
                tbl_ass.setStyle(TableStyle([
                    ("ALIGN",  (0,0), (-1,-1), "CENTER"),
                    ("VALIGN", (0,0), (-1,-1), "BOTTOM"),
                ]))
                story.append(tbl_ass)
                story.append(Spacer(1, 6))

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

    def _stat_card(self, parent, label, value, color):
        frame = tk.Frame(parent, bg="#1e293b", padx=20, pady=8, relief="flat")
        frame.pack(side="left", expand=True, fill="x", padx=(0, 8))
        tk.Label(frame, text=label, font=("Consolas", 8), fg="#64748b", bg="#1e293b").pack(anchor="w")
        lbl = tk.Label(frame, text=value, font=("Segoe UI", 22, "bold"), fg=color, bg="#1e293b")
        lbl.pack(anchor="w")
        return lbl

    def _contar_pdfs(self):
        if not os.path.exists(PASTA_ENTRADA): return 0
        count = 0
        for _, _, files in os.walk(PASTA_ENTRADA):
            count += sum(1 for f in files if f.lower().endswith(('.pdf', '.dll')))
        return count

    def _log(self, msg):
        self.log_area.config(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        linha = f"[{timestamp}] {msg}"
        self._log_historico.append(linha)
        self.log_area.insert("end", linha + "\n")
        self.log_area.see("end")
        self.log_area.config(state="disabled")
        self.root.update_idletasks()

    def _atualizar_stats(self, processados, duplicados):
        self.lbl_processados.config(text=str(processados))
        self.lbl_duplicados.config(text=str(duplicados))
        self.root.update_idletasks()

    def _iniciar_robo(self):
        if self.rodando:
            return
        pdfs = self._contar_pdfs()
        if pdfs == 0:
            self._log("⚠️ Nenhum PDF encontrado na pasta ENTRADA!")
            return
        self.rodando = True
        self.btn_rodar.config(text="⏳  PROCESSANDO...", state="disabled", bg="#475569")
        self.lbl_processados.config(text="0")
        self.lbl_duplicados.config(text="0")
        self.lbl_entrada.config(text=str(pdfs))
        self.progress.start(10)
        thread = threading.Thread(target=self._rodar_em_thread, daemon=True)
        thread.start()

    def _rodar_em_thread(self):
        try:
            self._log_historico = []  # Limpa histórico a cada rodada
            log_path = os.path.join(PASTA_ATUAL, f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")

            def log_com_arquivo(msg):
                timestamp = datetime.now().strftime("%H:%M:%S")
                linha = f"[{timestamp}] {msg}"
                self._log_historico.append(linha)
                try:
                    with open(log_path, "a", encoding="utf-8") as f:
                        f.write(linha + "\n")
                except: pass
                self._log(msg)

            self._log(f"🚀 Iniciando processamento de {self._contar_pdfs()} arquivo(s)...")
            self._log("="*50)
            # Cria o arquivo de log com cabeçalho
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(f"Log gerado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n")

            total = organizar_arquivos(log_com_arquivo, self._atualizar_stats)
            self.root.after(0, self._finalizar, total)
        except Exception as e:
            self._log(f"❌ Erro crítico: {e}")
            self.root.after(0, self._finalizar, 0)

    def _finalizar(self, total):
        self.progress.stop()
        self.rodando = False
        if total > 0:
            self.btn_rodar.config(text="✅  CONCLUÍDO — Rodar Novamente", state="normal", bg="#059669")
        else:
            self.btn_rodar.config(text="▶   RODAR ROBÔ AGORA", state="normal", bg="#0ea5e9")
        self.lbl_entrada.config(text=str(self._contar_pdfs()))

    def _abrir_saida(self):
        os.makedirs(PASTA_SAIDA, exist_ok=True)
        os.startfile(PASTA_SAIDA)

    def ler_xml_nfe(self, caminho_xml):
        import xml.etree.ElementTree as ET
        import re
        try:
            tree = ET.parse(caminho_xml)
            root = tree.getroot()
            
            # Helper para buscar ignorando namespaces
            def _find_ns(node, tag):
                for el in node.iter():
                    if el.tag.split('}')[-1] == tag:
                        return el
                return None

            infNFe = _find_ns(root, 'infNFe')
            if infNFe is None: return None

            nNF_node = _find_ns(infNFe, 'nNF')
            nNF = nNF_node.text if nNF_node is not None else "0"
            
            finNFe_node = _find_ns(infNFe, 'finNFe')
            finNFe = finNFe_node.text if finNFe_node is not None else "1"
            
            dhEmi_node = _find_ns(infNFe, 'dhEmi')
            if dhEmi_node is None:
                dhEmi_node = _find_ns(infNFe, 'dEmi')
            data_emissao = dhEmi_node.text[:10] if dhEmi_node is not None and dhEmi_node.text else ""
            
            chave_original = None
            if finNFe == "4":
                refNFe_node = _find_ns(infNFe, 'refNFe')
                if refNFe_node is not None: chave_original = refNFe_node.text

            uf_node = _find_ns(infNFe, 'UF')
            uf = uf_node.text if uf_node is not None else "GERAL"
            
            vNF_node = _find_ns(infNFe, 'vNF')
            vNF = float(vNF_node.text) if vNF_node is not None and vNF_node.text else 0.0
            
            vICMS_node = _find_ns(infNFe, 'vICMS')
            valor_icms = float(vICMS_node.text) if vICMS_node is not None and vICMS_node.text else 0.0
            
            vICMSUFDest_node = _find_ns(infNFe, 'vICMSUFDest')
            valor_difal = float(vICMSUFDest_node.text) if vICMSUFDest_node is not None and vICMSUFDest_node.text else 0.0

            return {
                "nf_numero": nNF, "finalidade": "DEVOLUCAO" if finNFe == "4" else "VENDA",
                "data_emissao": data_emissao, "uf_destino": uf, "valor_total_nota": vNF,
                "valor_imposto_estadual": valor_icms + valor_difal,
                "chave_nf_original": chave_original, "caminho_arquivo": caminho_xml
            }
        except Exception: return None

    def _build_aba_restituicoes(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        frame_top = tk.Frame(parent, bg=surface, pady=15, padx=20)
        frame_top.pack(fill="x", pady=10)
        tk.Label(frame_top, text="🏛️ Dossiê de Restituição de GNRE / ICMS", font=("Segoe UI", 14, "bold"), bg=surface, fg=accent).pack(anchor="w")
        tk.Label(frame_top, text="Cruza Notas Fiscais de Devolução com GNREs pagas e envia automaticamente.", font=("Segoe UI", 10), bg=surface, fg=muted).pack(anchor="w", pady=(0, 10))

        self.btn_gerar_dossie = tk.Button(
            frame_top, text="🔍 Analisar Notas de Devolução (XML) e Gerar Dossiês", 
            font=("Segoe UI", 11, "bold"), bg="#f59e0b", fg="#0f172a", 
            command=self._iniciar_analise_restituicao, relief="flat", padx=15, pady=8, cursor="hand2"
        )
        self.btn_gerar_dossie.pack(anchor="w")
        self.log_restituicao = scrolledtext.ScrolledText(parent, height=25, font=("Consolas", 10), bg="#0a0f1e", fg=text, insertbackground=accent)
        self.log_restituicao.pack(fill="both", expand=True, padx=20, pady=10)

    def _ui_log_rest(self, mensagem):
        """Método seguro para atualizar a interface a partir da thread."""
        self.root.after(0, lambda: self.log_restituicao.insert(tk.END, mensagem))
        self.root.after(0, lambda: self.log_restituicao.see(tk.END))

    def _iniciar_analise_restituicao(self):
        self.btn_gerar_dossie.config(state="disabled", text="⏳ A analisar...")
        self.log_restituicao.insert(tk.END, "A iniciar varredura de XMLs de devolução...\n")
        self.root.update()
        threading.Thread(target=self._processar_dossies_thread, daemon=True).start()

    def _processar_dossies_thread(self):
        try:
            pasta_xmls = os.path.join(PASTA_ATUAL, "XML_NOTAS") 
            if not os.path.exists(pasta_xmls):
                os.makedirs(pasta_xmls)
                self._ui_log_rest(f"❌ Pasta de XMLs criada em: {pasta_xmls}\nColoque os ficheiros e tente novamente.\n")
                return

            self._ui_log_rest("1. A extrair dados dos ficheiros XML...\n")
            notas_devolucao = [self.ler_xml_nfe(os.path.join(r, f)) for r, _, fs in os.walk(pasta_xmls) for f in fs if f.lower().endswith(".xml")]
            notas_devolucao = [n for n in notas_devolucao if n and n["finalidade"] == "DEVOLUCAO"]

            if not notas_devolucao:
                self._ui_log_rest("⚠️ Nenhuma nota de devolução encontrada.\n")
                return

            self._ui_log_rest(f"➜ Encontradas {len(notas_devolucao)} notas de devolução. A procurar GNREs pagas...\n")
            
            dossies_gerados = []
            
            for nf in notas_devolucao:
                if nf["valor_imposto_estadual"] <= 0: continue
                
                chave = nf["chave_nf_original"] or ""
                self._ui_log_rest(f"\n➜ Analisando Devolução NF {nf['nf_numero']} ({nf['uf_destino']}) - Valor a restituir: R$ {nf['valor_imposto_estadual']:.2f}\n")
                
                gnre_encontrada = None
                if chave:
                    # Busca a GNRE paga na pasta GNRE_SEFAZ
                    gnre_dir = os.path.join(PASTA_SAIDA, "GNRE_SEFAZ", nf['uf_destino'])
                    if os.path.exists(gnre_dir):
                        for f in os.listdir(gnre_dir):
                            if chave in f:
                                gnre_encontrada = os.path.join(gnre_dir, f)
                                break
                    
                    if not gnre_encontrada:
                        # Busca genérica em toda a PASTA_SAIDA
                        for root_dir, _, files in os.walk(PASTA_SAIDA):
                            for f in files:
                                if f.lower().endswith('.pdf') and chave in f:
                                    gnre_encontrada = os.path.join(root_dir, f)
                                    break
                            if gnre_encontrada: break

                if gnre_encontrada:
                    self._ui_log_rest(f"   ✅ GNRE Correspondente encontrada: {os.path.basename(gnre_encontrada)}\n")
                    dossies_gerados.append({
                        "nf": nf,
                        "gnre_pdf": gnre_encontrada
                    })
                else:
                    self._ui_log_rest(f"   ⚠️ GNRE com a chave {chave} não encontrada no arquivo.\n")

            if dossies_gerados:
                self._ui_log_rest("\n2. A enviar dossiês por e-mail...\n")
                
                emails_anexos = [d["gnre_pdf"] for d in dossies_gerados]
                emails_anexos.extend([d["nf"]["caminho_arquivo"] for d in dossies_gerados])
                
                corpo = "Segue em anexo o dossiê de restituição de impostos (GNRE e XMLs correspondentes):\n\n"
                for d in dossies_gerados:
                    corpo += f"- NF {d['nf']['nf_numero']} ({d['nf']['uf_destino']}): R$ {d['nf']['valor_imposto_estadual']:.2f}\n"
                
                try:
                    enviar_email(EMAIL_DESTINO_CONTAB, "Dossiê de Restituição SEFAZ", corpo, emails_anexos)
                    self._ui_log_rest("   ✅ E-mail enviado com sucesso para a contabilidade!\n")
                except Exception as e_mail:
                    self._ui_log_rest(f"   ⚠️ Não foi possível enviar e-mail: {e_mail}\n")
                    zip_path = os.path.join(PASTA_ATUAL, f"Dossie_Restituicao_{datetime.now().strftime('%Y%m%d')}.zip")
                    import zipfile
                    with zipfile.ZipFile(zip_path, 'w') as zf:
                        for arq in emails_anexos:
                            zf.write(arq, os.path.basename(arq))
                    self._ui_log_rest(f"   💾 Arquivos salvos localmente num ZIP seguro: {zip_path}\n")

            self._ui_log_rest("\n✨ Processo finalizado com sucesso!\n")
            
        except Exception as e:
            self._ui_log_rest(f"❌ Erro crítico: {str(e)}\n")
        finally:
            self.root.after(0, lambda: self.btn_gerar_dossie.config(state="normal", text="🔍 Analisar Notas de Devolução (XML) e Gerar Dossiês"))

# ==============================================================================
# --- MAIN ---
# ==============================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = RoboApp(root)
    root.mainloop()
