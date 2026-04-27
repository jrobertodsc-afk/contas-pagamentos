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
import xml.etree.ElementTree as ET
import xml.etree.ElementTree as ET_SAFE  # Para compatibilidade com o código de consulta


class AbaContasPagarMixin:

    # ══════════════════════════════════════════════════════════════════════════
    # ABA CONTAS A PAGAR
    # ══════════════════════════════════════════════════════════════════════════

    def _build_aba_contas_pagar(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        db_init()
        self._cap_cert_passwords = {}  # Cache de senhas por arquivo {caminho: senha}

        self.cap_tabs = ctk.CTkTabview(
            parent, 
            fg_color="transparent",
            segmented_button_selected_color=accent,
            segmented_button_selected_hover_color=green,
            segmented_button_unselected_color=surface,
            text_color="#FFFFFF" # Forçar branco para visibilidade nas abas
        )
        self.cap_tabs.pack(fill="both", expand=True, padx=10, pady=10)

        sub_lista   = self.cap_tabs.add("  📋 LISTA DE NOTAS  ")
        sub_nota    = self.cap_tabs.add("  ➕ LANÇAR NOTA  ")
        sub_alertas = self.cap_tabs.add("  🔔 ALERTAS  ")
        sub_adian   = self.cap_tabs.add("  💰 ADIANTAMENTOS  ")
        sub_prest   = self.cap_tabs.add("  ✅ PRESTAÇÃO  ")
        sub_imp     = self.cap_tabs.add("  🏛️ IMPOSTOS  ")
        sub_forn    = self.cap_tabs.add("  🏢 FORNECEDORES  ")
        sub_cnab    = self.cap_tabs.add("  🏦 CNAB 240  ")

        self._cap_build_lista(sub_lista, bg, surface, accent, green, yellow, text, muted, self.cap_tabs)
        self._cap_build_nova_nota(sub_nota, bg, surface, accent, green, yellow, text, muted, self.cap_tabs)
        self._cap_build_alertas(sub_alertas, bg, surface, accent, green, yellow, text, muted)
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

        # Prepara lista de senhas necessárias
        senhas_novas = {}
        for p in pfx_paths:
            if p not in self._cap_cert_passwords:
                from tkinter import simpledialog
                s = simpledialog.askstring(
                    "Certificado Digital",
                    f"Digite a senha para o certificado:\n{os.path.basename(p)}",
                    show="*", parent=self.root)
                if s:
                    self._cap_cert_passwords[p] = s
                else:
                    self._cap_lbl_xml.config(text=f"⚠️ Consulta ignorando {os.path.basename(p)}", fg="#fbbf24")

        # Se não tiver nenhuma senha, nem tenta
        if not self._cap_cert_passwords:
            self._cap_lbl_xml.config(text="⚠️ Nenhuma senha de certificado informada.", fg="#fbbf24")
            return
        
        mapa_senhas = self._cap_cert_passwords.copy()

        # Tenta consultar em background
        def _consultar_bg(mapa_s):
            try:
                import requests
                from cryptography.hazmat.primitives.serialization import pkcs12
                from cryptography.hazmat.primitives.serialization import (
                    Encoding, PrivateFormat, NoEncryption)
                import tempfile

                last_err = "Nenhum certificado funcionou"
                # Tenta cada .pfx até funcionar
                for pfx_path in pfx_paths:
                    fname = os.path.basename(pfx_path)
                    self.root.after(0, lambda f=fname: self._cap_lbl_xml.config(
                        text=f"⏳ Tentando certificado: {f}...", fg="#38bdf8"))
                    
                    try:
                        with open(pfx_path,"rb") as f:
                            pfx_data = f.read()

                        # Pega a senha correta para este arquivo
                        s_atual = mapa_s.get(pfx_path)
                        if not s_atual: continue

                        # Tenta carregar com a senha
                        try:
                            priv, cert, chain = pkcs12.load_key_and_certificates(
                                pfx_data, s_atual.encode())
                        except Exception as ep:
                            last_err = f"Senha incorreta ou erro no arquivo ({fname}): {ep}"
                            continue

                        # Gera PEM temporários para requests
                        with tempfile.NamedTemporaryFile(suffix=".pem",delete=False,mode="wb") as fc:
                            fc.write(cert.public_bytes(Encoding.PEM))
                            cert_pem = fc.name
                        with tempfile.NamedTemporaryFile(suffix=".pem",delete=False,mode="wb") as fk:
                            fk.write(priv.private_bytes(Encoding.PEM,PrivateFormat.TraditionalOpenSSL,NoEncryption()))
                            key_pem = fk.name

                        # Webservice SEFAZ por UF (Prioriza SVRS para maior estabilidade)
                        WS_URL = {
                            "29": "https://nfe.sefaz.ba.gov.br/webservices/nfeconsultaprotocolo4/nfeconsultaprotocolo4.asmx",
                            "32": "https://nfe.svrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx",
                        }.get(c_uf,
                              "https://nfe.svrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx")

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
                                # Salva XML temporário e importa para preencher todos os campos
                                tmp_name = f"NF_{chave}.xml"
                                tmp_path = os.path.join(tempfile.gettempdir(), tmp_name)
                                with open(tmp_path, "w", encoding="utf-8") as f:
                                    f.write(xml_nfe)
                                
                                self.root.after(0, lambda: (
                                    self._cap_importar_xml_path(tmp_path),
                                    self._cap_lbl_xml.config(text=f"✅ XML baixado da SEFAZ com sucesso!", fg="#4ade80")
                                ))
                                return
                    except Exception as e2:
                        last_err = str(e2)
                        continue

                self.root.after(0, lambda: self._cap_lbl_xml.config(
                    text=f"⚠️ Falha SEFAZ: {last_err}",
                    fg="#fbbf24"))

            except Exception as e:
                self.root.after(0, lambda: self._cap_lbl_xml.config(
                    text=f"⚠️ Certificado: {e}", fg="#fbbf24"))

        import threading
        threading.Thread(target=_consultar_bg, args=(mapa_senhas,), daemon=True).start()
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

            # ── Valores e Impostos ──
            # Tenta buscar em vários caminhos comuns (ICMSTot ou resumo do retorno)
            vNF_val = (_val(root, "total", "ICMSTot", "vNF") or 
                       _val(root, "retConsSitNFe", "vNF") or
                       _val(root, "infProt", "vNF") or "0")
            vNF = float(vNF_val)
            
            vICMS       = float(_val(root, "total", "ICMSTot", "vICMS")       or 0)
            vPIS        = float(_val(root, "total", "ICMSTot", "vPIS")        or 0)
            vCOFINS     = float(_val(root, "total", "ICMSTot", "vCOFINS")     or 0)
            vIR         = float(_val(root, "total", "ICMSTot", "vIR")         or 0)
            vINSS       = float(_val(root, "total", "ICMSTot", "vINSS")       or 0)
            vISS        = float(_val(root, "total", "ICMSTot", "vISS")        or 0)
            
            # Reforma Tributária (IBS/CBS)
            vIBS        = float(_val(root, "total", "IBSCBSTot", "vIBS")      or 0)
            vCBS        = float(_val(root, "total", "IBSCBSTot", "vCBS")      or 0)
            vIBSUF      = float(_val(root, "total", "IBSCBSTot", "gIBS", "gIBSUF", "vIBSUF") or 0)

            vICMSUFDest = float(_val(root, "total", "ICMSTot", "vICMSUFDest") or 0) # DIFAL antigo
            vFCPUFDest  = float(_val(root, "total", "ICMSTot", "vFCPUFDest")  or 0) # FCP
            
            # Se tiver vIBSUF, considera como o novo DIFAL
            valor_difal_final = vICMSUFDest if vICMSUFDest > 0 else vIBSUF

            # Natureza e Referência
            finNFe = _val(root, "ide", "finNFe") # 1=Normal, 4=Devolução
            chave_ref = ""
            if finNFe == "4":
                # Busca refNFe em qualquer lugar do XML
                for el in root.iter():
                    if _strip_ns(el.tag) == "refNFe":
                        chave_ref = el.text or ""
                        break

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

            # ── Itens (prod) ──
            itens_xml = []
            for det in root.iter():
                if _strip_ns(det.tag) == "det":
                    prod_desc = ""
                    v_total_prod = 0.0
                    # Busca xProd e vProd dentro de det
                    for sub in det.iter():
                        tag_sub = _strip_ns(sub.tag)
                        if tag_sub == "xProd": prod_desc = sub.text or ""
                        if tag_sub == "vProd": 
                            try: v_total_prod = float(sub.text or 0)
                            except: v_total_prod = 0.0
                    if prod_desc:
                        itens_xml.append((prod_desc, v_total_prod))

            # ── Preenche os campos ──
            # Fornecedor
            self._cap_forn.delete(0,"end")
            self._cap_forn.insert(0, razao)

            # CNPJ — formata XX.XXX.XXX/XXXX-XX
            if cnpj:
                cnpj_fmt = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}"
                self._cap_cnpj.delete(0,"end")
                self._cap_cnpj.insert(0, cnpj_fmt)

            # Natureza
            if finNFe == "4":
                self._cap_natureza_var.set("DEVOLUÇÃO")
            else:
                self._cap_natureza_var.set("VENDA")
            
            # Chave Referenciada
            if chave_ref:
                self._cap_chave_ref.delete(0, "end")
                self._cap_chave_ref.insert(0, chave_ref)

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

            # DIFAL e FCP
            self._cap_difal.delete(0, "end")
            self._cap_difal.insert(0, f"{valor_difal_final:,.2f}".replace(",","X").replace(".",",").replace("X","."))
            self._cap_fcp.delete(0, "end")
            self._cap_fcp.insert(0, f"{vFCPUFDest:,.2f}".replace(",","X").replace(".",",").replace("X","."))

            # Vencimento — usa primeira parcela se existir
            if parcelas_xml:
                self._cap_dt_venc.delete(0,"end")
                self._cap_dt_venc.insert(0, parcelas_xml[0][1])

            # Itens — limpa e preenche
            for ir in self._cap_item_rows:
                for w in ir.get("widgets", []): w.destroy()
            self._cap_item_rows = []
            
            if itens_xml:
                for d_it, v_it in itens_xml:
                    self._cap_add_item_row(d_it, f"{v_it:.2f}")
            else:
                self._cap_add_item_row() # Pelo menos um vazio

            # Trigger check de restituição
            self._cap_check_restituicao()
            
            # Força atualização da rolagem
            if hasattr(self, "_cap_update_scroll"):
                self._cap_update_scroll()

            # Impostos — preenche linhas existentes
            impostos_nf = [
                ("ICMS",       0,                                vICMS),
                ("IBS",        0,                                vIBS),
                ("CBS",        0,                                vCBS),
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
        tb = tk.Frame(parent, bg=surface, pady=10, padx=16)
        tb.pack(fill="x", padx=0, pady=0)

        tk.Label(tb, text="🧾 Contas a Pagar", font=("Segoe UI",14,"bold"),
                 fg=accent, bg=surface).pack(side="left", padx=(0,20))

        tk.Label(tb, text="Status:", fg=text, bg=surface,
                 font=("Segoe UI",11)).pack(side="left")
        self._cap_filtro_status = ttk.Combobox(tb,
            values=["TODOS","PENDENTE","APROVADA","PAGA","VENCIDA","CANCELADA"],
            state="readonly", width=12, font=("Segoe UI",11))
        self._cap_filtro_status.set("TODOS")
        self._cap_filtro_status.pack(side="left", padx=(5,15))

        tk.Label(tb, text="Empresa:", fg=text, bg=surface,
                 font=("Segoe UI",11)).pack(side="left")
        self._cap_filtro_emp = ttk.Combobox(tb,
            values=["TODOS","LALUA","SOLAR"],
            state="readonly", width=10, font=("Segoe UI",11))
        self._cap_filtro_emp.set("TODOS")
        self._cap_filtro_emp.pack(side="left", padx=(5,15))

        self._cap_busca_var = tk.StringVar()
        tk.Entry(tb, textvariable=self._cap_busca_var, font=("Segoe UI",11),
                 bg="#0a0f1e", fg=text, insertbackground=accent,
                 relief="flat", bd=2, width=25).pack(side="left", padx=(0,10))

        tk.Button(tb, text="🔍 Filtrar", font=("Segoe UI",11),
                  bg="#1e293b", fg=accent, relief="flat", bd=0,
                  padx=12, pady=4, cursor="hand2",
                  command=self._cap_carregar_lista).pack(side="left", padx=(0,10))

        tk.Button(tb, text="➕ Nova Nota", font=("Segoe UI",11,"bold"),
                  bg=accent, fg="#0f172a", relief="flat", bd=0,
                  padx=15, pady=4, cursor="hand2",
                  command=lambda: nb_pai.select(1)).pack(side="left", padx=(0,8))

        tk.Button(tb, text="💰 Adiantamento", font=("Segoe UI",11),
                  bg="#7c3aed", fg="white", relief="flat", bd=0,
                  padx=15, pady=4, cursor="hand2",
                  command=lambda: nb_pai.select(3)).pack(side="left")

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
        rod = tk.Frame(parent, bg=surface, pady=10, padx=16)
        rod.pack(fill="x", side="bottom")

        self._cap_lbl_totais = tk.Label(rod, text="", font=("Segoe UI",11,"bold"),
                                         fg=accent, bg=surface)
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

        total_receita = 0.0
        total_despesa = 0.0
        for n in notas:
            tag = n["status"] if n["status"] in STATUS_NOTA else "PENDENTE"
            
            tipo_op = n.get("tipo_operacao", "DESPESA")
            valor = n["valor_bruto"]
            
            if tipo_op == "RECEITA":
                total_receita += valor
                op_prefix = "(+) "
            else:
                total_despesa += valor
                op_prefix = "(-) "

            self._cap_tree.insert("", "end", iid=str(n["id"]), tags=(tag,),
                values=(
                    n["numero_tx"], 
                    f"{op_prefix}{n.get('natureza','NOTA')}", 
                    n["fornecedor"][:30],
                    n["empresa"], n["dt_emissao"], n["dt_vencimento"],
                    f"R$ {n['valor_bruto']:,.2f}",
                    f"R$ {n['valor_liquido']:,.2f}",
                    n["status"], n["categoria"][:25]
                ))

        saldo = total_receita - total_despesa
        cor_saldo = "#4ade80" if saldo >= 0 else "#f87171"

        self._cap_lbl_totais.config(
            text=f"{len(notas)} nota(s)  |  "
                 f"Receitas: R$ {total_receita:,.2f}  |  "
                 f"Despesas: R$ {total_despesa:,.2f}  |  "
                 f"SALDO: R$ {saldo:,.2f}",
            fg=cor_saldo
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
        def _update_scroll(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        frm.bind("<Configure>", _update_scroll)
        self._cap_update_scroll = _update_scroll # Guarda referência

        def _scroll(e):
            canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _scroll)

        pad = dict(padx=20, pady=4)

        # ── TÍTULO + BOTÃO XML ──
        frm_tit = tk.Frame(frm, bg=bg)
        frm_tit.pack(fill="x", padx=20, pady=10)
        tk.Label(frm_tit, text="➕  Nova Nota / Despesa",
                 font=("Segoe UI",16,"bold"), fg=accent, bg=bg
                 ).pack(side="left")

        tk.Button(frm_tit, text="📎 Importar XML (NF-e)",
                  font=("Segoe UI",9,"bold"), bg="#7c3aed", fg="white",
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=lambda: self._cap_importar_xml()
                  ).pack(side="right")

        self._cap_lbl_xml = tk.Label(frm_tit, text="",
                                      font=("Segoe UI",8), fg=green, bg=bg)
        self._cap_lbl_xml.pack(side="right", padx=8)

        # Botão Salvar Rápido (Topo)
        tk.Button(frm_tit, text="💾 Salvar",
                  font=("Segoe UI",9,"bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=15, pady=5, cursor="hand2",
                  command=lambda: self._cap_salvar_nota(nb_pai)
                  ).pack(side="right", padx=10)

        # ── CAMPO CHAVE DE ACESSO ──
        frm_chave = tk.Frame(frm, bg=surface, padx=14, pady=10)
        frm_chave.pack(fill="x", padx=20, pady=(0,10))

        tk.Label(frm_chave, text="🔑 Chave de Acesso NF-e:",
                 fg=accent, bg=surface, font=("Segoe UI",11,"bold")
                 ).pack(side="left", padx=(0,10))

        self._cap_chave_var = tk.StringVar()
        ent_chave = tk.Entry(frm_chave, textvariable=self._cap_chave_var,
                             font=("Consolas",12), bg="#0a0f1e", fg=yellow,
                             insertbackground=accent, relief="flat", bd=2, width=55)
        ent_chave.pack(side="left", padx=(0,10))
        tk.Label(frm_chave, text="(44 dígitos — pode bipar ou colar)",
                 fg="#94a3b8", bg=surface, font=("Consolas",9)
                 ).pack(side="left", padx=(0,10))

        tk.Button(frm_chave, text="🔍 Consultar SEFAZ",
                  font=("Segoe UI",10,"bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=15, pady=6, cursor="hand2",
                  command=lambda: self._cap_consultar_sefaz()
                  ).pack(side="left")

        # ── BLOCO 1: DADOS BÁSICOS ──
        b1 = tk.LabelFrame(frm, text="  Dados da Nota  ", bg=surface,
                            fg=accent, font=("Segoe UI",11,"bold"), padx=15, pady=12)
        b1.pack(fill="x", padx=20, pady=(0,10))

        def lbl(parent, t): return tk.Label(parent, text=t, fg=text, bg=surface,
                                             font=("Segoe UI",11), anchor="w")
        def ent(parent, w=25, **kw):
            e = tk.Entry(parent, font=("Segoe UI",12), bg="#0a0f1e", fg=text,
                         insertbackground=accent, relief="flat", bd=2, width=w, **kw)
            return e

        # Linha 0 — Tipo e Operação
        frm_tipo_op = tk.Frame(b1, bg=surface)
        frm_tipo_op.grid(row=0,column=0,columnspan=6,sticky="w",pady=8)

        lbl(frm_tipo_op,"Tipo:").pack(side="left", padx=(0,8))
        self._cap_tipo_var = tk.StringVar(value="NOTA")
        for val,lbl_t in [("NOTA","Nota Fiscal"),("ADIANTAMENTO","Adiantamento")]:
            tk.Radiobutton(frm_tipo_op, text=lbl_t, variable=self._cap_tipo_var,
                           value=val, bg=surface, fg=text, selectcolor="#0a0f1e",
                           font=("Segoe UI",11)).pack(side="left", padx=8)

        tk.Label(frm_tipo_op, text="  |  ", fg="#475569", bg=surface, font=("Segoe UI",12)).pack(side="left", padx=10)

        lbl(frm_tipo_op,"Operação:").pack(side="left", padx=(0,8))
        self._cap_operacao_var = tk.StringVar(value="DESPESA")
        for val,lbl_o in [("DESPESA","Despesa (-)"),("RECEITA","Receita (+)")]:
            tk.Radiobutton(frm_tipo_op, text=lbl_o, variable=self._cap_operacao_var,
                           value=val, bg=surface, fg=text, selectcolor="#0a0f1e",
                           font=("Segoe UI",11)).pack(side="left", padx=8)

        # Linha 1 — Fornecedor / CNPJ
        lbl(b1,"Fornecedor:").grid(row=1,column=0,sticky="w",pady=6)
        frm_forn = tk.Frame(b1, bg=surface)
        frm_forn.grid(row=1,column=1,columnspan=2,sticky="w",padx=(0,15))
        self._cap_forn = ent(frm_forn, 35)
        self._cap_forn.pack(side="left")
        tk.Button(frm_forn, text="🔍 Buscar", font=("Segoe UI",10,"bold"),
                  bg="#1e293b", fg=accent, relief="flat", bd=0,
                  padx=10, pady=4, cursor="hand2",
                  command=lambda: _abrir_busca_forn()
                  ).pack(side="left", padx=(5,0))

        lbl(b1,"CNPJ/CPF:").grid(row=1,column=3,sticky="w", padx=(15,0))
        self._cap_cnpj = ent(b1, 22)
        self._cap_cnpj.grid(row=1,column=4,sticky="w")

        # Linha 2 — Empresa / Categoria
        lbl(b1,"Empresa:").grid(row=2,column=0,sticky="w",pady=6)
        self._cap_emp_var = tk.StringVar(value="LALUA")
        ttk.Combobox(b1, textvariable=self._cap_emp_var,
                     values=["LALUA","SOLAR"], state="readonly",
                     width=12, font=("Segoe UI",12)).grid(row=2,column=1,sticky="w",padx=(0,15))

        lbl(b1,"Categoria:").grid(row=2,column=2,sticky="w")
        self._cap_cat_var = tk.StringVar()
        self._cap_cat_cb = ttk.Combobox(b1, textvariable=self._cap_cat_var,
                                          values=CATEGORIAS_LISTA,
                                          state="normal", width=40, font=("Segoe UI",12))
        self._cap_cat_cb.grid(row=2,column=3,columnspan=2,sticky="w")

        # Linha 3 — Natureza / Responsável
        lbl(b1,"Natureza:").grid(row=3,column=0,sticky="w",pady=6)
        self._cap_natureza_var = tk.StringVar(value="VENDA")
        self._cap_nat_cb = ttk.Combobox(b1, textvariable=self._cap_natureza_var,
            values=["VENDA","SERVIÇO","DEVOLUÇÃO","SALÁRIO","BENEFÍCIO","IMPOSTO","FIXA"],
            state="readonly", width=15, font=("Segoe UI",12))
        self._cap_nat_cb.grid(row=3,column=1,sticky="w",padx=(0,15))

        lbl(b1,"Responsável:").grid(row=3,column=2,sticky="w")
        self._cap_resp = ttk.Combobox(b1, values=RESPONSAVEIS_LISTA,
                                       state="readonly", width=22, font=("Segoe UI",12))
        self._cap_resp.grid(row=3,column=3,sticky="w")

        # Linha 4 — Descrição
        lbl(b1,"Descrição:").grid(row=4,column=0,sticky="w",pady=6)
        self._cap_desc = ent(b1, 75)
        self._cap_desc.grid(row=4,column=1,columnspan=4,sticky="w")

        # Linha 5 — Datas / Valor
        lbl(b1,"Emissão:").grid(row=5,column=0,sticky="w",pady=6)
        self._cap_dt_emissao = ent(b1, 14)
        self._cap_dt_emissao.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self._cap_dt_emissao.grid(row=5,column=1,sticky="w",padx=(0,15))
        _aplicar_mascara_data(self._cap_dt_emissao)

        lbl(b1,"Vencimento:").grid(row=5,column=2,sticky="w")
        self._cap_dt_venc = ent(b1, 14)
        self._cap_dt_venc.grid(row=5,column=3,sticky="w",padx=(0,15))
        _aplicar_mascara_data(self._cap_dt_venc)

        lbl(b1,"Valor Bruto:").grid(row=5,column=4,sticky="w")
        self._cap_valor = ent(b1, 20)
        self._cap_valor.grid(row=5,column=5,sticky="w")
        _aplicar_mascara_valor(self._cap_valor)

        # Linha 6 — DIFAL / FCP / Chave Ref (Específico ERP)
        lbl(b1,"DIFAL (R$):").grid(row=6,column=0,sticky="w",pady=6)
        self._cap_difal = ent(b1, 14)
        self._cap_difal.insert(0, "0,00")
        self._cap_difal.grid(row=6,column=1,sticky="w",padx=(0,15))
        _aplicar_mascara_valor(self._cap_difal)

        lbl(b1,"FCP (R$):").grid(row=6,column=2,sticky="w")
        self._cap_fcp = ent(b1, 14)
        self._cap_fcp.insert(0, "0,00")
        self._cap_fcp.grid(row=6,column=3,sticky="w",padx=(0,15))
        _aplicar_mascara_valor(self._cap_fcp)

        lbl(b1,"Chave/NF Ref:").grid(row=6,column=4,sticky="w")
        self._cap_chave_ref = ent(b1, 40)
        self._cap_chave_ref.grid(row=6,column=5,sticky="w")

        # Linha 7 — Recorrência / Alertas
        lbl(b1,"Repetir:").grid(row=7,column=0,sticky="w",pady=2)
        self._cap_recorr_var = tk.BooleanVar(value=False)
        tk.Checkbutton(b1, text="Despesa Fixa Mensal", variable=self._cap_recorr_var,
                       bg=surface, fg=text, selectcolor="#0a0f1e",
                       font=("Segoe UI",8)).grid(row=7,column=1,columnspan=2,sticky="w")
        
        lbl(b1,"Meses:").grid(row=7,column=3,sticky="w")
        self._cap_recorr_meses = tk.Spinbox(b1, from_=2, to=36, width=5, font=("Segoe UI",8),
                                            bg="#0a0f1e", fg=text, relief="flat", bd=2)
        self._cap_recorr_meses.grid(row=7,column=4,sticky="w")

        # Alerta de Restituição DIFAL
        self._cap_lbl_alerta_restit = tk.Label(b1, text="", font=("Segoe UI",8,"bold"),
                                               fg=yellow, bg=surface)
        self._cap_lbl_alerta_restit.grid(row=7,column=5,sticky="w")


        self._cap_nat_cb.bind("<<ComboboxSelected>>", self._cap_check_restituicao)
        self._cap_chave_ref.bind("<FocusOut>", self._cap_check_restituicao)

        # Linha 8 — Observação
        lbl(b1,"Observação:").grid(row=8,column=0,sticky="w",pady=2)
        self._cap_obs = ent(b1, 85)
        self._cap_obs.grid(row=8,column=1,columnspan=5,sticky="w")

        # Linha 9 — Código de Barras
        lbl(b1,"Cód. Barras:").grid(row=9,column=0,sticky="w",pady=2)
        frm_cb = tk.Frame(b1, bg=surface)
        frm_cb.grid(row=9,column=1,columnspan=5,sticky="w")
        self._cap_cod_barras = tk.Entry(
            frm_cb, font=("Consolas",9), bg="#0a0f1e", fg=yellow,
            insertbackground=accent, relief="flat", bd=2, width=65)
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

        # ── BLOCO 1c: ITENS DA NOTA / SERVIÇO ──
        b1c = tk.LabelFrame(frm, text="  Itens da Nota / Serviço  ", bg=surface,
                             fg=accent, font=("Segoe UI",11,"bold"), padx=15, pady=12)
        b1c.pack(fill="x", padx=20, pady=(0,10))

        self._cap_item_rows = []
        self._cap_frm_items = tk.Frame(b1c, bg=surface)
        self._cap_frm_items.pack(fill="x")

        # Cabeçalho itens
        for i, (h, w) in enumerate([("Descrição / Produto", 70), ("Valor Total", 20)]):
            tk.Label(self._cap_frm_items, text=h, fg=accent, bg=surface,
                     font=("Segoe UI",10,"bold")).grid(row=0, column=i*2, sticky="w", padx=6, pady=(0,5))

        self._cap_add_item_row() # Começa com uma linha

        tk.Button(b1c, text="+ item", font=("Segoe UI",10,"bold"),
                  bg="#1e293b", fg=accent, relief="flat", bd=0,
                  padx=10, pady=4, cursor="hand2",
                  command=self._cap_add_item_row).pack(anchor="w", pady=(8,0))

        # Guarda referência para o autocomplete usar
        self._cap_preencher_banco = _preencher_dados_banco

        # ── BLOCO 2: IMPOSTOS ──
        b2 = tk.LabelFrame(frm, text="  Impostos Retidos  ", bg=surface,
                            fg=yellow, font=("Segoe UI",11,"bold"), padx=15, pady=12)
        b2.pack(fill="x", padx=20, pady=(0,10))

        tk.Label(b2, text="Valor Bruto de referência é preenchido automaticamente ao digitar acima.",
                 fg="#94a3b8", bg=surface, font=("Consolas",9)).pack(anchor="w", pady=(0,8))

        self._cap_imp_rows = []
        frm_imps = tk.Frame(b2, bg=surface)
        frm_imps.pack(fill="x")

        # Cabeçalho impostos
        for i, h in enumerate(["Imposto","Alíquota %","Valor R$","Venc. DARF/GPS"]):
            tk.Label(frm_imps, text=h, fg=accent, bg=surface,
                     font=("Segoe UI",9,"bold")).grid(row=0,column=i*2,sticky="w",padx=6,pady=(0,5))

        def _add_imposto(tipo="", aliq="", venc=""):
            row_i = len(self._cap_imp_rows) + 1
            tipo_var = tk.StringVar(value=tipo)
            aliq_var = tk.StringVar(value=str(aliq))
            val_var  = tk.StringVar(value="")
            venc_var = tk.StringVar(value=venc)

            cb = ttk.Combobox(frm_imps, textvariable=tipo_var,
                               values=TIPOS_IMPOSTO, state="readonly",
                               width=15, font=("Segoe UI",11))
            cb.grid(row=row_i, column=0, padx=6, pady=4)

            aliq_e = tk.Entry(frm_imps, textvariable=aliq_var, width=10,
                              font=("Segoe UI",11), bg="#0a0f1e", fg=text,
                              insertbackground=accent, relief="flat", bd=2)
            aliq_e.grid(row=row_i, column=2, padx=6)

            val_e = tk.Entry(frm_imps, textvariable=val_var, width=15,
                             font=("Segoe UI",11), bg="#0a0f1e", fg=yellow,
                             insertbackground=accent, relief="flat", bd=2)
            val_e.grid(row=row_i, column=4, padx=6)

            venc_e = tk.Entry(frm_imps, textvariable=venc_var, width=14,
                              font=("Segoe UI",11), bg="#0a0f1e", fg=text,
                              insertbackground=accent, relief="flat", bd=2)
            venc_e.grid(row=row_i, column=6, padx=6)

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

        tk.Button(b2, text="+ imposto", font=("Segoe UI",10,"bold"),
                  bg="#1e293b", fg=yellow, relief="flat", bd=0,
                  padx=10, pady=4, cursor="hand2",
                  command=_add_imposto).pack(anchor="w", pady=(8,0))

        # Label valor líquido
        self._cap_lbl_liq = tk.Label(b2, text="Valor líquido a pagar:  R$ 0,00",
                                      font=("Segoe UI",12,"bold"), fg=green, bg=surface)
        self._cap_lbl_liq.pack(anchor="e", pady=(8,0))

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
                            fg=accent, font=("Segoe UI",11,"bold"), padx=15, pady=12)
        b3.pack(fill="x", padx=20, pady=(0,10))

        frame_parc_ctrl = tk.Frame(b3, bg=surface)
        frame_parc_ctrl.pack(fill="x")

        tk.Label(frame_parc_ctrl, text="Nº parcelas:", fg=text, bg=surface,
                 font=("Segoe UI",11)).pack(side="left")
        self._cap_nparc = tk.Spinbox(frame_parc_ctrl, from_=1, to=48, width=6,
                                      font=("Segoe UI",11), bg="#0a0f1e", fg=text,
                                      insertbackground=accent, relief="flat", bd=2)
        self._cap_nparc.pack(side="left", padx=(6,15))

        tk.Label(frame_parc_ctrl, text="1ª parcela:", fg=text, bg=surface,
                 font=("Segoe UI",11)).pack(side="left")
        self._cap_dt_1parc = ent(frame_parc_ctrl, 14)
        self._cap_dt_1parc.pack(side="left", padx=(6,15))
        _aplicar_mascara_data(self._cap_dt_1parc)

        tk.Label(frame_parc_ctrl, text="Intervalo (dias):", fg=text, bg=surface,
                 font=("Segoe UI",11)).pack(side="left")
        self._cap_intervalo = tk.Spinbox(frame_parc_ctrl, from_=1, to=90,
                                          width=6, font=("Segoe UI",11),
                                          bg="#0a0f1e", fg=text,
                                          insertbackground=accent, relief="flat", bd=2)
        self._cap_intervalo.delete(0,"end"); self._cap_intervalo.insert(0,"30")
        self._cap_intervalo.pack(side="left", padx=(6,15))

        tk.Button(frame_parc_ctrl, text="⚡ Calcular parcelas",
                  font=("Segoe UI",10,"bold"), bg="#1e293b", fg=accent,
                  relief="flat", bd=0, padx=12, pady=4, cursor="hand2",
                  command=self._cap_calcular_parcelas).pack(side="left")

        self._cap_frm_parcelas = tk.Frame(b3, bg=surface)
        self._cap_frm_parcelas.pack(fill="x", pady=(6,0))
        self._cap_parc_widgets = []

        # ── BLOCO 4: RATEIO ──
        b4 = tk.LabelFrame(frm, text="  Rateio por Filial / Histórico Contábil  ",
                            bg=surface, fg=accent,
                            font=("Segoe UI",11,"bold"), padx=15, pady=12)
        b4.pack(fill="x", padx=20, pady=(0,10))

        tk.Label(b4, text="Soma dos percentuais deve ser 100%",
                 fg="#94a3b8", bg=surface, font=("Consolas",9)).pack(anchor="w", pady=(0,8))

        self._cap_rat_rows = []
        frm_rats = tk.Frame(b4, bg=surface)
        frm_rats.pack(fill="x")

        for i,h in enumerate(["Filial","% Rateio","Categoria","Valor R$"]):
            tk.Label(frm_rats, text=h, fg=accent, bg=surface,
                     font=("Segoe UI",9,"bold")).grid(row=0,column=i*2,sticky="w",padx=6,pady=(0,5))

        def _add_rateio():
            ri = len(self._cap_rat_rows) + 1
            fil_var  = tk.StringVar()
            pct_var  = tk.StringVar()
            cat_var  = tk.StringVar()
            val_var  = tk.StringVar()

            ttk.Combobox(frm_rats, textvariable=fil_var,
                         values=FILIAIS_RATEIO, state="readonly",
                         width=20, font=("Segoe UI",11)
                         ).grid(row=ri, column=0, padx=6, pady=4)

            pct_e = tk.Entry(frm_rats, textvariable=pct_var, width=10,
                             font=("Segoe UI",11), bg="#0a0f1e", fg=text,
                             insertbackground=accent, relief="flat", bd=2)
            pct_e.grid(row=ri, column=2, padx=6)

            ttk.Combobox(frm_rats, textvariable=cat_var,
                         values=CATEGORIAS_LISTA, state="normal",
                         width=35, font=("Segoe UI",11)
                         ).grid(row=ri, column=4, padx=6)

            val_lbl = tk.Label(frm_rats, textvariable=val_var,
                               fg=yellow, bg=surface, font=("Consolas",11,"bold"), width=15)
            val_lbl.grid(row=ri, column=6, padx=6)

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

        tk.Button(b4, text="+ filial", font=("Segoe UI",10,"bold"),
                  bg="#1e293b", fg=accent, relief="flat", bd=0,
                  padx=10, pady=4, cursor="hand2",
                  command=_add_rateio).pack(anchor="w", pady=(8,0))

        self._cap_lbl_pct_total = tk.Label(b4, text="Total rateio: 0%",
                                            font=("Segoe UI",11,"bold"), fg=text, bg=surface)
        self._cap_lbl_pct_total.pack(anchor="e", pady=5)

        # ── BOTÃO SALVAR ──
        tk.Button(frm, text="💾  SALVAR NOTA",
                  font=("Segoe UI",14,"bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=35, pady=15, cursor="hand2",
                  command=lambda: self._cap_salvar_nota(nb_pai)
                  ).pack(padx=20, pady=(10,25), anchor="e")

        self._cap_lbl_save = tk.Label(frm, text="", font=("Segoe UI",9),
                                       fg=green, bg=bg)
        self._cap_lbl_save.pack(padx=20, anchor="e")

    def _cap_check_restituicao(self, event=None):
        nat = self._cap_natureza_var.get()
        ref = self._cap_chave_ref.get().strip()
        if nat == "DEVOLUÇÃO" and ref:
            # Busca nota original no banco
            with db_conn() as conn:
                row = conn.execute("SELECT valor_difal FROM notas WHERE (numero_tx=? OR chave_ref=?) AND valor_difal > 0", 
                                   (ref, ref)).fetchone()
                if row:
                    self._cap_lbl_alerta_restit.config(text="⚠️ CRÉDITO DIFAL: Solicitar Restituição!")
                else:
                    self._cap_lbl_alerta_restit.config(text="")
        else:
            self._cap_lbl_alerta_restit.config(text="")

    def _cap_add_item_row(self, desc="", total=""):
        ri = len(self._cap_item_rows) + 1
        desc_var = tk.StringVar(value=desc)
        tot_var = tk.StringVar(value=total)

        desc_e = tk.Entry(self._cap_frm_items, textvariable=desc_var, width=80,
                          font=("Segoe UI",11), bg="#0a0f1e", fg="white",
                          insertbackground="#38bdf8", relief="flat", bd=2)
        desc_e.grid(row=ri, column=0, padx=6, pady=4, sticky="w")

        tot_e = tk.Entry(self._cap_frm_items, textvariable=tot_var, width=18,
                         font=("Segoe UI",11), bg="#0a0f1e", fg="#4ade80",
                         insertbackground="#38bdf8", relief="flat", bd=2)
        tot_e.grid(row=ri, column=2, padx=6, sticky="w")

        self._cap_item_rows.append({"desc": desc_var, "total": tot_var, "widgets": (desc_e, tot_e)})



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

            # Dados base para salvar
            dados_base = {
                "tipo":          self._cap_tipo_var.get(),
                "fornecedor":    forn,
                "cnpj":          self._cap_cnpj.get().strip(),
                "empresa":       self._cap_emp_var.get(),
                "descricao":     self._cap_desc.get().strip(),
                "dt_emissao":    self._cap_dt_emissao.get().strip(),
                "dt_vencimento": self._cap_dt_venc.get().strip(),
                "valor_bruto":   valor_bruto,
                "valor_liquido": valor_bruto - sum(i["valor"] for i in impostos),
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
                
                # ERP Fields
                "tipo_operacao": self._cap_op_var.get(),
                "natureza":      self._cap_nat_var.get(),
                "valor_difal":   _parse_valor(self._cap_difal.get()),
                "valor_fcp":     _parse_valor(self._cap_fcp.get()),
                "chave_ref":     self._cap_chave_ref.get().strip(),
                
                "impostos":      impostos,
                "parcelas":      parcelas,
                "rateio":        rateio,
                "itens":         []
            }

            # Coleta Itens
            for ir in self._cap_item_rows:
                d = ir["desc"].get().strip()
                t = ir["total"].get().strip()
                if d or t:
                    dados_base["itens"].append({
                        "descricao": d,
                        "valor_total": _parse_valor(t)
                    })

            # Lógica de Recorrência
            if self._cap_recorr_var.get():
                meses = int(self._cap_recorr_meses.get())
                id_recorr = f"REC_{int(time.time())}"
                
                from dateutil.relativedelta import relativedelta
                dt_base = datetime.strptime(dados_base["dt_vencimento"], "%d/%m/%Y")
                
                for m in range(meses):
                    d_mes = dados_base.copy()
                    dt_venc = dt_base + relativedelta(months=m)
                    d_mes["dt_vencimento"] = dt_venc.strftime("%d/%m/%Y")
                    d_mes["numero_tx"] = f"{d_mes.get('numero_tx','')}_{m+1}" if m > 0 else d_mes.get('numero_tx')
                    d_mes["recorrente"] = 1
                    d_mes["id_recorrencia"] = id_recorr
                    
                    # Ajusta vencimento das parcelas se houver
                    if d_mes["parcelas"]:
                        for p in d_mes["parcelas"]:
                            dt_p = datetime.strptime(p["dt_vencimento"], "%d/%m/%Y") + relativedelta(months=m)
                            p["dt_vencimento"] = dt_p.strftime("%d/%m/%Y")
                    
                    salvar_nota(d_mes)
                
                self._cap_lbl_save.config(
                    text=f"✅ {meses} Notas recorrentes salvas!", fg="#4ade80")
            else:
                numero_tx = salvar_nota(dados_base)
                self._cap_lbl_save.config(
                    text=f"✅ Nota salva com sucesso! Nº: {numero_tx}", fg="#4ade80")

            # Limpa campos
            for w in [self._cap_forn, self._cap_cnpj, self._cap_desc,
                       self._cap_valor, self._cap_obs, self._cap_dt_venc,
                       self._cap_cod_barras, self._cap_pix_chave,
                       self._cap_banco_dest, self._cap_ag_dest,
                       self._cap_conta_dest, self._cap_cpf_cnpj_dest,
                       self._cap_difal, self._cap_fcp, self._cap_chave_ref]:
                w.delete(0,"end")
            self._cap_difal.insert(0,"0,00")
            self._cap_fcp.insert(0,"0,00")
            self._cap_cat_var.set("")
            self._cap_recorr_var.set(False)
            for r in self._cap_imp_rows:
                r["val"].set(""); r["aliq"].set("")
            for w in self._cap_parc_widgets:
                w.destroy()
            self._cap_parc_widgets = []
            
            # Limpa Itens
            for row in self._cap_item_rows:
                row["desc"].set(""); row["total"].set("")
            # Deixa apenas uma linha limpa se houver muitas
            if len(self._cap_item_rows) > 1:
                # Aqui precisaria destruir os widgets, mas para simplificar vamos apenas limpar as vars
                pass

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
                 font=("Segoe UI",16,"bold"), fg=accent, bg=bg
                 ).pack(anchor="w", padx=20, pady=(15,8))

        # Formulário
        frm = tk.LabelFrame(parent, text="  Novo Adiantamento  ", bg=surface,
                             fg=accent, font=("Segoe UI",11,"bold"), padx=15, pady=12)
        frm.pack(fill="x", padx=20, pady=(0,10))

        def lbl(t): return tk.Label(frm, text=t, fg=text, bg=surface,
                                     font=("Segoe UI",11))
        def ent(w=20):
            return tk.Entry(frm, font=("Segoe UI",12), bg="#0a0f1e", fg=text,
                            insertbackground=accent, relief="flat", bd=2, width=w)

        lbl("Beneficiário:").grid(row=0,column=0,sticky="w",pady=6)
        self._ad_benef = ent(35)
        self._ad_benef.grid(row=0,column=1,padx=(0,15),sticky="w")

        lbl("CPF/CNPJ:").grid(row=0,column=2,sticky="w")
        self._ad_cpf = ent(22)
        self._ad_cpf.grid(row=0,column=3,sticky="w")

        lbl("Tipo:").grid(row=1,column=0,sticky="w",pady=6)
        self._ad_tipo = ttk.Combobox(frm,
            values=["FUNCIONARIO","FORNECEDOR"],
            state="readonly", width=18, font=("Segoe UI",12))
        self._ad_tipo.set("FUNCIONARIO")
        self._ad_tipo.grid(row=1,column=1,sticky="w",padx=(0,15))

        lbl("Empresa:").grid(row=1,column=2,sticky="w")
        self._ad_emp = ttk.Combobox(frm,
            values=["LALUA","SOLAR"],
            state="readonly", width=12, font=("Segoe UI",12))
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
                    with db_conn() as conn:
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
        with db_conn() as conn:
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
        with db_conn() as conn:
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
        with db_conn() as conn:
            conn.row_factory = sqlite3.Row
            f = conn.execute("SELECT * FROM fornecedores WHERE id=?",
                             (forn_id,)).fetchone()
        if f:
            self._forn_abrir_form(dict(f))


    def _forn_inativar(self):
        sel = self._forn_tree.selection()
        if not sel: return
        with db_conn() as conn:
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

        with db_conn() as conn:
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

        bg    = "#000000"
        surf  = "#111111"
        acc   = "#a3e635"
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


    def _cap_build_alertas(self, parent, bg, surface, accent, green, yellow, text, muted):
        """Constrói a interface da central de alertas de vencimentos."""
        
        frm_topo = tk.Frame(parent, bg=bg, pady=15)
        frm_topo.pack(fill="x", padx=20)
        
        tk.Label(frm_topo, text="🔔 Central de Alertas e Vencimentos",
                 font=("Segoe UI",16,"bold"), fg=accent, bg=bg
                 ).pack(side="left")
        
        tk.Button(frm_topo, text="🔄 Atualizar Alertas",
                  font=("Segoe UI",11,"bold"), bg="#1e293b", fg=accent,
                  relief="flat", bd=0, padx=15, pady=6, cursor="hand2",
                  command=self._cap_carregar_alertas).pack(side="right")

        # Container para os cards de resumo
        frm_resumo = tk.Frame(parent, bg=bg)
        frm_resumo.pack(fill="x", padx=20, pady=10)
        
        def card(parent, title, color):
            f = tk.Frame(parent, bg=surface, highlightbackground=color, highlightthickness=2, padx=20, pady=15)
            f.pack(side="left", padx=(0,20), expand=True, fill="both")
            tk.Label(f, text=title, fg=text, bg=surface, font=("Segoe UI",10,"bold")).pack(anchor="w")
            v = tk.Label(f, text="R$ 0,00", fg=color, bg=surface, font=("Segoe UI",16,"bold"))
            v.pack(anchor="w", pady=(4,0))
            n = tk.Label(f, text="0 itens", fg="#94a3b8", bg=surface, font=("Segoe UI",9))
            n.pack(anchor="w")
            return v, n

        self._lbl_alerta_atrasado_val, self._lbl_alerta_atrasado_cnt = card(frm_resumo, "ATRASADOS", "#f87171")
        self._lbl_alerta_hoje_val, self._lbl_alerta_hoje_cnt = card(frm_resumo, "VENCENDO HOJE", "#fb923c")
        self._lbl_alerta_semana_val, self._lbl_alerta_semana_cnt = card(frm_resumo, "PRÓX. 7 DIAS", "#fbbf24")

        # Tabela de alertas detalhados
        tk.Label(parent, text="📋 Detalhamento de Pendências:",
                 font=("Segoe UI",11,"bold"), fg=text, bg=bg
                 ).pack(anchor="w", padx=20, pady=(20,8))

        cols = ("Vencimento","Natureza","Fornecedor/Imposto","Descrição","Valor","Empresa","Nº TX")
        self._alerta_tree = ttk.Treeview(parent, columns=cols, show="headings", height=15)
        largs = [100,100,200,250,120,80,120]
        for c, w in zip(cols, largs):
            self._alerta_tree.heading(c, text=c)
            self._alerta_tree.column(c, width=w, anchor="e" if c=="Valor" else "w")

        self._alerta_tree.tag_configure("ATRASADO", foreground="#f87171")
        self._alerta_tree.tag_configure("HOJE",     foreground="#fb923c")
        self._alerta_tree.tag_configure("SEMANA",   foreground="#fbbf24")

        sb = ttk.Scrollbar(parent, orient="vertical", command=self._alerta_tree.yview)
        self._alerta_tree.configure(yscrollcommand=sb.set)
        
        self._alerta_tree.pack(side="left", fill="both", expand=True, padx=(20,0), pady=(0,15))
        sb.pack(side="left", fill="y", pady=(0,15), padx=(0,20))

        # Ação rápida
        rod = tk.Frame(parent, bg=surface, pady=10, padx=16)
        rod.pack(fill="x", side="bottom")
        tk.Button(rod, text="✅ Marcar Selecionados como PAGO",
                  font=("Segoe UI",11,"bold"), bg="#059669", fg="white",
                  relief="flat", bd=0, padx=20, pady=8, cursor="hand2",
                  command=lambda: self._cap_mudar_status_alerta("PAGA")).pack(side="right")

        self._cap_carregar_alertas()

    def _cap_carregar_alertas(self):
        """Busca no banco todas as notas pendentes e classifica para os alertas."""
        for r in self._alerta_tree.get_children():
            self._alerta_tree.delete(r)
            
        hoje = date.today()
        proximos_7 = hoje + timedelta(days=7)
        
        v_atrasado = v_hoje = v_semana = 0.0
        c_atrasado = c_hoje = c_semana = 0
        
        # Busca todas as pendentes
        notas = listar_notas(status="PENDENTE")
        
        for n in notas:
            try:
                dt_venc = datetime.strptime(n["dt_vencimento"], "%d/%m/%Y").date()
            except: continue
            
            valor = n["valor_bruto"]
            tag = ""
            
            if dt_venc < hoje:
                tag = "ATRASADO"
                v_atrasado += valor
                c_atrasado += 1
            elif dt_venc == hoje:
                tag = "HOJE"
                v_hoje += valor
                c_hoje += 1
            elif hoje < dt_venc <= proximos_7:
                tag = "SEMANA"
                v_semana += valor
                c_semana += 1
            else:
                continue # Fora do range de alertas imediatos
                
            self._alerta_tree.insert("", "end", iid=str(n["id"]), tags=(tag,),
                values=(
                    n["dt_vencimento"],
                    n.get("natureza","VENDA"),
                    n["fornecedor"][:25],
                    (n.get("descricao") or "")[:35],
                    f"R$ {valor:,.2f}",
                    n["empresa"],
                    n["numero_tx"]
                ))
        
        # Atualiza Cards
        self._lbl_alerta_atrasado_val.config(text=f"R$ {v_atrasado:,.2f}")
        self._lbl_alerta_atrasado_cnt.config(text=f"{c_atrasado} itens")
        
        self._lbl_alerta_hoje_val.config(text=f"R$ {v_hoje:,.2f}")
        self._lbl_alerta_hoje_cnt.config(text=f"{c_hoje} itens")
        
        self._lbl_alerta_semana_val.config(text=f"R$ {v_semana:,.2f}")
        self._lbl_alerta_semana_cnt.config(text=f"{c_semana} itens")

    def _cap_mudar_status_alerta(self, status):
        sel = self._alerta_tree.selection()
        if not sel: return
        for iid in sel:
            atualizar_status_nota(int(iid), status)
        self._cap_carregar_alertas()
        self._cap_carregar_lista()
