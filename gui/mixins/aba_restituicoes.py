import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import customtkinter as ctk

from config import *
from utils.gnre_processor import processar_lote_gnre, renomear_gnre_via_xml
from utils.xml_utils import ler_xml_nfe

from utils.smart_processor import SmartProcessor

class AbaRestituicoesMixin:
    def _build_aba_restituicoes(self, parent, bg, surface, border, accent, green, yellow, text, muted):
        # ── HEADER ──
        header_f = ctk.CTkFrame(parent, fg_color="transparent")
        header_f.pack(fill="x", padx=25, pady=(20, 10))
        
        ctk.CTkLabel(header_f, text="🏛️ MONITOR DE RESTITUIÇÕES & ROI", font=("Segoe UI", 22, "bold"), text_color=accent).pack(side="left")
        
        self.btn_sync = ctk.CTkButton(
            header_f, text="⚡ SINCRONIZAR TUDO", font=("Segoe UI", 12, "bold"), 
            fg_color=accent, text_color="#000000", hover_color=yellow,
            height=45, width=200, command=self._sincronizar_smart
        )
        self.btn_sync.pack(side="right")

        # ── INSTRUÇÕES (ROI FOCUSED) ──
        inst_f = ctk.CTkFrame(parent, fg_color=surface, corner_radius=10)
        inst_f.pack(fill="x", padx=25, pady=(0, 20))
        
        msg = f"📂 Deposite seus arquivos (XML/PDF) na pasta de entrada para análise automática."
        ctk.CTkLabel(inst_f, text=msg, font=("Segoe UI", 11), text_color=yellow, justify="left").pack(side="left", padx=20, pady=10)
        
        btns_f = ctk.CTkFrame(inst_f, fg_color="transparent")
        btns_f.pack(side="right", padx=10)

        ctk.CTkButton(
            btns_f, text="📂 ABRIR ENTRADA", font=("Segoe UI", 10, "bold"),
            fg_color="transparent", border_width=1, border_color=accent, text_color=accent,
            width=130, height=30, command=lambda: os.startfile(PASTA_INPUT_SMART)
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            btns_f, text="📄 VER RELATÓRIOS", font=("Segoe UI", 10, "bold"),
            fg_color="transparent", border_width=1, border_color=green, text_color=green,
            width=130, height=30, command=lambda: os.startfile(PASTA_DOSSIES)
        ).pack(side="left", padx=5)

        # ── DASHBOARD DE ROI (CARDS) ──
        cards_f = ctk.CTkFrame(parent, fg_color="transparent")
        cards_f.pack(fill="x", padx=25, pady=(0, 20))
        
        # Card 1: Valor Recuperado (Destaque)
        self.card_restituido = self._criar_card_roi(cards_f, "💰 RESTITUIÇÃO LOCALIZADA", "R$ 0,00", accent, "#1A1A1A", accent)
        self.card_restituido.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # Card 2: Biblioteca
        self.card_docs = self._criar_card_roi(cards_f, "📚 NOTAS VINCULADAS", "0 notas", "#FFFFFF", "#1A1A1A", border)
        self.card_docs.pack(side="left", fill="both", expand=True, padx=5)
        
        # Card 3: Eficiência
        self.card_status = self._criar_card_roi(cards_f, "🚀 STATUS OPERACIONAL", "Sistema Ocioso", yellow, "#1A1A1A", border)
        self.card_status.pack(side="left", fill="both", expand=True, padx=(10, 0))

        # ── MONITOR DE ATIVIDADE ──
        frame_work = ctk.CTkFrame(parent, fg_color="#0A0A0A", corner_radius=12, border_width=1, border_color="#1A1A1A")
        frame_work.pack(fill="both", expand=True, padx=25, pady=(0, 25))
        
        ctk.CTkLabel(frame_work, text="🖥️ LOG DE PROCESSAMENTO AUTOMÁTICO", font=("Segoe UI", 10, "bold"), text_color=muted).pack(anchor="w", padx=15, pady=(10, 5))
        
        self.log_restituicao = ctk.CTkTextbox(
            frame_work, font=("Consolas", 11), fg_color="#000000", text_color=accent, 
            border_width=0, corner_radius=10, activate_scrollbars=True
        )
        self.log_restituicao.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _criar_card_roi(self, parent, titulo, valor, color, bg, border):
        f = ctk.CTkFrame(parent, fg_color=bg, border_width=2, border_color=border, height=120)
        f.pack_propagate(False) # Mantém a altura fixa de 120
        
        ctk.CTkLabel(f, text=titulo, font=("Segoe UI", 10, "bold"), text_color="#AAAAAA").pack(pady=(20, 0))
        
        val_lbl = ctk.CTkLabel(f, text=valor, font=("Segoe UI", 26, "bold"), text_color=color)
        val_lbl.pack(pady=(5, 20))
        
        f.val_lbl = val_lbl
        return f

    def _ui_log_rest(self, mensagem):
        self.root.after(0, lambda: self.log_restituicao.insert(tk.END, mensagem + "\n"))
        self.root.after(0, lambda: self.log_restituicao.see(tk.END))

    def _sincronizar_smart(self):
        self.btn_sync.configure(state="disabled", text="⏳ PROCESSANDO...")
        self.log_restituicao.delete("1.0", tk.END)
        self._ui_log_rest("🔍 Iniciando varredura inteligente na pasta de IMPORTAÇÃO...")
        
        threading.Thread(target=self._run_smart_sync, daemon=True).start()

    def _run_smart_sync(self):
        try:
            processor = SmartProcessor(self._ui_log_rest)
            stats = processor.process_all()
            
            # Atualiza Dashboard
            self.root.after(0, lambda: self.card_restituido.val_lbl.configure(text=f"R$ {stats['valor_recuperado']:,.2f}"))
            self.root.after(0, lambda: self.card_docs.val_lbl.configure(text=f"{stats['xml_venda'] + stats['xml_devolucao'] + stats['pdfs_batizados']} docs"))
            self.root.after(0, lambda: self.card_status.val_lbl.configure(text="Sincronizado", text_color=green))
            
            # Se houver devoluções, avisa sobre os dossiês
            if stats["xml_devolucao"] > 0:
                self._ui_log_rest(f"✨ GANHO IDENTIFICADO: R$ {stats['valor_recuperado']:.2f} pronto para restituição.")
                self._run_cruzamento_automatico()
            else:
                self.root.after(0, lambda: self.card_status.val_lbl.configure(text="Sincronizado", text_color=green))

        except Exception as e:
            self._ui_log_rest(f"❌ Erro na sincronização: {e}")
        finally:
            self.root.after(0, lambda: self.btn_sync.configure(state="normal", text="⚡ SINCRONIZAR TUDO"))

    def _run_cruzamento_automatico(self):
        """Versão silenciosa do cruzamento para o fluxo SaaS."""
        self._ui_log_rest("🔨 Montando dossiês de devolução automaticamente...")
        # Reutiliza a lógica de cruzamento, mas sem popups chatos
        threading.Thread(target=self._processar_dossies_thread, daemon=True).start()

    def _processar_dossies_thread(self):
        # ... (Mantém a lógica de cruzamento que já funcionava, mas adaptada para ser silenciosa)
        try:
            pasta_xmls = PASTA_RESTITUICOES_XML
            notas_devolucao = []
            for r, _, fs in os.walk(pasta_xmls):
                for f in fs:
                    if f.lower().endswith(".xml"):
                        res = ler_xml_nfe(os.path.join(r, f))
                        if res and res["finalidade"] == "DEVOLUCAO":
                            notas_devolucao.append(res)
            
            if not notas_devolucao: return

            dossies_gerados = []
            for nf in notas_devolucao:
                valor_alvo = nf["valor_imposto_estadual"]
                if valor_alvo <= 0: continue
                chave = nf["chave_nf_original"] or ""
                
                gnre_encontrada = None
                uf_folder = os.path.join(PASTA_BIBLIOTECA_GNRE, nf['uf_destino'])
                if os.path.exists(uf_folder):
                    for r, _, fs in os.walk(uf_folder):
                        for f in fs:
                            if (chave and chave in f) or ((f"NF_{nf['nf_numero']}" in f) and (f"R${valor_alvo:.2f}" in f)):
                                gnre_encontrada = os.path.join(r, f)
                                break
                        if gnre_encontrada: break
                
                if gnre_encontrada:
                    dossies_gerados.append({"nf": nf, "gnre_pdf": gnre_encontrada})

            if dossies_gerados:
                # Gera ZIP na pasta de Dossiês Finais
                zip_name = f"Dossie_Restituicao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
                zip_path = os.path.join(PASTA_DOSSIES, zip_name)
                
                import zipfile
                with zipfile.ZipFile(zip_path, 'w') as zf:
                    for d in dossies_gerados:
                        zf.write(d["gnre_pdf"], os.path.basename(d["gnre_pdf"]))
                        zf.write(d["nf"]["caminho_arquivo"], os.path.basename(d["nf"]["caminho_arquivo"]))
                
                self._ui_log_rest(f"✅ {len(dossies_gerados)} dossiês gerados em: {zip_name}")
                self.root.after(0, lambda: self.card_status.val_lbl.configure(text="Dossiê Pronto", text_color=green))
        except Exception as e:
            self._ui_log_rest(f"⚠️ Erro ao gerar dossiês: {e}")
