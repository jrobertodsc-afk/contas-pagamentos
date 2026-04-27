import os
import shutil
import pypdf
from datetime import datetime
from config import PASTA_INPUT_SMART, PASTA_RESTITUICOES_XML, PASTA_BIBLIOTECA_GNRE, PASTA_RESTITUICOES_PDF
from utils.xml_utils import ler_xml_nfe
from utils.gnre_processor import extrair_dados_basicos, processar_lote_gnre, renomear_gnre_via_xml

class SmartProcessor:
    def __init__(self, log_func):
        self.log_func = log_func
        self.stats = {
            "xml_venda": 0,
            "xml_devolucao": 0,
            "pdfs_batizados": 0,
            "valor_recuperado": 0.0,
            "erros": 0
        }

    def process_all(self):
        """Varredura completa na pasta de importação automática."""
        if not os.path.exists(PASTA_INPUT_SMART):
            self.log_func("⚠️ Pasta de importação automática não encontrada.")
            return self.stats

        files = os.listdir(PASTA_INPUT_SMART)
        if not files:
            self.log_func("📭 Nada para processar na pasta de importação.")
            return self.stats

        self.log_func(f"🚀 Iniciando Processamento Inteligente de {len(files)} arquivos...")

        # 1. Identificação e Separação Inicial
        xmls = []
        pdfs = []
        relatorios = []

        for f in files:
            path = os.path.join(PASTA_INPUT_SMART, f)
            if f.lower().endswith(".xml"):
                xmls.append(path)
            elif f.lower().endswith(".pdf"):
                # Identifica se é relatório ou comprovante
                if self._is_bank_report(path):
                    relatorios.append(path)
                else:
                    pdfs.append(path)

        # 2. Processar XMLs (Mover para as pastas corretas)
        for xml_path in xmls:
            dados = ler_xml_nfe(xml_path)
            if not dados:
                self.stats["erros"] += 1
                continue

            if dados["finalidade"] == "DEVOLUCAO":
                dest = os.path.join(PASTA_RESTITUICOES_XML, os.path.basename(xml_path))
                shutil.copy(xml_path, dest) # Mantém na entrada para o batismo se necessário, ou move?
                self.stats["xml_devolucao"] += 1
                self.stats["valor_recuperado"] += dados["valor_imposto_estadual"]
            else:
                self.stats["xml_venda"] += 1
            
            # Nota: O XML de venda é usado no Batismo. 
            # No fluxo SaaS, vamos tentar batizar os PDFs que estão na pasta usando os XMLs que estão lá.

        # 3. Processar Relatórios (Se houver PDFs brutos)
        # Se houver relatórios, tentamos processar como lote
        for rel in relatorios:
            # Busca PDFs que possam ser lotes (geralmente arquivos maiores ou com nome específico)
            for pdf in pdfs:
                if os.path.getsize(pdf) > 500000: # Ex: Mais de 500KB costuma ser lote
                    processar_lote_gnre(pdf, rel, self.log_func)

        # 4. Vinculação Automática (Cruzar o que sobrou na pasta)
        renomeados = renomear_gnre_via_xml(PASTA_INPUT_SMART, PASTA_INPUT_SMART, self.log_func)
        self.stats["pdfs_batizados"] += renomeados

        # 5. Limpeza / Organização Final
        # Move PDFs vinculados para a Biblioteca se possível
        self._organizar_biblioteca_automatica()

        self.log_func("✨ Organização de arquivos concluída.")
        return self.stats

    def _is_bank_report(self, pdf_path):
        """Verifica se o PDF é um relatório bancário (Itaú, etc)."""
        try:
            reader = pypdf.PdfReader(pdf_path)
            text = reader.pages[0].extract_text()
            return "RELATÓRIO" in text.upper() or "EXTRATO" in text.upper() or "EFETUADO" in text.lower()
        except:
            return False

    def _organizar_biblioteca_automatica(self):
        """Move arquivos batizados da entrada para a biblioteca estruturada."""
        for f in os.listdir(PASTA_INPUT_SMART):
            if f.startswith("PAG_GNRE_") and f.endswith(".pdf"):
                path = os.path.join(PASTA_INPUT_SMART, f)
                dados = extrair_dados_basicos(path)
                if dados:
                    # Tenta extrair UF do nome (já que foi batizado)
                    # Formato: PAG_GNRE_NF_123_SP_R$100.00.pdf
                    parts = f.split("_")
                    uf = parts[4] if len(parts) > 4 else "GERAL"
                    
                    dt_str = dados["data"]
                    try:
                        dt = datetime.strptime(dt_str, "%d/%m/%Y")
                        ano, mes = str(dt.year), f"{dt.month:02d}"
                    except:
                        ano, mes = "9999", "00"

                    dest_dir = os.path.join(PASTA_BIBLIOTECA_GNRE, uf, ano, mes)
                    os.makedirs(dest_dir, exist_ok=True)
                    
                    shutil.move(path, os.path.join(dest_dir, f))
                    self.log_func(f"📦 Arquivado na Biblioteca: {f}")
