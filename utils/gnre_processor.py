import os
import re
import pypdf
from datetime import datetime
from config import PASTA_RESTITUICOES

# Estrutura de pastas da biblioteca
PASTA_BIBLIOTECA = os.path.join(PASTA_RESTITUICOES, "BIBLIOTECA_GNRE")

def split_multi_page_pdf(pdf_path, output_folder):
    """Separa um PDF de múltiplas páginas em arquivos individuais."""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    reader = pypdf.PdfReader(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    
    files_created = []
    for i, page in enumerate(reader.pages):
        writer = pypdf.PdfWriter()
        writer.add_page(page)
        
        output_filename = f"{base_name}_pag_{i+1:03d}.pdf"
        output_path = os.path.join(output_folder, output_filename)
        
        with open(output_path, "wb") as f:
            writer.write(f)
        files_created.append(output_path)
        
    return files_created

def extrair_dados_basicos(pdf_path):
    """Extrai valor e data de um comprovante Itaú."""
    try:
        reader = pypdf.PdfReader(pdf_path)
        text = reader.pages[0].extract_text()
        
        # Busca Valor (ex: R$ 544,22)
        match_valor = re.search(r"Valor.*?R\$\s*([\d\.,]+)", text)
        valor = 0.0
        if match_valor:
            val_str = match_valor.group(1).replace(".", "").replace(",", ".")
            valor = float(val_str)
            
        # Busca Data (ex: 08/04/2026)
        match_data = re.search(r"(\d{2}/\d{2}/\d{4})", text)
        data = match_data.group(1) if match_data else datetime.now().strftime("%d/%m/%Y")
        
        # Busca Autenticação
        match_auth = re.search(r"Autentica..o:\s*([A-Z0-9\s]+)", text)
        auth = match_auth.group(1).strip() if match_auth else ""
        
        return {
            "valor": valor,
            "data": data,
            "auth": auth,
            "texto": text
        }
    except Exception as e:
        print(f"Erro ao extrair dados de {pdf_path}: {e}")
        return None

def ler_relatorio_bancario(report_path):
    """Lê o relatório do banco e retorna uma lista de pagamentos com UF."""
    pagamentos = []
    try:
        reader = pypdf.PdfReader(report_path)
        for page in reader.pages:
            text = page.extract_text()
            # Padrão: Nome da SEFAZ + Valor + Status
            # Ex: SEFAZ SAO PAULO-DARE- ... 166,95 efetuado
            lines = text.split("\n")
            for line in lines:
                if "SEFAZ" in line or "GNRE" in line:
                    # Tenta extrair a UF da linha
                    uf = "GERAL"
                    if "SAO PAULO" in line: uf = "SP"
                    elif "MINAS GERAIS" in line: uf = "MG"
                    elif "DF" in line or "DISTRITO" in line: uf = "DF"
                    elif "RJ" in line or "RIO DE JANEIRO" in line: uf = "RJ"
                    
                    # Tenta extrair o valor no final da linha (antes de 'efetuado')
                    match_val = re.search(r"([\d\.,]+)\s+efetuado", line)
                    if match_val:
                        val_str = match_val.group(1).replace(".", "").replace(",", ".")
                        pagamentos.append({
                            "uf": uf,
                            "valor": float(val_str),
                            "descricao": line.strip()
                        })
        return pagamentos
    except Exception as e:
        print(f"Erro ao ler relatório: {e}")
        return []

def processar_lote_gnre(pdf_lote, report_path, log_func):
    """Fluxo completo de processamento de um lote de GNREs."""
    log_func(f"🚀 Iniciando processamento do lote: {os.path.basename(pdf_lote)}")
    
    # 1. Separar páginas
    temp_folder = os.path.join(os.path.dirname(pdf_lote), "TEMP_SPLIT")
    arquivos_individuais = split_multi_page_pdf(pdf_lote, temp_folder)
    log_func(f"📦 PDF separado em {len(arquivos_individuais)} comprovantes individuais.")
    
    # 2. Ler relatório
    base_dados_banco = ler_relatorio_bancario(report_path)
    log_func(f"📊 Relatório bancário lido: {len(base_dados_banco)} lançamentos encontrados.")
    
    sucessos = 0
    for pdf in arquivos_individuais:
        dados = extrair_dados_basicos(pdf)
        if not dados: continue
        
        # 3. Cruzamento
        # Busca no relatório por valor (com margem de erro pequena se necessário)
        match_banco = next((p for p in base_dados_banco if abs(p['valor'] - dados['valor']) < 0.01), None)
        
        uf = match_banco['uf'] if match_banco else "DESCONHECIDO"
        data_dt = datetime.strptime(dados['data'], "%d/%m/%Y")
        ano = str(data_dt.year)
        mes = f"{data_dt.month:02d}"
        
        # 4. Organização Final
        dest_folder = os.path.join(PASTA_BIBLIOTECA, uf, ano, mes)
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)
            
        nome_final = f"PAG_GNRE_{uf}_{data_dt.strftime('%Y-%m-%d')}_R${dados['valor']:.2f}.pdf"
        dest_path = os.path.join(dest_folder, nome_final)
        
        import shutil
import shutil
from utils.xml_utils import ler_xml_nfe

def extrair_dados_complexos_ai(pdf_path, prompt_extra):
    """
    Placeholder para extração via IA (Gemini Flash).
    Reservado para casos onde o layout do PDF muda ou regex falha.
    """
    # Esta função será implementada quando integrarmos a API do Gemini.
    # Por enquanto, retorna None para cair no fluxo padrão.
    return None

def renomear_gnre_via_xml(xml_dir, pdf_dir, log_func):
    """
    Cruza XMLs de NFe com PDFs de comprovantes e renomeia os PDFs (Batismo).
    Usa valor do imposto e UF para garantir o cruzamento correto.
    """
    log_func(f"🔍 Escaneando XMLs em {xml_dir}...")
    banco_notas = []
    for f in os.listdir(xml_dir):
        if f.lower().endswith(".xml"):
            res = ler_xml_nfe(os.path.join(xml_dir, f))
            # Só interessam notas que tenham algum imposto estadual (GNRE/DARE)
            if res and res["valor_imposto_estadual"] > 0:
                banco_notas.append(res)
    
    log_func(f"✅ {len(banco_notas)} notas com impostos identificadas.")
    
    renomeados = 0
    pdfs = [f for f in os.listdir(pdf_dir) if f.lower().endswith(".pdf")]
    
    for pdf_name in pdfs:
        if "NF_" in pdf_name: continue # Já batizado
        
        pdf_path = os.path.join(pdf_dir, pdf_name)
        dados_pdf = extrair_dados_basicos(pdf_path)
        if not dados_pdf:
            # Tenta extração via IA se o regex falhar (reservado)
            dados_pdf = extrair_dados_complexos_ai(pdf_path, "Extraia valor, data e autenticação.")
            if not dados_pdf: continue
        
        # Critério de match: Valor (tolerância 0.05) + UF (se disponível no PDF ou via relatório)
        # Por enquanto, focamos no valor, mas preparamos para mais critérios.
        match_nota = next((n for n in banco_notas if abs(n['valor_imposto_estadual'] - dados_pdf['valor']) < 0.05), None)
        
        if match_nota:
            # Nome padrão: PAG_GNRE_NF_[NUMERO]_[UF]_R$[VALOR].pdf
            novo_nome = f"PAG_GNRE_NF_{match_nota['nf_numero']}_{match_nota['uf_destino']}_R${dados_pdf['valor']:.2f}.pdf"
            novo_path = os.path.join(pdf_dir, novo_nome)
            
            # Evita colisões se houver notas com mesmo valor (raro, mas possível)
            counter = 1
            while os.path.exists(novo_path):
                novo_nome = f"PAG_GNRE_NF_{match_nota['nf_numero']}_{match_nota['uf_destino']}_R${dados_pdf['valor']:.2f}_{counter}.pdf"
                novo_path = os.path.join(pdf_dir, novo_nome)
                counter += 1
                
            os.rename(pdf_path, novo_path)
            log_func(f"🏷️ Batizado: {pdf_name} -> {novo_nome}")
            renomeados += 1
            
    log_func(f"✨ Batismo concluído: {renomeados} arquivos processados com sucesso.")
    return renomeados
