import re
from datetime import datetime

__all__ = ['_parse_valor', '_aplicar_mascara_valor', 'calcular_periodo_sugerido', 'auto_classificar']

def _parse_valor(texto):
    """Converte string '1.234,56' ou '1234.56' em float."""
    if not texto: return 0.0
    # Remove R$, espaços e converte formato brasileiro
    limpo = texto.replace("R$", "").replace(" ", "")
    if "," in limpo and "." in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")
    elif "," in limpo:
        limpo = limpo.replace(",", ".")
    try:
        return float(limpo)
    except:
        return 0.0

def _aplicar_mascara_valor(entry):
    """Aplica máscara de valor monetário em tempo real a um tk.Entry."""
    def formatar(event):
        if event.keysym in ("BackSpace", "Delete", "Left", "Right"): return
        texto = entry.get().replace(",", "").replace(".", "").replace("R$", "").strip()
        if not texto: return
        try:
            val = float(texto) / 100
            entry.delete(0, "end")
            entry.insert(0, f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        except: pass
    entry.bind("<KeyRelease>", formatar)

def calcular_periodo_sugerido():
    """Sugere o período de ontem ou do mês atual."""
    hoje = datetime.now()
    ontem = hoje
    # Se for segunda, sugere sexta a domingo? Não, simplificar:
    return hoje, hoje, "Hoje"

def auto_classificar(nome_fornecedor, valor=0.0, cnpj=""):
    """
    Inteligência de classificação baseada na base de dados real do ERP.
    Prioridade: 1. CNPJ exato | 2. Nome do Colaborador | 3. Nome do Fornecedor | 4. Regras Genéricas
    """
    from config import (DB_FORNECEDORES_CATEGORIA, DB_FORNECEDORES_NOME, 
                        DB_COLABORADORES, DB_FORNECEDORES_FIXOS)
    
    nome = nome_fornecedor.upper().strip()
    cnpj_clean = cnpj.strip()

    # 1. Busca por CNPJ exato na base de 266 fornecedores
    if cnpj_clean in DB_FORNECEDORES_CATEGORIA:
        r, c = DB_FORNECEDORES_CATEGORIA[cnpj_clean]
        # Regra especial: GNRE sempre GERENTE ONLINE
        if "GNRE" in c.upper():
            return "GERENTE ONLINE", c
        return r, c

    # 2. Busca por Nome do Colaborador (RH)
    if nome in DB_COLABORADORES:
        return "RH", "12201 - Salários"

    # 3. Regra especial por nome (GNRE)
    if "GNRE" in nome:
        return "GERENTE ONLINE", "22309 - GNRE"

    # 4. Busca por Nome do Fornecedor (Base DB_FORNECEDORES_NOME)
    for chave_nome, (resp, cat) in DB_FORNECEDORES_NOME.items():
        if chave_nome.upper() in nome:
            return resp, cat

    # 4. Regras Genéricas de Categoria (Fallback)
    if any(x in nome for x in ["FGTS", "GRRF"]):
        return "RH", "12205 - FGTS"
    if any(x in nome for x in ["INSS", "IRRF", "DARF", "GPS", "SIMPLES NACIONAL", "DAS", "PIS", "COFINS"]):
        return "ADM/FINANCEIRO", "22308 - Simples Nacional"
    if any(x in nome for x in ["SALARIO", "FOLHA", "PROVENTO", "FERIAS", "RESCISAO"]):
        return "RH", "12201 - Salários"
    if any(x in nome for x in ["ALUGUEL", "CONDOMINIO"]):
        return "ADM/FINANCEIRO", "21101 - Aluguel"
    if any(x in nome for x in ["ENERGIA", "COELBA", "AGUA", "EMBASA", "TELEFONE", "INTERNET"]):
        return "ADM/FINANCEIRO", "21104 - Energia Eletrica"
    if any(x in nome for x in ["MARKETING", "FACEBOOK", "GOOGLE", "ADS", "INSTAGRAM"]):
        return "MARKETING", "21701 - Comunicação/Mídia Digital"
    if any(x in nome for x in ["TRANSF", "PIX FILIAL", "MATRIZ"]):
        return "ADM/FINANCEIRO", "TRANSFERENCIA"
    if any(x in nome for x in ["FRETE", "LOGISTICA", "TRANSPORTE"]):
        return "LOGISTICA", "12112 - Frete/Transporte - Produção"

    # Fornecedores (fallback por valor)
    if valor > 500:
        return "ADM/FINANCEIRO", "12104 - Produtos Para Revenda"

    return "ADM/FINANCEIRO", "A CLASSIFICAR"
