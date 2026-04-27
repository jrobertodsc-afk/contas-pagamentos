"""
config.py — Configurações e constantes globais do Robô Financeiro BOAH/SOLAR.

Todos os caminhos de pasta, flags e parâmetros do sistema estão centralizados
aqui para facilitar manutenção.
"""

import os

# --- FLAGS DE BIBLIOTECAS ---
try:
    import openpyxl
    import xlrd
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    import fitz  # PyMuPDF
    PDF_OK = True
except ImportError:
    PDF_OK = False


# ==============================================================================
# ── PASTAS PRINCIPAIS ──────────────────────────────────────────────────────────
# ==============================================================================

# Raiz do projeto (pasta onde este config.py está)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Pasta de TRABALHO do robô — PDFs de entrada e logs são gerados aqui
# Equivalente ao Desktop/FINANCEIRO (pasta raiz de operação do dia a dia)
PASTA_ATUAL = r"C:\Users\Roberto\Desktop\FINANCEIRO"

# Pasta de ENTRADA — PDFs brutos colocados para o robô processar
PASTA_ENTRADA = os.path.join(PASTA_ATUAL, "ENTRADA")

# Pasta de SAÍDA ORGANIZADA — comprovantes já classificados por empresa/mês/dia
PASTA_SAIDA = os.path.join(PASTA_ATUAL, "SAIDA_ORGANIZADA")

# Pasta de EXTRATOS / FLUXO DE CAIXA (xlsx)
PASTA_EXTRATOS = os.path.join(PASTA_ATUAL, "EXTRATOS")

# Pasta de BACKUP LOCAL dos extratos
PASTA_BACKUP_EXTRATOS = os.path.join(PASTA_ATUAL, "BACKUP_EXTRATOS")

# Pasta de LOGS do robô (arquivos log_YYYYMMDD_HHMMSS.txt)
PASTA_LOGS = os.path.join(PASTA_ATUAL, "LOGS")

# Pasta de RESTITUICOES SEFAZ
PASTA_RESTITUICOES = os.path.join(PASTA_ATUAL, "RESTITUICOES")
PASTA_RESTITUICOES_XML = os.path.join(PASTA_RESTITUICOES, "XML_NOTAS")
PASTA_RESTITUICOES_PDF = os.path.join(PASTA_RESTITUICOES, "PDF_GUIAS")
PASTA_BIBLIOTECA_GNRE  = os.path.join(PASTA_RESTITUICOES, "BIBLIOTECA_GNRE")
PASTA_MODELOS_GNRE     = os.path.join(os.path.expanduser("~"), "Desktop", "MODELOS_GNRE")

# ==============================================================================
# ── SERVIDOR / BACKUP DE COMPROVANTES ─────────────────────────────────────────
# ==============================================================================

# Caminho no servidor de rede onde os PDFs de comprovantes serão espelhados
PASTA_SERVIDOR = r"Y:\01-Administrativo\02-Financeiro\04-Comprovantes"

# Pasta de BACKUP LOCAL dos comprovantes PDF (na própria máquina)
PASTA_BACKUP_PDF = os.path.join(PASTA_ATUAL, "BACKUP_COMPROVANTES")

# ==============================================================================
# ── ROTAÇÃO / LIMPEZA DE LOGS ─────────────────────────────────────────────────
# ==============================================================================

# Quantidade de dias de retenção dos arquivos de log.
# Logs mais antigos que este limite serão removidos automaticamente.
LOG_RETENCAO_DIAS = 30

# ==============================================================================
# ── BANCO DE DADOS ────────────────────────────────────────────────────────────
# ==============================================================================

DB_PATH = os.path.join(BASE_DIR, "robo_boah.db")

# ==============================================================================
# ── PARÂMETROS FINANCEIROS ────────────────────────────────────────────────────
# ==============================================================================

# Alíquotas padrão de impostos retidos (usadas no formulário de notas)
ALIQUOTAS_PADRAO = {
    "PIS":          0.65,
    "COFINS":       3.00,
    "IRRF":         1.50,
    "INSS retido":  11.00,
    "ISS":          5.00,
}

# Status válidos para notas no Contas a Pagar
STATUS_NOTA = {"PENDENTE", "APROVADA", "PAGA", "VENCIDA", "CANCELADA"}

# ==============================================================================
# ── GARANTE CRIAÇÃO DAS PASTAS LOCAIS ─────────────────────────────────────────
# ==============================================================================

# Pasta de ENTRADA INTELIGENTE — Onde o usuário joga tudo (XML, PDF, Relatórios)
PASTA_INPUT_SMART = os.path.join(PASTA_ATUAL, "IMPORTACAO_AUTOMATICA")
PASTA_DOSSIES     = os.path.join(PASTA_RESTITUICOES, "DOSSIES_FINAIS")

def ensure_directories():
    """Garante que todas as pastas de operação existam."""
    for _p in (PASTA_ATUAL, PASTA_ENTRADA, PASTA_SAIDA,
               PASTA_EXTRATOS, PASTA_BACKUP_EXTRATOS,
               PASTA_LOGS, PASTA_BACKUP_PDF,
               PASTA_RESTITUICOES, PASTA_RESTITUICOES_XML, 
               PASTA_RESTITUICOES_PDF, PASTA_BIBLIOTECA_GNRE,
               PASTA_MODELOS_GNRE, PASTA_INPUT_SMART, PASTA_DOSSIES):
        try:
            os.makedirs(_p, exist_ok=True)
        except Exception as e:
            print(f"Erro ao criar pasta {_p}: {e}")

# --- CONFIGURAÇÕES DE RH E DP ---
VALORES_DOMINGO = [33.75, 32.20, 65.95, 68.00]
CATEGORIAS_DP = ["RH - SALARIOS", "RH - DOMINGOS E FERIADOS", "RH - FGTS", "RH - RESCISAO", "RH - FERIAS", "RH - ADIANTAMENTO"]

DE_PARA_FILIAL_MAP = {
    "Filial_Barra": "Boah - Barra", "Filial_Horto": "Boah - Horto",
    "Matriz": "Boah - Matriz", "Filial_Paseo": "Boah - Paseo",
    "Filial_SDB": "Boah - SDB", "Filial_Vilas": "Boah - Vilas",
    "Empresa_Solar": "Solar - Geral", "Geral": "A Classificar"
}

DB_COLABORADORES = {
    "CINDIA SANTOS SILVA": "Filial_Barra",
    "ELLEN KAYANNE VITOR DA SILVA": "Filial_SDB",
    "GEISA SACRAMENTO MARQUES": "Filial_Barra",
    "GILMARIA DOS SANTOS RIBEIRO": "Filial_Paseo",
    "JULIA LIMA DE SOUSA MOREIRA": "Filial_Barra",
    "SOFIA DE ALMEIDA SANTOS PINH": "Filial_SDB",
    "ANA LUIZA LEMOS": "Filial_Barra", "AUGUSTO GONCALO": "Filial_Barra", "BEATRIZ FREITAS": "Filial_Barra",
    "CELIA DA COSTA": "Filial_Barra", "ISANA LAVINIA": "Filial_Barra", "JESSICA PINHEIRO": "Filial_Barra",
    "JILVANETE SANTOS": "Filial_Barra", "JOANA NATHALIA": "Filial_Barra", "JULIANA SOUZA DA CRUZ": "Filial_Barra",
    "LAIS OLIVEIRA": "Filial_Barra", "LORENA MENDONCA": "Filial_Barra", "MAIRA VITORIA": "Filial_Barra",
    "MARIA AMALIA": "Filial_Barra", "MARIA LUIZA VITA": "Filial_Barra", "MICHELE ELAINE": "Filial_Barra",
    "MILENA SILVA": "Filial_Barra", "NAIRAN SILVA": "Filial_Barra", "NATHALIA PORTO": "Filial_Barra",
    "PATRICIA DOS SANTOS": "Filial_Barra", "QUECIA VITORIA": "Filial_Barra", "SARAH FERNANDES": "Filial_Barra",
    "STEPHANIE DOS SANTOS": "Filial_Barra", "THAINA REGO": "Filial_Barra", "THAIS MILLA": "Filial_Barra",
    "WAILLYN VIEIRA": "Filial_Barra",
    "BARBARA BELA": "Filial_Horto", "CARLA PATRICIA": "Filial_Horto", "JOCASTA AVELAR": "Filial_Horto",
    "LUANA LUZIA": "Filial_Horto", "VERONICA PINTO": "Filial_Horto",
    "ALEX FERREIRA": "Matriz", "AMARILDO FERREIRA": "Matriz", "CAMILLA RODRIGUES": "Matriz",
    "DANIELLE TEIXEIRA": "Matriz", "EMILE BRANDAO": "Matriz", "JANICLEIDE OLIVEIRA": "Matriz",
    "JULIANA SOUZA SOARES": "Matriz", "LILIANE CONCEICAO": "Matriz", "LUCILEIDE NOGUEIRA": "Matriz",
    "NAIARA LANDIM": "Matriz", "NAIARA SANTOS GOMES": "Matriz", "RENILDA VIEIRA": "Matriz",
    "RUTH FERREIRA": "Matriz", "SERGIO DOS SANTOS": "Matriz",
    "ALINE CARDOSO": "Filial_Paseo", "ANA CECILIA": "Filial_Paseo", "DANIELA BONIFACIO": "Filial_Paseo",
    "DANIELA PAIXAO": "Filial_Paseo", "EDILENE GONCALVES": "Filial_Paseo", "GISELE DANTAS": "Filial_Paseo",
    "JAMILY FERREIRA": "Filial_Paseo", "LIVIA RIBEIRO": "Filial_Paseo", "MIRIAN DE SOUZA": "Filial_Paseo",
    "REGINA FABIANA": "Filial_Paseo", "RUTILENE REIS": "Filial_Paseo", "SANDRA REGINA": "Filial_Paseo",
    "SARA REGINA": "Filial_Paseo", "STERPHANNIE COSTA": "Filial_Paseo",
    "AILA MIQUELE": "Filial_SDB", "AMBRE AMINA": "Filial_SDB", "ARI DE CARVALHO": "Filial_SDB",
    "BARBARA LUCIA": "Filial_SDB", "BIANCA CATARINA": "Filial_SDB", "BRUNA MELO": "Filial_SDB",
    "BRUNA SIERPINSKA": "Filial_SDB", "ELMA GUEDES": "Filial_SDB", "GABRIELA ALBUQUERQUE": "Filial_SDB",
    "GABRIELA PEREIRA": "Filial_SDB", "JAMYLLE LEAO": "Filial_SDB", "JAQUELINE PINHEIRO": "Filial_SDB",
    "LETICIA DA SILVA": "Filial_SDB", "MARIA CLARA SALES": "Filial_SDB", "MYLENA KAROLYNA": "Filial_SDB",
    "RAFAELA SANTANA": "Filial_SDB", "REBECA DANIELE": "Filial_SDB", "SIMONE DO VALE": "Filial_SDB",
    "VIVIANE PLANZO": "Filial_SDB",
    "ALINE MESQUITA": "Filial_Vilas", "ANA PAULA SANTANA": "Filial_Vilas", "BIANCA CSEKO": "Filial_Vilas",
    "BIANCA HABIB": "Filial_Vilas", "ERICA LIMA": "Filial_Vilas", "FERNANDA BARRETO": "Filial_Vilas",
    "GABRIELA TEIXEIRA": "Filial_Vilas", "STEPHANIE ALVES": "Filial_Vilas", "STEPHANIE DANIELE": "Filial_Vilas",
    "ANA CRISTINA": "Empresa_Solar", "ANA MARIA DE SANTANA": "Empresa_Solar", "APARECIDA RIBEIRO": "Empresa_Solar",
    "ARY WANDERLEY": "Empresa_Solar", "DAISY MAIRA": "Empresa_Solar", "DANIEL MUNIZ": "Empresa_Solar",
    "EDCARLA DOS SANTOS": "Empresa_Solar", "ELEN MENDES": "Empresa_Solar", "FINELON GOMES": "Empresa_Solar",
    "JOAO ESTEVAM": "Empresa_Solar", "JOSE ROBERTO": "Empresa_Solar", "LILIA DE JESUS": "Empresa_Solar",
    "LORENA DANTAS": "Empresa_Solar", "MAGNO SENA": "Empresa_Solar", "RAIETE CAROL": "Empresa_Solar",
    "RUTE MARIA": "Empresa_Solar", "SOFIA CORREIA": "Empresa_Solar"
}

DB_FORNECEDORES_FIXOS = {
    "ALAMEDA VERT":       "Filial_Horto",
    "CONDOMINIO ALAMEDA": "Filial_Horto",
    "CONDOMINIO PASEO":   "Filial_Paseo",
    "CONDOMINIO SHOPPING BARRA": "Filial_Barra",
    "SHOPPING DA BAHIA":  "Filial_SDB",
    "COELBA":             "Matriz",
    "EMBASA":             "Matriz",
    "TIM CELULAR":        "Matriz",
    "CLARO":              "Matriz",
}

MESES = {
    1:"01_Janeiro", 2:"02_Fevereiro", 3:"03_Marco", 4:"04_Abril",
    5:"05_Maio", 6:"06_Junho", 7:"07_Julho", 8:"08_Agosto",
    9:"09_Setembro", 10:"10_Outubro", 11:"11_Novembro", 12:"12_Dezembro"
}

DB_FORNECEDORES_CATEGORIA = {
    "38.462.298/0001-24": ("ADM/FINANCEIRO", "12202 - Salários - Meis/PJ"),
    "03.165.536/0001-55": ("COMPRAS", "21707 - Material Gráfico"),
    "36.193.540/0001-86": ("MARKETING", "21703 - Lookbook"),
    "41.258.802/0001-83": ("ADM/FINANCEIRO", "12200 - Salários"),
    "50.942.841/0001-96": ("MARKETING", "21703 - Lookbook"),
    "51.313.763/0001-23": ("MARKETING", "21703 - Lookbook"),
    "52.244.938/0001-50": ("COMPRAS", "12102 - Aviamentos"),
    "60.662.136/0001-99": ("ADM/FINANCEIRO", "12202 - Salários - Meis/PJ"),
    "62.336.608/0001-49": ("ADM/FINANCEIRO", "12202 - Salários"),
    "64.082.848/0001-90": ("GERENTE LOJA", "21707 - Material Gráfico"),
    "23.827.668/0001-02": ("ADM/FINANCEIRO", "21102 - Condominio"),
    "33.583.036/0001-02": ("ADM/FINANCEIRO", "21116 - Sindicato E Associcoes"),
    "01.397.927/0001-70": ("ADM/FINANCEIRO", "21101 - Aluguel"),
    "61.009.520/0001-50": ("SUPRIMENTOS", "22206 - Embalagens (Sacolas, Envelopes e Papel Seda)"),
    "44.926.364/0001-72": ("ADM/FINANCEIRO", "21101 - Aluguel"),
    "61.573.796/0001-66": ("ADM/FINANCEIRO", "21109 - Seguros Loja/Imóvel"),
    "32.346.596/0001-72": ("ADM/FINANCEIRO", "12200 - Salários"),
    "19.161.341/0001-77": ("ADM/FINANCEIRO", "21116 - Sindicato E Associcoes"),
    "05.480.257/0001-01": ("COMPRAS", "12102 - Aviamentos"),
    "48.740.351/0001-65": ("ADM/FINANCEIRO", "12112 - Frete/Transporte - Produção"),
    "00.360.305/0001-04": ("ADM/FINANCEIRO", "12207 - Rescisão / FGTS"),
    "29.221.576/0001-60": ("ADM/FINANCEIRO", "12203 - Transportes"),
    "47.201.850/0001-11": ("COMPRAS", "12102 - Aviamentos"),
    "15.139.629/0001-94": ("ADM/FINANCEIRO", "21104 - Energia Eletrica"),
    "08.883.683/0001-84": ("COMPRAS", "12101 - Tecidos"),
    "16.188.955/0001-54": ("ADM/FINANCEIRO", "21101 - Aluguel"),
    "10.420.899/0001-55": ("ADM/FINANCEIRO", "21102 - Condominio"),
    "54.259.756/0001-89": ("ADM/FINANCEIRO", "21102 - Condominio"),
    "31.872.549/0001-08": ("ADM/FINANCEIRO", "12203 - Transportes"),
    "32.726.717/0001-01": ("ADM/FINANCEIRO", "21101 - Aluguel"),
    "22.267.284/0001-10": ("ADM/FINANCEIRO", "12203 - Transportes"),
    "83.310.441/0001-17": ("ADM/FINANCEIRO", "21116 - Sindicato E Associcoes"),
    "60.507.500/0001-46": ("MARKETING", "21705 - Relac com o Cliente - Ações Clientes"),
    "00.955.829/0001-48": ("SUPRIMENTOS", "12106 - Sacolas"),
    "47.738.604/0001-01": ("COMPRAS", "12102 - Aviamentos"),
    "14.406.482/0001-99": ("ADM/FINANCEIRO", "21112 - Servicos Contabeis"),
    "22.014.850/0001-81": ("COMPRAS", "12105 - Produtos Para Revenda"),
    "34.028.316/0001-03": ("ADM/FINANCEIRO", "22201 - Entregas On-line"),
    "27.080.571/0001-30": ("GERENTE ONLINE", "22309 - GNRE"),
    "47.379.714/0001-16": ("COMPRAS", "12101 - Tecidos"),
    "02.384.871/0001-81": ("COMPRAS", "12101 - Tecidos"),
    "14.796.606/0001-90": ("PLANEJAMENTO", "21701 - Desp. Operacionais/Facebook/Email"),
    "10.465.955/0001-78": ("COMPRAS", "12112 - Frete/Transporte - Produção"),
    "40.883.098/0001-97": ("ADM/FINANCEIRO", "12202 - Salários - Meis/PJ"),
    "33.815.420/0001-85": ("MARKETING", "21703 - Lookbook"),
    "11.498.408/0001-51": ("PLANEJAMENTO", "21701 - Desp. Operacionais/Facebook/Email"),
    "13.727.240/0001-34": ("COMPRAS", "12112 - Frete/Transporte - Produção"),
    "15.243.447/0001-69": ("MARKETING", "21703 - Lookbook"),
    "43.182.349/0001-02": ("MARKETING", "21801 - Endomarketing"),
    "33.988.601/0001-03": ("COMPRAS", "12101 - Tecidos"),
    "00.930.402/0001-95": ("COMPRAS", "12101 - Tecidos"),
    "43.284.638/0001-04": ("MARKETING", "21705 - Relac com o Cliente - Ações Clientes"),
    "06.990.590/0001-23": ("PLANEJAMENTO", "21701 - Desp. Operacionais/Facebook/Email"),
    "82.645.862/0001-36": ("ADM/FINANCEIRO", "12107 - Etiquetas roupas/Acessorios"),
    "27.914.259/0001-02": ("SUPRIMENTOS", "21206 - Material de Limpeza"),
    "22.761.991/0001-68": ("ADM/FINANCEIRO", "21114 - Telefonia Fixa/Internet"),
    "45.251.219/0001-00": ("MARKETING", "21703 - Lookbook"),
    "10.820.791/0001-50": ("COMPRAS", "12101 - Tecidos"),
    "25.912.503/0001-64": ("ADM/FINANCEIRO", "21120 - Sistemas e Softwares"),
    "37.134.249/0001-08": ("MARKETING", "21702 - Editorial para Campanha"),
    "08.772.214/0001-98": ("ADM/FINANCEIRO", "21114 - Telefonia Fixa/Internet"),
    "26.937.210/0001-02": ("LOGISTICA", "21304 - Manutenção - Moto/Carro"),
    "55.952.885/0001-10": ("ADM/FINANCEIRO", "21801 - Endomarketing"),
    "48.542.702/0001-23": ("SUPRIMENTOS", "21706 - Visual Merchandising"),
    "10.436.619/0003-69": ("ADM/FINANCEIRO", "TRANSFERENCIA"),
    "38.265.141/0001-09": ("ADM/FINANCEIRO", "21101 - Aluguel"),
    "16.906.199/0001-51": ("COMPRAS", "12101 - Tecidos"),
    "09.411.448/0001-72": ("COMPRAS", "12112 - Frete/Transporte - Produção"),
    "25.694.899/0001-10": ("SUPRIMENTOS", "21209 - Uso e Consumo Lojas (Copos)"),
    "05.140.762/0001-07": ("COMPRAS", "12101 - Tecidos"),
    "48.204.195/0001-18": ("ADM/FINANCEIRO", "12200 - Salários"),
    "61.721.785/0001-86": ("MARKETING", "21703 - Lookbook"),
    "10.474.553/0001-30": ("COMPRAS", "12101 - Tecidos"),
    "73.186.116/0001-30": ("GERENTE ONLINE", "21120 - Sistemas e Softwares"),
    "07.225.209/0001-00": ("ADM/FINANCEIRO", "12101 - Tecidos"),
    "39.664.088/0001-81": ("COMPRAS", "12101 - Tecidos"),
    "35.654.917/0001-94": ("COMPRAS", "12101 - Tecidos"),
    "15.089.323/0001-70": ("COMPRAS", "22201 - Entregas On-line"),
    "29.469.420/0001-01": ("COMPRAS", "12101 - Tecidos"),
    "42.154.687/0001-60": ("ADM/FINANCEIRO", "12101 - Tecidos"),
    "14.955.141/0001-72": ("COMPRAS", "12101 - Tecidos"),
    "12.202.612/0001-46": ("COMPRAS", "12101 - Tecidos"),
    "35.616.849/0001-79": ("ADM/FINANCEIRO", "21707 - Material Gráfico"),
    "18.191.228/0001-71": ("GERENTE ONLINE", "21120 - Sistemas e Softwares"),
    "00.822.602/0001-24": ("COMPRAS", "12107 - Etiquetas roupas/Acessorios"),
    "05.278.412/0001-01": ("ADM/FINANCEIRO", "22410 - Investimento - Reformas"),
    "55.343.764/0001-71": ("GERENTE LOJA", "21801 - Endomarketing"),
    "27.657.581/0001-95": ("COMPRAS", "12104 - Produtos Para Revenda"),
    "37.967.052/0001-41": ("PLANEJAMENTO", "21701 - Desp. Operacionais/Facebook/Email"),
    "23.414.602/0001-90": ("ADM/FINANCEIRO", "21113 - Consultorias e Auditorias"),
    "47.292.795/0001-12": ("ADM/FINANCEIRO", "12209 - Exames Clínicos (Dem/Adm/Per)"),
    "14.014.761/0001-07": ("COMPRAS", "12101 - Tecidos"),
    "80.446.990/0001-25": ("COMPRAS", "12102 - Aviamentos"),
    "90.400.888/0001-42": ("ADM/FINANCEIRO", "21120 - Sistemas e Softwares"),
    "15.239.478/0001-46": ("ADM/FINANCEIRO", "21117 - Sindicato E Associcoes"),
    "32.700.213/0001-12": ("ADM/FINANCEIRO", "21116 - Sindicato E Associcoes"),
    "04.316.357/0001-34": ("COMPRAS", "12104 - Produtos Para Revenda"),
    "26.607.445/0001-28": ("ADM/FINANCEIRO", "TRANSFERENCIA"),
    "30.777.576/0001-20": ("COMPRAS", "21707 - Material Gráfico"),
    "21.933.166/0001-30": ("PLANEJAMENTO", "21120 - Sistemas e Softwares"),
    "48.618.124/0001-61": ("ADM/FINANCEIRO", "21113 - Consultorias e Auditorias"),
    "38.764.395/0001-71": ("SUPRIMENTOS", "12107 - Etiquetas roupas/Acessorios"),
    "73.939.449/0001-93": ("GERENTE ONLINE", "22201 - Entregas On-line"),
    "43.248.764/0001-03": ("COMPRAS", "12101 - Tecidos"),
    "05.003.162/0001-05": ("COMPRAS", "12101 - Tecidos"),
    "17.490.828/0001-78": ("COMPRAS", "12102 - Aviamentos"),
    "47.866.934/0001-74": ("ADM/FINANCEIRO", "12204 - Alimentação"),
    "03.506.307/0001-57": ("LOGISTICA", "21301 - Combustíveis Motoboy"),
    "02.421.421/0001-11": ("ADM/FINANCEIRO", "21114 - Telefonia Fixa/Internet"),
    "11.247.476/0001-48": ("COMPRAS", "12101 - Tecidos"),
    "22.972.620/0001-25": ("ADM/FINANCEIRO", "21501 - Manutenção - Predial"),
    "29.709.951/0001-16": ("COMPRAS", "21202 - Dedetização"),
    "12.803.527/0001-33": ("GERENTE ONLINE", "21120 - Sistemas e Softwares"),
    "02.535.864/0007-29": ("ADM/FINANCEIRO", "22201 - Entregas On-line"),
    "46.407.843/0001-08": ("ADM/FINANCEIRO", "12202 - Salários - Meis/PJ"),
    "23.904.297/0001-15": ("PRODUCAO", "12111 - Faccionista - Mão de Obra"),
    "34.350.031/0001-94": ("ADM/FINANCEIRO", "12202 - Salários - Meis/PJ"),
    "07.014.198/0003-73": ("SUPRIMENTOS", "21204 - Material de Escritório"),
    "85.236.743/0001-18": ("ADM/FINANCEIRO", "21120 - Sistemas e Softwares"),
    "27.004.715/0001-79": ("ADM/FINANCEIRO", "12203 - Salários - Meis/PJ"),
    "34.600.339/0001-03": ("COMPRAS", "12104 - Produtos Para Revenda"),
    "34.600.339/0001-40": ("COMPRAS", "12104 - Produtos Para Revenda"),
    "54.517.628/0001-98": ("ADM/FINANCEIRO", "21120 - Sistemas e Softwares"),
    "05.231.614/0001-06": ("COMPRAS", "12101 - Tecidos"),
    "46.390.771/0001-33": ("PRODUCAO", "12111 - Faccionista - Mão de Obra"),
    "61.083.952/0001-00": ("COMPRAS", "12102 - Aviamentos"),
    "***.666.465-**": ("MARKETING",  "21703 - Lookbook"),           # PAULA CALINE DA SILVA OLIVEIRA
    "***.208.315-**": ("PRODUCAO",   "12111 - Faccionista - Mão de Obra"),  # ANA CAROLINE MAGALHAES RAMOS
    "***.144.215-**": ("PRODUCAO",   "12111 - Faccionista - Mão de Obra"),  # JUVANICE DA CONCEICAO SOUSA
    "***.511.455-**": ("PRODUCAO",   "12111 - Faccionista - Mão de Obra"),  # LEIDE FERREIRA DOS SANTOS
    "***.850.015-**": ("PRODUCAO",   "12111 - Faccionista - Mão de Obra"),  # REGINA MARIA DIAS DA SILVA
    "***.381.965-**": ("MARKETING",  "21703 - Lookbook"),           # JULIA GARCEZ DE SENA SARMENTO
}

DB_FORNECEDORES_NOME = {
    "LINX SIST": ("ADM/FINANCEIRO", "21120 - Sistemas e Softwares"),
    "CONSORCIO NACIGUAT": ("ADM/FINANCEIRO", "21101 - Aluguel"),
    "TEX COURIER": ("GERENTE ONLINE", "22201 - Entregas On-line"),
    "MULTIPLIKE": ("COMPRAS", "12101 - Tecidos"),
    "TEXTIL FAVERO": ("COMPRAS", "12101 - Tecidos"),
    "MENEGOTTI TEXTIL": ("COMPRAS", "12101 - Tecidos"),
    "CRZ COM E REPRES": ("COMPRAS", "12102 - Aviamentos"),
    "GOOGLE BRASIL": ("PLANEJAMENTO", "21701 - Desp. Operacionais/Facebook/Email"),
    "TICKET SERVICOS": ("ADM/FINANCEIRO", "12204 - Alimentação"),
    "TICKET SOLUCOES": ("LOGISTICA", "21301 - Combustíveis Motoboy"),
    "HACO ETIQUETAS": ("ADM/FINANCEIRO", "12107 - Etiquetas roupas/Acessorios"),
    "EXCIM IMP E EXP": ("COMPRAS", "12101 - Tecidos"),
    "MICROPOST": ("GERENTE ONLINE", "21120 - Sistemas e Softwares"),
    "ECOBOTOES": ("COMPRAS", "12105 - Produtos Para Revenda"),
    "ALAMEDA VERT": ("ADM/FINANCEIRO", "21101 - Aluguel"),
    "COMP ELE EST DA BA": ("ADM/FINANCEIRO", "21104 - Energia Eletrica"),
    "COMPANHIA DE ELETRICIDADE": ("ADM/FINANCEIRO", "21104 - Energia Eletrica"),
    "COELB": ("ADM/FINANCEIRO", "21104 - Energia Eletrica"),
    "EMBASA": ("ADM/FINANCEIRO", "21103 - Agua E Esgotos"),
    "SEND4": ("GERENTE ONLINE", "21120 - Sistemas e Softwares"),
    "FORMATO TRANSPORTES": ("COMPRAS", "12112 - Frete/Transporte - Produção"),
    "O S S CREDITOS": ("COMPRAS", "12101 - Tecidos"),
    "THALIA AVIAMENTOS": ("COMPRAS", "12102 - Aviamentos"),
    "ITS TELECOMUNICACOES": ("ADM/FINANCEIRO", "21114 - Telefonia Fixa/Internet"),
    "HUP TELECOM": ("ADM/FINANCEIRO", "21114 - Telefonia Fixa/Internet"),
    "TIM S/A": ("ADM/FINANCEIRO", "21114 - Telefonia Fixa/Internet"),
    "TIM CELULAR": ("ADM/FINANCEIRO", "21114 - Telefonia Fixa/Internet"),
    "CLARO S.A": ("ADM/FINANCEIRO", "21114 - Telefonia Fixa/Internet"),
    "CLARO": ("ADM/FINANCEIRO", "21114 - Telefonia Fixa/Internet"),
    "SEFAZ BAHIA": ("ADM/FINANCEIRO", "22301 - ICMS (Parc paseo)"),
    "SEFAZ": ("ADM/FINANCEIRO", "22309 - GNRE"),
    "GNRE": ("ADM/FINANCEIRO", "22309 - GNRE"),
    "RECEITA FED": ("ADM/FINANCEIRO", "12205 - IRRF - Imposto de Renda"),
    "DAS-SIMPLES": ("ADM/FINANCEIRO", "22308 - Simples Nacional"),
    "SIMPLES NACIONAL": ("ADM/FINANCEIRO", "22308 - Simples Nacional"),
    "DARF": ("ADM/FINANCEIRO", "12205 - IRRF - Imposto de Renda"),
    "FLOWBIZ": ("PLANEJAMENTO", "21701 - Desp. Operacionais/Facebook/Email"),
    "FACEBOOK": ("PLANEJAMENTO", "21701 - Desp. Operacionais/Facebook/Email"),
    "RG MARK": ("PLANEJAMENTO", "21701 - Desp. Operacionais/Facebook/Email"),
    "SINDICOF": ("ADM/FINANCEIRO", "21116 - Sindicato E Associcoes"),
    "SANCRIS LINHAS": ("COMPRAS", "12102 - Aviamentos"),
    "MS OPEN FUNDO": ("COMPRAS", "12101 - Tecidos"),
    "MS MULTI FUNDO": ("COMPRAS", "12101 - Tecidos"),
    "ASSOCIACAO DOS LOJISTAS": ("ADM/FINANCEIRO", "21116 - Sindicato E Associcoes"),
    "SINDICATO DOS EMPREGADOS": ("ADM/FINANCEIRO", "21117 - Sindicato E Associcoes"),
    "ABIV": ("ADM/FINANCEIRO", "21116 - Sindicato E Associcoes"),
    "SOLAR SERV": ("ADM/FINANCEIRO", "TRANSFERENCIA"),
    "SOLAR SERVIÇOS": ("ADM/FINANCEIRO", "TRANSFERENCIA"),
    "IGLU LOC": ("MARKETING", "21703 - Lookbook"),
    "ALLIANZ SEGUROS": ("ADM/FINANCEIRO", "21109 - Seguros Loja/Imóvel"),
    "BRASIL BOTOES": ("COMPRAS", "12102 - Aviamentos"),
    "CENTRAL VT SERVICOS": ("ADM/FINANCEIRO", "12203 - Transportes"),
    "CONSORCIO SALVADOR TRANSCARD": ("ADM/FINANCEIRO", "12203 - Transportes"),
    "CONSORCIO METROPASSE": ("ADM/FINANCEIRO", "12203 - Transportes"),
    "DMF SERVICOS": ("ADM/FINANCEIRO", "21112 - Servicos Contabeis"),
    "SHOPPING ALAMEDA VERT": ("ADM/FINANCEIRO", "21102 - Condominio"),
    "CONDOMINIO SHOPPING": ("ADM/FINANCEIRO", "21102 - Condominio"),
    "CONDOMINIO PASEO": ("ADM/FINANCEIRO", "21102 - Condominio"),
    "M LEITAO": ("COMPRAS", "12101 - Tecidos"),
    "CONCEPT TXTL": ("COMPRAS", "12101 - Tecidos"),
    "CAIXA ECONOMICA FEDERAL": ("ADM/FINANCEIRO", "12207 - Rescisão / FGTS"),
    "CEF MATRIZ": ("ADM/FINANCEIRO", "RH - FGTS"),
    "SPOT METRICS": ("PLANEJAMENTO", "21120 - Sistemas e Softwares"),
    "SOCIALCRED": ("COMPRAS", "12104 - Produtos Para Revenda"),
    "COMERCIAL DE ZIPERES": ("COMPRAS", "12102 - Aviamentos"),
    "MT LOG SOLUCOES": ("COMPRAS", "22201 - Entregas On-line"),
    "TEXTIL SUICA": ("COMPRAS", "12101 - Tecidos"),
    "LAPA PATRIMONIAL": ("ADM/FINANCEIRO", "21101 - Aluguel"),
    "ABR PLANEJAMENTO": ("ADM/FINANCEIRO", "21101 - Aluguel"),
    "INNOVATIV INDUSTRIA": ("COMPRAS", "12101 - Tecidos"),
    "GALVANOTECNICA": ("COMPRAS", "12101 - Tecidos"),
    "GIMENEZ JACOB": ("COMPRAS", "12101 - Tecidos"),
    "V.L.O TEXTIL": ("COMPRAS", "12101 - Tecidos"),
    "VNDA SERVICOS": ("GERENTE ONLINE", "21120 - Sistemas e Softwares"),
    "VNDA SERVIÇOS": ("GERENTE ONLINE", "21120 - Sistemas e Softwares"),
    "VR BENEFICIOS": ("ADM/FINANCEIRO", "22201 - Entregas On-line"),
    "RIOMED": ("ADM/FINANCEIRO", "12209 - Exames Clínicos (Dem/Adm/Per)"),
    "SST MAIS": ("ADM/FINANCEIRO", "21113 - Consultorias e Auditorias"),
    "BEM MAIS GESTORA": ("ADM/FINANCEIRO", "21116 - Sindicato E Associcoes"),
    "INTEGRADO SISTEMAS": ("ADM/FINANCEIRO", "21120 - Sistemas e Softwares"),
    "SANTANDER NEGOCIOS": ("ADM/FINANCEIRO", "21120 - Sistemas e Softwares"),
    "EMPRESA BRASILEIRA DE CORREIOS": ("ADM/FINANCEIRO", "22201 - Entregas On-line"),
    "BRASPRESS": ("ADM/FINANCEIRO", "12112 - Frete/Transporte - Produção"),
    "LUCAS FIGUEIREDO": ("SUPRIMENTOS", "21209 - Uso e Consumo Lojas (Copos)"),
    "VALTER VANDERLEI": ("ADM/FINANCEIRO", "21501 - Manutenção - Predial"),
    "CORK COMERCIO": ("MARKETING", "21705 - Relac com o Cliente - Ações Clientes"),
    "MARGO CONFEITARIA": ("MARKETING", "21705 - Relac com o Cliente - Ações Clientes"),
    "JIREH BRINDES": ("ADM/FINANCEIRO", "21801 - Endomarketing"),
    "GABRIELA REGIS": ("MARKETING", "21801 - Endomarketing"),
    "PJBANK": ("GERENTE ONLINE", "21120 - Sistemas e Softwares"),
    "AUDACES": ("ADM/FINANCEIRO", "21120 - Sistemas e Softwares"),
    "PREF MUN SALVADOR": ("ADM/FINANCEIRO", "21902 - Taxas Municipais"),
    "IPTU": ("ADM/FINANCEIRO", "21901 - IPTU"),
    "JAC SERVICOS": ("LOGISTICA", "21304 - Manutenção - Moto/Carro"),
    "L3E COMERCIO": ("SUPRIMENTOS", "21706 - Visual Merchandising"),
    "ATACADAO DO PAPEL": ("SUPRIMENTOS", "21204 - Material de Escritório"),
    "CROMAGRAFYC": ("SUPRIMENTOS", "12106 - Sacolas"),
    "HIGIFORTE": ("SUPRIMENTOS", "21206 - Material de Limpeza"),
    "SUPRILINX": ("SUPRIMENTOS", "12107 - Etiquetas roupas/Acessorios"),
    "AG ANTECIPA": ("SUPRIMENTOS", "22206 - Embalagens (Sacolas, Envelopes e Papel Seda)"),
    "LOLA FEITO": ("COMPRAS", "12104 - Produtos Para Revenda"),
    "LITORAL COMERCIO EXTERIOR": ("COMPRAS", "12101 - Tecidos"),
    "ROYAL BLUE": ("COMPRAS", "12101 - Tecidos"),
    "LDB LOGISTICA": ("COMPRAS", "12101 - Tecidos"),
    "LDB TRANSPORTES": ("COMPRAS", "12112 - Frete/Transporte - Produção"),
    "ESTAMPARIA SALETE": ("COMPRAS", "12101 - Tecidos"),
    "RELIC OPTICAL": ("COMPRAS", "12104 - Produtos Para Revenda"),
    "PLOTAG PAPEIS": ("COMPRAS", "12107 - Etiquetas roupas/Acessorios"),
    "SOUL PRINT": ("COMPRAS", "21707 - Material Gráfico"),
    "3D SERVICOS GRAFICOS": ("COMPRAS", "21707 - Material Gráfico"),
    "PIRES PINHEIRO": ("ADM/FINANCEIRO", "21707 - Material Gráfico"),
    "VICTORIA DE CASTRO": ("COMPRAS", "21202 - Dedetização"),
    "PUBLISILK": ("ADM/FINANCEIRO", "22410 - Investimento - Reformas"),
    "40 GRAUS BA": ("MARKETING", "21703 - Lookbook"),
    "MCL PUBLICIDADE": ("MARKETING", "21703 - Lookbook"),
    "IONARA ARGOLO": ("MARKETING", "21702 - Editorial para Campanha"),
    "FUNDACAO MUSEU": ("MARKETING", "21703 - Lookbook"),
    "QZ COMERCIO": ("GERENTE LOJA", "21801 - Endomarketing"),
    "MILE MONEY": ("COMPRAS", "12101 - Tecidos"),
    "KAIQUE RAMOS": ("ADM/FINANCEIRO", "21505 - Manutenção - Refrigeração/Ar-condicionado"),
    "NAESON GOMES": ("ADM/FINANCEIRO", "21502 - Manutenção - Elétrica"),
    "EDUARDO NASCIMENTO": ("ADM/FINANCEIRO", "21501 - Manutenção - Predial"),
    "CASA FERRO": ("COMPRAS", "12102 - Aviamentos"),
    "FELICITY INDUSTRIA": ("COMPRAS", "12112 - Frete/Transporte - Produção"),
}

CATEGORIAS_LISTA = sorted(set([
    "11101 - Acervo Stilling",
    "11102 - Estilista Terçeiros (Salário)",
    "11103 - Peças de Inspiração/Pilotos",
    "11104 - Peças Piloto de Terceiros",
    "11105 - Custos Com Viagens",
    "11201 - Produção - Cadista/Corte/Expedição/Modelagem/Pilotagem",
    "12101 - Tecidos", "12102 - Aviamentos", "12103 - Insumos Gerais (produção)",
    "12104 - Produtos Para Revenda", "12105 - Embalagens", "12106 - Sacolas",
    "12107 - Etiquetas roupas/Acessorios", "12108 - Lacres",
    "12109 - Prestação de Serviço - Corte", "12110 - Faccionista - Conserto",
    "12111 - Faccionista - Mão de Obra", "12112 - Frete/Transporte - Produção",
    "12113 - Faccionista - Confecção de Pilotos",
    "12201 - Salários", "12202 - Salários - Meis/PJ", "12203 - Prolabore",
    "12203 - Transportes", "12204 - Alimentação", "12205 - FGTS", "12205 - INSS",
    "12205 - IRRF - Imposto de Renda PF", "12206 - Férias", "12207 - Rescisão",
    "12207 - Multa de FGTS", "12208 - Domingos e Feriados Trabalhados",
    "12208 - Cursos e Treinamentos", "12209 - Exames Clínicos (Dem/Adm/Per)",
    "12210 - Despesas com Estágio", "12211 - 13º Salário", "12212 - Sindicatos",
    "12213 - Extras", "12214 - Fardamento", "12215 - ISS Substituto Tributário",
    "21101 - Aluguel", "21102 - Condominio", "21103 - Agua E Esgotos",
    "21104 - Energia Eletrica", "21105 - Ar Condicionado",
    "21106 - Fundo De Promoção/Reserva", "21107 - Reembolso de Despesas Operacionais",
    "21109 - Seguros Loja/Imóvel", "21110 - Gráficas Em Geral",
    "21111 - Serviços Advocatícios", "21112 - Servicos Contabeis",
    "21113 - Consultorias e Auditorias", "21114 - Telefonia Fixa/Internet",
    "21115 - Telefonia Movel", "21116 - Sindicato E Associcoes", "21117 - Seguro Geral",
    "21118 - Correios, Cartorios E Periodicos", "21119 - Consulta SPC/Serasa",
    "21120 - Sistemas e Softwares", "21121 - Dominios, Emails E Site",
    "21201 - Copa e Cozinha", "21202 - Dedetização", "21203 - Recarga de Extintores",
    "21204 - Material de Escritório", "21205 - Material de Informática",
    "21206 - Material de Limpeza", "21207 - Despesas com PET",
    "21208 - Prestações de Serviços Operacionais", "21209 - Uso e Consumo Lojas",
    "21301 - Combustíveis Motoboy", "21302 - Estacionamento/Pedágio",
    "21303 - Licenciamento e Multas/IPVA Moto/Carro", "21304 - Manutenção - Moto/Carro",
    "21305 - Seguros - Moto/Carro", "21306 - Táxi/Uber", "21501 - Manutenção - Predial",
    "21502 - Manutenção - Elétrica", "21503 - Manutenção - Informática",
    "21504 - Manutenção - Maquinas e Equipamentos",
    "21505 - Manutenção - Refrigeração/Ar-condicionado",
    "21506 - Manutenção - Mobiliário/Decoração", "21601 - Tarifas Bancárias",
    "21602 - Aluguel de Maquinetas", "21603 - Taxa de Adm Cartões/Pix/Boletos",
    "21701 - Comunicação/Mídia Digital", "21702 - Editorial para Campanha",
    "21703 - Lookbook", "21704 - Marketing de Influência - Influencers",
    "21705 - Relacionamento com o Cliente", "21706 - Visual Merchandising",
    "21707 - Material Gráfico", "21801 - Endomarketing", "21901 - IPTU",
    "21902 - Taxas Municipais", "21903 - Taxas Estaduais",
    "22101 - Bonificação / Premiação", "22102 - Comissões", "22201 - Entregas On-line",
    "22202 - Entregas Atacado", "22203 - Plataforma de Vendas On-line",
    "22204 - Devolução de Vendas", "22205 - Aluguel Percentual",
    "22206 - Embalagens (Sacolas, Envelopes e Papel Seda)", "22207 - Ações Comerciais",
    "22301 - ICMS", "22302 - ICMS Substituição Tributária", "22303 - ICMS Antecipação Parcial",
    "22304 - PIS (8109)", "22305 - COFINS (2172)", "22306 - IRPJ (2089)",
    "22307 - CSLL (2372)", "22308 - Simples Nacional", "22309 - GNRE",
    "22309 - GNRE - Alagoas", "22309 - GNRE - Santa Catarina",
    "22309 - GNRE - Minas Gerais", "22309 - GNRE - DF", "22309 - GNRE - Paraná",
    "22309 - GNRE - Rio Grande do Norte", "22309 - GNRE - Pernambuco",
    "22309 - GNRE - Rio de Janeiro", "22309 - GNRE - Ceará", "22310 - ISS",
    "22401 - Investimento - Predial", "22402 - Investimento - Informática",
    "22403 - Investimento - Elétrica", "22404 - Investimento - Maquinas e Equipamentos",
    "22405 - Investimento - Refrigeração/Ar-condicionado",
    "22406 - Investimento - Mobiliário/Decoração",
    "22407 - Investimento - Consultorias e Prestações de Serviços",
    "22408 - Investimento - Marcas e Patentes", "22409 - Investimento - Novas Unidades",
    "22410 - Investimento - Reformas", "22501 - Retirada de Sócios",
    "22601 - Juros Cheque Especial / IOF", "22602 - Juros por Atraso de Pagamentos",
    "22701 - Contrato de Mutuo - Débito", "TRANSFERENCIA",
    "TRANSFERENCIA MATRIZ - PIX Filial", "CARTÃO DE CREDITO", "A CLASSIFICAR",
]))

RESPONSAVEIS_LISTA = [
    "ADM/FINANCEIRO", "COMPRAS", "MARKETING", "PLANEJAMENTO", "GERENTE ONLINE",
    "GERENTE LOJA", "GERENTE PRODUCAO", "PRODUCAO", "SUPRIMENTOS", "LOGISTICA",
    "RH", "DIRETORIA"
]

TIPOS_IMPOSTO = ["PIS", "COFINS", "IRRF", "INSS retido", "ISS", "CSLL"]

STATUS_NOTA = ["PENDENTE", "APROVADA", "PAGA", "VENCIDA", "CANCELADA"]

ALIQUOTAS_PADRAO = {
    "PIS": 0.65,
    "COFINS": 3.0,
    "IRRF": 1.5,
    "INSS retido": 11.0,
    "ISS": 5.0,
}

# --- CONFIGURAÇÕES CNAB ---
CNAB_CONTAS = {
    "ITAÚ": {"agencia": "0000", "conta": "00000-0"},
    "SANTANDER": {"agencia": "1111", "conta": "11111-1"},
}

# --- APIS E INTEGRAÇÕES ---
PAGARME_API_URL = "https://api.pagar.me/1"
PAGARME_API_KEY = "sua_chave_aqui"

REDE_API_URL = "https://api.userede.com.br"
REDE_CLIENT_ID = "seu_id"
REDE_CLIENT_SECRET = "seu_secret"
REDE_PV = "seu_pv"

GETNET_API_URL = "https://api.getnet.com.br"
GETNET_CLIENT_ID = "seu_id"
GETNET_CLIENT_SECRET = "seu_secret"
GETNET_SELLER_ID = "seu_seller"

MICROVIX_API_URL = "https://api.microvix.com.br"
MICROVIX_USUARIO = "usuario"
MICROVIX_SENHA   = "senha"
MICROVIX_CNPJ    = "00000000000000"
MICROVIX_CHAVE   = "chave"

# Configurações de E-mail (Destinatários padrão)
EMAIL_DESTINO_DP      = "dp@suaempresa.com"
EMAIL_DESTINO_EXTRATO = "financeiro@suaempresa.com"
EMAIL_DESTINO_CONTAB  = "contabil@suaempresa.com"
EMAIL_REMETENTE       = "robo@suaempresa.com"
EMAIL_SENHA           = "sua_senha"
SMTP_HOST             = "smtp.gmail.com"
SMTP_PORT             = 465

# --- RATEIO E FILIAIS ---
FILIAIS_RATEIO = ["LALUA MATRIZ", "LALUA FILIAL 1", "LALUA FILIAL 2", "SOLAR MATRIZ"]
DE_PARA_FILIAL = {
    "LALUA": "LALUA MATRIZ",
    "SOLAR": "SOLAR MATRIZ",
}


