import os
import xlrd
from datetime import datetime
from utils.data_processing import auto_classificar

def ler_extrato_pdf(caminho):
    """Stub para leitura de extrato PDF."""
    return {}, "Período não identificado"

def ler_getnet_excel(caminho):
    """Stub para leitura de excel Getnet."""
    return []

def ler_itau_pagamentos(caminho):
    """Lê pagamentos do Excel (.xls) do Itaú de forma robusta."""
    try:
        workbook = xlrd.open_workbook(caminho)
        sheet = workbook.sheet_by_index(0)
    except Exception as e:
        print(f"Erro ao abrir XLS: {e}")
        return []

    pagamentos = []
    header_idx = -1
    empresa_detectada = "LALUA"

    # 1. Identifica a empresa
    for r in range(min(sheet.nrows, 20)):
        txt = " ".join([str(x) for x in sheet.row_values(r)]).upper()
        if "SOLAR" in txt: empresa_detectada = "SOLAR"
        elif "LALUA" in txt: empresa_detectada = "LALUA"

    # 2. Busca o cabeçalho
    for r in range(sheet.nrows):
        row_vals = [str(c).strip().lower() for c in sheet.row_values(r)]
        if any("favorecido" in v or "beneficiário" in v for v in row_vals):
            header_idx = r
            break
    
    if header_idx == -1:
        return []

    # 3. Processa as linhas
    for r in range(header_idx + 1, sheet.nrows):
        row = sheet.row_values(r)
        if not row or not row[0] or str(row[0]).strip().lower() == "total:":
            continue
        
        try:
            nome = str(row[0]).strip()
            cnpj = str(row[1]).strip()
            tipo = str(row[2]).strip()
            data_str = str(row[4]).strip()
            valor = 0.0
            try:
                valor = float(row[5])
            except: pass

            data_obj = None
            if data_str:
                try:
                    data_obj = datetime.strptime(data_str, "%d/%m/%Y")
                except: pass

            resp, cat = auto_classificar(nome, valor, cnpj)

            pagamentos.append({
                "nome": nome, "cnpj": cnpj, "tipo": tipo, "data": data_str,
                "data_obj": data_obj, "valor": valor,
                "status": str(row[6]).strip() if len(row) > 6 else "Aprovada",
                "responsavel": resp, "categoria": cat, "empresa": empresa_detectada,
                "observacao": f"Importado de {os.path.basename(caminho)}",
                "manual": False
            })
        except: continue

    return pagamentos

def ler_itau_folha(caminho):
    """Lê folha de pagamentos do Itaú (simplificado)."""
    if caminho.lower().endswith(".xls"):
        itens = ler_itau_pagamentos(caminho)
        for i in itens:
            i["categoria"] = "RH - SALARIOS"
        return itens
    return []
