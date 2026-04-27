import xml.etree.ElementTree as ET
import os

def _get_text(node, tag, default=""):
    """Busca o texto de uma tag ignorando namespaces."""
    for el in node.iter():
        if el.tag.split('}')[-1] == tag:
            return el.text if el.text else default
    return default

def _get_float(node, tag, default=0.0):
    """Busca um valor float de uma tag ignorando namespaces."""
    text = _get_text(node, tag)
    try:
        return float(text) if text else default
    except ValueError:
        return default

def ler_xml_nfe(caminho_xml):
    """
    Extrai dados detalhados de um XML de NFe (Venda ou Devolução).
    Focado em precisão para cruzamento com GNREs.
    """
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        
        # O nó principal costuma ser NFe ou nfeProc
        infNFe = None
        for el in root.iter():
            if el.tag.split('}')[-1] == 'infNFe':
                infNFe = el
                break
        
        if infNFe is None: return None

        # Dados Identificação
        nNF = _get_text(infNFe, 'nNF', "0")
        serie = _get_text(infNFe, 'serie', "0")
        finNFe = _get_text(infNFe, 'finNFe', "1") # 1=Normal, 4=Devolução
        
        # Data Emissão
        dhEmi = _get_text(infNFe, 'dhEmi')
        if not dhEmi:
            dhEmi = _get_text(infNFe, 'dEmi')
        data_emissao = dhEmi[:10] if dhEmi else ""
        
        # Chave de Acesso (ID da infNFe ou via protNFe)
        chave = infNFe.get('Id', '')
        if chave.startswith('NFe'):
            chave = chave[3:]
        
        # Referência (para Devoluções)
        chave_original = None
        if finNFe == "4":
            chave_original = _get_text(infNFe, 'refNFe')

        # UF e Dados do Destinatário
        uf = "GERAL"
        dest = None
        for el in infNFe.iter():
            if el.tag.split('}')[-1] == 'dest':
                dest = el
                break
        
        if dest is not None:
            uf = _get_text(dest, 'UF', "GERAL")
            cnpj_dest = _get_text(dest, 'CNPJ')
            cpf_dest = _get_text(dest, 'CPF')
        else:
            cnpj_dest = cpf_dest = ""

        # Totais de Impostos (Onde a GNRE se baseia)
        # vICMS: ICMS Próprio
        # vST: ICMS Substituição Tributária
        # vICMSUFDest: DIFAL (Destino)
        # vFCPUFDest: FCP (Destino)
        
        vNF = _get_float(infNFe, 'vNF')
        vICMS = _get_float(infNFe, 'vICMS')
        vST = _get_float(infNFe, 'vST')
        vICMSUFDest = _get_float(infNFe, 'vICMSUFDest')
        vFCPUFDest = _get_float(infNFe, 'vFCPUFDest')
        
        # Valor total de imposto estadual que pode gerar uma GNRE
        valor_imposto_estadual = vICMS + vST + vICMSUFDest + vFCPUFDest

        return {
            "nf_numero": nNF,
            "serie": serie,
            "chave_acesso": chave,
            "finalidade": "DEVOLUCAO" if finNFe == "4" else "VENDA",
            "data_emissao": data_emissao,
            "uf_destino": uf,
            "cnpj_dest": cnpj_dest or cpf_dest,
            "valor_total_nota": vNF,
            "valor_icms_proprio": vICMS,
            "valor_icms_st": vST,
            "valor_difal": vICMSUFDest,
            "valor_fcp": vFCPUFDest,
            "valor_imposto_estadual": round(valor_imposto_estadual, 2),
            "chave_nf_original": chave_original,
            "caminho_arquivo": caminho_xml
        }
    except Exception as e:
        print(f"Erro ao ler XML {caminho_xml}: {e}")
        return None
