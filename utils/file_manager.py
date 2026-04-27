"""
utils/file_manager.py — Gerenciamento de arquivos do Robô Financeiro BOAH/SOLAR.

Responsável por:
  - Organizar PDFs de comprovantes da pasta ENTRADA → SAIDA_ORGANIZADA
  - Espelhar comprovantes no servidor (Y:\\...)
  - Fazer backup local dos comprovantes
  - Rotação automática de logs (remove arquivos mais antigos que LOG_RETENCAO_DIAS)
"""

from __future__ import annotations

import os
import re
import shutil
import time
import glob
from datetime import datetime, timedelta
from typing import Callable, Optional

from config import (
    PASTA_ENTRADA,
    PASTA_SAIDA,
    PASTA_ATUAL,
    PASTA_LOGS,
    PASTA_SERVIDOR,
    PASTA_BACKUP_PDF,
    LOG_RETENCAO_DIAS,
)


# ==============================================================================
# ── UTILITÁRIOS INTERNOS ──────────────────────────────────────────────────────
# ==============================================================================

def _normalizar(texto: str) -> str:
    """Remove caracteres inválidos para nome de pasta/arquivo."""
    return re.sub(r'[<>:"/\\|?*]', '_', texto).strip()


def _copiar_seguro(origem: str, destino: str, log_fn: Optional[Callable] = None) -> bool:
    """
    Copia `origem` para `destino` de forma segura:
      - Cria diretório de destino se necessário
      - Não sobrescreve arquivos idênticos (verifica tamanho)
      - Retorna True em sucesso, False em falha (sem lançar exceção)
    """
    try:
        os.makedirs(os.path.dirname(destino), exist_ok=True)
        # Evita sobrescrever se arquivo já existe e tem mesmo tamanho
        if os.path.exists(destino):
            if os.path.getsize(destino) == os.path.getsize(origem):
                return True  # já existe, skip silencioso
        shutil.copy2(origem, destino)
        return True
    except Exception as exc:
        if log_fn:
            log_fn(f"⚠️ Não foi possível copiar para {destino!r}: {exc}")
        return False


# ==============================================================================
# ── ROTAÇÃO DE LOGS ───────────────────────────────────────────────────────────
# ==============================================================================

def limpar_logs_antigos(log_fn: Optional[Callable] = None) -> int:
    """
    Remove arquivos de log com mais de LOG_RETENCAO_DIAS dias da PASTA_LOGS.
    Retorna o número de arquivos removidos.
    """
    limite = datetime.now() - timedelta(days=LOG_RETENCAO_DIAS)
    removidos = 0

    padrao = os.path.join(PASTA_LOGS, "log_*.txt")
    for caminho in glob.glob(padrao):
        try:
            mtime = datetime.fromtimestamp(os.path.getmtime(caminho))
            if mtime < limite:
                os.remove(caminho)
                removidos += 1
        except Exception as exc:
            if log_fn:
                log_fn(f"⚠️ Erro ao remover log {os.path.basename(caminho)}: {exc}")

    if removidos and log_fn:
        log_fn(f"🧹 Rotação de logs: {removidos} arquivo(s) removido(s) "
               f"(mais antigos que {LOG_RETENCAO_DIAS} dias)")
    return removidos


# ==============================================================================
# ── ESPELHAMENTO: SERVIDOR + BACKUP LOCAL ─────────────────────────────────────
# ==============================================================================

def _espelhar_pdf(
    caminho_origem: str,
    caminho_relativo: str,
    log_fn: Optional[Callable] = None,
) -> None:
    """
    Copia um PDF gerado para:
      1. Servidor: PASTA_SERVIDOR / caminho_relativo
      2. Backup local: PASTA_BACKUP_PDF / caminho_relativo

    `caminho_relativo` deve ser o sub-caminho a partir da raiz de saída,
    ex: "LALUA/2026/04_Abril/23_Dia/FORNECEDORES/boleto_fulano.pdf"
    """
    # ── Cópia para servidor ──────────────────────────────────────────────────
    destino_srv = os.path.join(PASTA_SERVIDOR, caminho_relativo)
    ok_srv = _copiar_seguro(caminho_origem, destino_srv, log_fn)
    if ok_srv and log_fn:
        log_fn(f"☁️  Servidor: {os.path.basename(caminho_origem)} → salvo")

    # ── Backup local ─────────────────────────────────────────────────────────
    destino_bkp = os.path.join(PASTA_BACKUP_PDF, caminho_relativo)
    ok_bkp = _copiar_seguro(caminho_origem, destino_bkp, log_fn)
    if ok_bkp and log_fn:
        log_fn(f"💾 Backup local: {os.path.basename(caminho_origem)} → salvo")


# ==============================================================================
# ── ORGANIZAÇÃO DE COMPROVANTES ───────────────────────────────────────────────
# ==============================================================================

# Expressão regular para identificar empresa no nome do PDF
_RE_EMPRESA = re.compile(
    r'\b(LALUA|SOLAR|BOAH)\b', re.IGNORECASE
)

# Mapeamento de número de mês para nome de pasta
_MESES = {
    1: "01_Janeiro", 2: "02_Fevereiro", 3: "03_Marco",
    4: "04_Abril",   5: "05_Maio",      6: "06_Junho",
    7: "07_Julho",   8: "08_Agosto",    9: "09_Setembro",
    10: "10_Outubro", 11: "11_Novembro", 12: "12_Dezembro",
}

# Categorias inferidas pelo nome do arquivo
_CATEGORIAS_REGEX = [
    (re.compile(r'\bFORNECED', re.I),    "FORNECEDORES"),
    (re.compile(r'\bSALARIO|RH\b', re.I), "RH - SALARIOS"),
    (re.compile(r'\bFGTS\b', re.I),       "RH - FGTS"),
    (re.compile(r'\bIMPOSTO|GNRE|DARF\b', re.I), "IMPOSTOS"),
    (re.compile(r'\bTRANSF', re.I),       "TRANSFERENCIA"),
]


def _inferir_categoria(nome_arquivo: str) -> str:
    for regex, cat in _CATEGORIAS_REGEX:
        if regex.search(nome_arquivo):
            return cat
    return "A CLASSIFICAR"


def _inferir_empresa(nome_arquivo: str) -> str:
    m = _RE_EMPRESA.search(nome_arquivo)
    return m.group(0).upper() if m else "A Classificar"


def _destino_organizado(
    nome_arquivo: str,
    data_ref: Optional[datetime] = None,
    empresa: Optional[str] = None,
    categoria: Optional[str] = None,
) -> tuple[str, str]:
    """
    Calcula o caminho de destino dentro de PASTA_SAIDA para um PDF.

    Retorna (caminho_absoluto, caminho_relativo).
    Estrutura: PASTA_SAIDA / {empresa} / {ano} / {mês} / {dia}_Dia / {categoria} / arquivo.pdf
    """
    dt = data_ref or datetime.now()
    empresa_dir  = _normalizar(empresa  or _inferir_empresa(nome_arquivo))
    cat_dir      = _normalizar(categoria or _inferir_categoria(nome_arquivo))
    ano_dir      = str(dt.year)
    mes_dir      = _MESES.get(dt.month, f"{dt.month:02d}_Mes")
    dia_dir      = f"{dt.day:02d}_Dia"

    nome_upper = nome_arquivo.upper()
    if "SEFAZ" in nome_upper or "IMPOSTO_ESTADUAL" in nome_upper or cat_dir == "IMPOSTOS":
        # Extrai a UF se existir no nome, ex: SEFAZ_SP
        match_uf = re.search(r'SEFAZ[_\-\/]?([A-Z]{2})', nome_upper)
        uf_dir = match_uf.group(1) if match_uf else "GERAL"
        relativo = os.path.join(empresa_dir, ano_dir, mes_dir, "GNRE_SEFAZ", uf_dir, nome_arquivo)
    else:
        relativo = os.path.join(empresa_dir, ano_dir, mes_dir, dia_dir, cat_dir, nome_arquivo)

    absoluto = os.path.join(PASTA_SAIDA, relativo)
    return absoluto, relativo


def mover_para_saida(
    caminho_origem: str,
    data_ref: Optional[datetime] = None,
    empresa: Optional[str] = None,
    categoria: Optional[str] = None,
    log_fn: Optional[Callable] = None,
    espelhar: bool = True,
) -> Optional[str]:
    """
    Move um PDF de comprovante para a estrutura SAIDA_ORGANIZADA e,
    opcionalmente, espelha no servidor e no backup local.

    Retorna o caminho final do arquivo em SAIDA_ORGANIZADA, ou None em erro.
    """
    nome = os.path.basename(caminho_origem)
    destino_abs, destino_rel = _destino_organizado(nome, data_ref, empresa, categoria)

    try:
        os.makedirs(os.path.dirname(destino_abs), exist_ok=True)

        # Se já existe arquivo com mesmo nome, adiciona timestamp
        if os.path.exists(destino_abs):
            base, ext = os.path.splitext(nome)
            ts = datetime.now().strftime("%H%M%S")
            destino_abs = destino_abs.replace(nome, f"{base}_{ts}{ext}")
            destino_rel = destino_rel.replace(nome, f"{base}_{ts}{ext}")

        shutil.move(caminho_origem, destino_abs)

        if log_fn:
            log_fn(f"📂 Organizado: {nome} → {destino_rel}")

        # ── Espelha no servidor + backup local ──────────────────────────────
        if espelhar:
            _espelhar_pdf(destino_abs, destino_rel, log_fn)

        return destino_abs

    except Exception as exc:
        if log_fn:
            log_fn(f"❌ Erro ao mover {nome}: {exc}")
        return None


def copiar_pdf_para_servidor(
    caminho_pdf: str,
    caminho_relativo: Optional[str] = None,
    log_fn: Optional[Callable] = None,
) -> bool:
    """
    Copia um PDF (já gerado em qualquer pasta) para o servidor e backup local.
    Use esta função para PDFs gerados diretamente em PASTA_ATUAL (relatórios,
    transferências, etc.) que não passam pelo fluxo de organização.

    Se `caminho_relativo` for None, usa apenas o nome do arquivo como sub-pasta
    datada: YYYY/MM_Mes/DD_Dia/<arquivo>.pdf
    """
    nome = os.path.basename(caminho_pdf)

    if caminho_relativo is None:
        dt = datetime.now()
        mes_dir = _MESES.get(dt.month, f"{dt.month:02d}_Mes")
        caminho_relativo = os.path.join(
            str(dt.year), mes_dir, f"{dt.day:02d}_Dia", nome
        )

    _espelhar_pdf(caminho_pdf, caminho_relativo, log_fn)
    return True


# ==============================================================================
# ── ORGANIZAÇÃO EM LOTE (chamada pelo robô) ───────────────────────────────────
# ==============================================================================

def organizar_arquivos(
    log_fn: Optional[Callable] = None,
    atualizar_stats: Optional[Callable] = None,
) -> int:
    """
    Processa todos os PDFs em PASTA_ENTRADA:
      1. Move cada PDF para a estrutura organizada em PASTA_SAIDA
      2. Espelha no servidor (Y:\\...) e no backup local
      3. Executa rotação de logs ao final

    Retorna o total de arquivos processados com sucesso.
    """
    pdfs = [
        f for f in os.listdir(PASTA_ENTRADA)
        if f.lower().endswith(".pdf")
    ]

    if not pdfs:
        if log_fn:
            log_fn("ℹ️ Nenhum PDF encontrado na pasta ENTRADA.")
        return 0

    processados = 0
    duplicados  = 0

    for nome in pdfs:
        caminho = os.path.join(PASTA_ENTRADA, nome)

        # Detecta possível duplicata pelo nome (arquivo já na saída)
        _, destino_rel = _destino_organizado(nome)
        destino_abs = os.path.join(PASTA_SAIDA, destino_rel)
        if os.path.exists(destino_abs):
            # Verifica por tamanho
            if os.path.getsize(caminho) == os.path.getsize(destino_abs):
                if log_fn:
                    log_fn(f"🔁 Duplicado ignorado: {nome}")
                duplicados += 1
                if atualizar_stats:
                    atualizar_stats(processados, duplicados)
                # Remove da entrada mesmo assim para não acumular
                try:
                    os.remove(caminho)
                except Exception:
                    pass
                continue

        resultado = mover_para_saida(caminho, log_fn=log_fn, espelhar=True)
        if resultado:
            processados += 1
        if atualizar_stats:
            atualizar_stats(processados, duplicados)

        # Pequena pausa para não sobrecarregar o servidor
        time.sleep(0.05)

    # ── Rotação de logs ──────────────────────────────────────────────────────
    limpar_logs_antigos(log_fn)

    if log_fn:
        log_fn(f"\n✅ Concluído: {processados} processados, {duplicados} duplicados.")

    return processados


# ==============================================================================
# ── UTILITÁRIOS PÚBLICOS ──────────────────────────────────────────────────────
# ==============================================================================

def contar_pdfs_entrada() -> int:
    """Retorna quantos PDFs estão aguardando na PASTA_ENTRADA."""
    try:
        return sum(1 for f in os.listdir(PASTA_ENTRADA) if f.lower().endswith(".pdf"))
    except Exception:
        return 0
