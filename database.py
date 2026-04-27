"""
database.py — Gerenciamento do banco de dados SQLite do Robô BOAH.
"""
import sqlite3
import os
from datetime import datetime

# Caminho padrão do banco de dados (mesma pasta do script)
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "robo_boah.db")


def _get_conn():
    """Retorna uma conexão com o banco de dados, criando as tabelas se necessário."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    _db_init(conn)
    return conn


def _db_init(conn):
    """Cria todas as tabelas necessárias se ainda não existirem."""
    cursor = conn.cursor()

    # ── Tabela principal de pagamentos autorizados ──────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pagamentos_autorizados (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            nome            TEXT    NOT NULL,
            cnpj            TEXT    DEFAULT '',
            tipo            TEXT    DEFAULT '',
            data            TEXT    NOT NULL,
            valor           REAL    NOT NULL,
            status          TEXT    DEFAULT 'Aprovada',
            responsavel     TEXT    DEFAULT '',
            categoria       TEXT    DEFAULT 'A CLASSIFICAR',
            empresa         TEXT    DEFAULT 'LALUA',
            observacao      TEXT    DEFAULT '',
            data_aprovacao  TEXT    NOT NULL
        )
    """)

    # ── Tabela de Notas Fiscais / Contas a Pagar ────────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS notas (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_tx       TEXT    UNIQUE,
            tipo            TEXT    DEFAULT 'NOTA',
            fornecedor      TEXT    NOT NULL,
            cnpj            TEXT,
            dt_emissao      TEXT,
            dt_vencimento   TEXT,
            valor_bruto     REAL,
            valor_liquido   REAL,
            status          TEXT    DEFAULT 'PENDENTE',
            categoria       TEXT,
            empresa         TEXT,
            observacao      TEXT,
            responsavel     TEXT,
            descricao       TEXT,
            tipo_operacao   TEXT    DEFAULT 'DESPESA',
            natureza        TEXT    DEFAULT 'VENDA',
            valor_difal     REAL    DEFAULT 0,
            valor_fcp       REAL    DEFAULT 0,
            chave_ref       TEXT,
            recorrente      INTEGER DEFAULT 0,
            id_recorrencia  TEXT,
            pix_chave       TEXT,
            banco_dest      TEXT,
            agencia_dest    TEXT,
            conta_dest      TEXT,
            cpf_cnpj_dest   TEXT,
            cod_barras      TEXT
        )
    """)

    # ── Migração: Adiciona colunas novas caso a tabela já exista ────────────
    try:
        # Tenta adicionar colunas uma a uma (ignora erro se já existir)
        novas_colunas = [
            ("descricao",     "TEXT"),
            ("tipo_operacao", "TEXT DEFAULT 'DESPESA'"),
            ("natureza",      "TEXT DEFAULT 'VENDA'"),
            ("valor_difal",   "REAL DEFAULT 0"),
            ("valor_fcp",     "REAL DEFAULT 0"),
            ("chave_ref",     "TEXT"),
            ("recorrente",    "INTEGER DEFAULT 0"),
            ("id_recorrencia","TEXT"),
            ("pix_chave",     "TEXT"),
            ("banco_dest",    "TEXT"),
            ("agencia_dest",  "TEXT"),
            ("conta_dest",    "TEXT"),
            ("cpf_cnpj_dest", "TEXT"),
            ("cod_barras",    "TEXT")
        ]
        for col, tip in novas_colunas:
            try:
                cursor.execute(f"ALTER TABLE notas ADD COLUMN {col} {tip}")
            except: pass
    except: pass
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS nota_itens (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            nota_id         INTEGER NOT NULL,
            descricao       TEXT,
            quantidade      REAL,
            unidade         TEXT,
            valor_unitario  REAL,
            valor_total     REAL,
            FOREIGN KEY(nota_id) REFERENCES notas(id)
        )
    """)
    conn.commit()

    # ── Tabela de Fornecedores ──────────────────────────────────────────────

    # ── Tabela de Fornecedores ──────────────────────────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS fornecedores (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            razao_social    TEXT    NOT NULL,
            nome_fantasia   TEXT,
            cnpj_cpf        TEXT    UNIQUE,
            tipo            TEXT,
            categoria       TEXT,
            responsavel     TEXT,
            empresa         TEXT,
            filial_padrao   TEXT,
            ativo           INTEGER DEFAULT 1
        )
    """)

    conn.commit()


# ── PAGAMENTOS AUTORIZADOS ──────────────────────────────────────────────────

def salvar_pagamentos_autorizados(lista_pagamentos: list) -> int:
    """
    Salva os pagamentos autorizados no banco de dados.

    Evita duplicidade verificando a combinação de (nome, valor, data, empresa).
    Retorna o número de novos registros inseridos.
    """
    if not lista_pagamentos:
        return 0

    conn = _get_conn()
    cursor = conn.cursor()
    inseridos = 0
    data_aprovacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    try:
        for p in lista_pagamentos:
            nome   = str(p.get("nome",   "")).strip()
            valor  = float(p.get("valor", 0.0))
            data   = str(p.get("data",   "")).strip()
            empresa = str(p.get("empresa", "LALUA")).strip()

            # Verifica duplicidade: mesmo nome + valor + data + empresa
            cursor.execute("""
                SELECT id FROM pagamentos_autorizados
                WHERE nome = ? AND valor = ? AND data = ? AND empresa = ?
                LIMIT 1
            """, (nome, valor, data, empresa))

            if cursor.fetchone():
                continue  # Já existe, pula

            cursor.execute("""
                INSERT INTO pagamentos_autorizados
                    (nome, cnpj, tipo, data, valor, status, responsavel,
                     categoria, empresa, observacao, data_aprovacao)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                nome,
                str(p.get("cnpj",        "")).strip(),
                str(p.get("tipo",        "")).strip(),
                data,
                valor,
                str(p.get("status",      "Aprovada")).strip(),
                str(p.get("responsavel", "")).strip(),
                str(p.get("categoria",   "A CLASSIFICAR")).strip(),
                empresa,
                str(p.get("observacao",  "")).strip(),
                data_aprovacao,
            ))
            inseridos += 1

        conn.commit()
    finally:
        conn.close()

    return inseridos


def listar_pagamentos_autorizados(
    data_de: str = None,
    data_ate: str = None,
    empresa: str = None
) -> list:
    """
    Retorna uma lista de dicts com todos os pagamentos autorizados salvos.

    Parâmetros opcionais de filtro (formato: "DD/MM/YYYY"):
        data_de  — data inicial inclusive
        data_ate — data final inclusive
        empresa  — "LALUA", "SOLAR" ou None para ambas
    """
    conn = _get_conn()
    cursor = conn.cursor()

    try:
        cursor.execute("""
            SELECT id, nome, cnpj, tipo, data, valor, status,
                   responsavel, categoria, empresa, observacao, data_aprovacao
            FROM pagamentos_autorizados
            ORDER BY data ASC, nome ASC
        """)
        rows = cursor.fetchall()
    finally:
        conn.close()

    # Converte para dicts e aplica filtros em Python (formato dd/mm/yyyy)
    resultado = []
    for row in rows:
        r = dict(row)

        # Filtro de empresa
        if empresa and r.get("empresa", "").upper() != empresa.upper():
            continue

        # Filtro de período: converte data "DD/MM/YYYY" → datetime para comparar
        try:
            dt = datetime.strptime(r["data"], "%d/%m/%Y")
            if data_de:
                dt_de = datetime.strptime(data_de, "%d/%m/%Y")
                if dt < dt_de:
                    continue
            if data_ate:
                dt_ate = datetime.strptime(data_ate, "%d/%m/%Y")
                if dt > dt_ate:
                    continue
            r["data_obj"] = dt
        except ValueError:
            r["data_obj"] = None

        resultado.append(r)

    return resultado


def excluir_pagamento_autorizado(id_registro: int) -> bool:
    """
    Remove um registro pelo id.
    Retorna True se excluído com sucesso, False caso não encontrado.
    """
    conn = _get_conn()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM pagamentos_autorizados WHERE id = ?", (id_registro,))
        conn.commit()
        return cursor.rowcount > 0
    finally:
        conn.close()

# ── NOTAS E FORNECEDORES ──────────────────────────────────────────────────

def db_conn(): return _get_conn()

def db_init():
    conn = _get_conn()
    _db_init_internal(conn)
    conn.close()

def _db_init_internal(conn):
    cursor = conn.cursor()
    # Corrige tabela nota_impostos caso já exista com schema antigo
    try:
        cursor.execute("SELECT tipo FROM nota_impostos LIMIT 1")
    except sqlite3.OperationalError:
        cursor.execute("DROP TABLE IF EXISTS nota_impostos")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS adiantamentos (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            fornecedor      TEXT    NOT NULL,
            data            TEXT    NOT NULL,
            valor           REAL    NOT NULL,
            status          TEXT    DEFAULT 'PENDENTE',
            observacao      TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS nota_impostos (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            nota_id         INTEGER NOT NULL,
            tipo            TEXT    NOT NULL,
            aliquota        REAL    DEFAULT 0,
            valor           REAL    NOT NULL,
            dt_venc_imp     TEXT,
            status          TEXT    DEFAULT 'PENDENTE',
            numero_doc      TEXT,
            FOREIGN KEY(nota_id) REFERENCES notas(id)
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS nota_parcelas (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            nota_id         INTEGER NOT NULL,
            numero          INTEGER,
            total           INTEGER,
            dt_vencimento   TEXT,
            valor           REAL,
            status          TEXT    DEFAULT 'PENDENTE',
            FOREIGN KEY(nota_id) REFERENCES notas(id)
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS nota_rateio (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            nota_id         INTEGER NOT NULL,
            filial          TEXT,
            categoria       TEXT,
            percentual      REAL,
            valor           REAL,
            FOREIGN KEY(nota_id) REFERENCES notas(id)
        )
    """)
    conn.commit()

def listar_notas(status=None, empresa=None, busca=None):
    conn = _get_conn()
    cursor = conn.cursor()
    query = "SELECT * FROM notas WHERE 1=1"
    params = []
    if status:
        query += " AND status = ?"
        params.append(status)
    if empresa:
        query += " AND empresa = ?"
        params.append(empresa)
    if busca:
        query += " AND (fornecedor LIKE ? OR numero_tx LIKE ?)"
        params.extend([f"%{busca}%", f"%{busca}%"])
    query += " ORDER BY dt_vencimento ASC"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def salvar_nota(dados):
    conn = _get_conn()
    cursor = conn.cursor()
    
    # Separa listas aninhadas
    impostos = dados.pop("impostos", [])
    parcelas = dados.pop("parcelas", [])
    rateio   = dados.pop("rateio",   [])
    itens    = dados.pop("itens",    [])
    
    # Gera numero_tx se não houver
    if not dados.get("numero_tx"):
        # Se for recorrente, a chave já pode estar vindo pronta
        dados["numero_tx"] = f"TX{int(time.time()*100)}"
        
    keys = dados.keys()
    vals = [dados[k] for k in keys]
    query = f"INSERT OR REPLACE INTO notas ({','.join(keys)}) VALUES ({','.join(['?']*len(keys))})"
    cursor.execute(query, vals)
    nota_id = cursor.lastrowid
    
    # Se foi REPLACE, o lastrowid pode ser 0 ou não ser o ID real. Busca pelo numero_tx.
    if not nota_id or nota_id == 0:
        row = cursor.execute("SELECT id FROM notas WHERE numero_tx=?", (dados["numero_tx"],)).fetchone()
        if row: nota_id = row[0]

    # Limpa aninhados antigos para evitar duplicidade em caso de update
    cursor.execute("DELETE FROM nota_impostos WHERE nota_id=?", (nota_id,))
    cursor.execute("DELETE FROM nota_parcelas WHERE nota_id=?", (nota_id,))
    cursor.execute("DELETE FROM nota_rateio   WHERE nota_id=?", (nota_id,))
    cursor.execute("DELETE FROM nota_itens    WHERE nota_id=?", (nota_id,))

    for imp in impostos:
        cursor.execute("""
            INSERT INTO nota_impostos (nota_id, tipo, aliquota, valor, dt_venc_imp)
            VALUES (?, ?, ?, ?, ?)
        """, (nota_id, imp["tipo"], imp.get("aliquota",0), imp["valor"], imp.get("dt_venc_imp")))

    for parc in parcelas:
        cursor.execute("""
            INSERT INTO nota_parcelas (nota_id, numero, total, dt_vencimento, valor)
            VALUES (?, ?, ?, ?, ?)
        """, (nota_id, parc["numero"], parc["total"], parc["dt_vencimento"], parc["valor"]))

    for rat in rateio:
        cursor.execute("""
            INSERT INTO nota_rateio (nota_id, filial, categoria, percentual, valor)
            VALUES (?, ?, ?, ?, ?)
        """, (nota_id, rat["filial"], rat["categoria"], rat["percentual"], rat["valor"]))

    for it in itens:
        cursor.execute("""
            INSERT INTO nota_itens (nota_id, descricao, quantidade, unidade, valor_unitario, valor_total)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (nota_id, it["descricao"], it.get("quantidade",1), it.get("unidade","UN"), it.get("valor_unitario",0), it["valor_total"]))

    conn.commit()
    conn.close()
    return dados["numero_tx"]

def atualizar_status_nota(id_nota, novo_status):
    conn = _get_conn()
    cursor = conn.cursor()
    cursor.execute("UPDATE notas SET status = ? WHERE id = ?", (novo_status, id_nota))
    conn.commit()
    conn.close()

def listar_fornecedores(busca=None, ativo_only=True):
    conn = _get_conn()
    cursor = conn.cursor()
    query = "SELECT * FROM fornecedores WHERE 1=1"
    params = []
    if busca:
        query += " AND (razao_social LIKE ? OR cnpj_cpf LIKE ?)"
        params.extend([f"%{busca}%", f"%{busca}%"])
    if ativo_only:
        query += " AND ativo = 1"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def buscar_fornecedor_por_cnpj(cnpj):
    conn = _get_conn()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM fornecedores WHERE cnpj_cpf = ?", (cnpj,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None

def salvar_fornecedor(dados):
    conn = _get_conn()
    cursor = conn.cursor()
    keys = dados.keys()
    vals = [dados[k] for k in keys]
    query = f"INSERT OR REPLACE INTO fornecedores ({','.join(keys)}) VALUES ({','.join(['?']*len(keys))})"
    cursor.execute(query, vals)
    conn.commit()
    conn.close()

def listar_adiantamentos(busca=None, status=None):
    conn = _get_conn()
    cursor = conn.cursor()
    query = "SELECT * FROM adiantamentos WHERE 1=1"
    params = []
    if busca:
        query += " AND (fornecedor LIKE ? OR observacao LIKE ?)"
        params.extend([f"%{busca}%", f"%{busca}%"])
    if status:
        query += " AND status = ?"
        params.append(status)
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def salvar_adiantamento(dados):
    conn = _get_conn()
    cursor = conn.cursor()
    # Garante que a tabela existe
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS adiantamentos (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            fornecedor      TEXT    NOT NULL,
            data            TEXT    NOT NULL,
            valor           REAL    NOT NULL,
            status          TEXT    DEFAULT 'PENDENTE',
            observacao      TEXT
        )
    """)
    keys = dados.keys()
    vals = [dados[k] for k in keys]
    query = f"INSERT OR REPLACE INTO adiantamentos ({','.join(keys)}) VALUES ({','.join(['?']*len(keys))})"
    cursor.execute(query, vals)
    conn.commit()
    conn.close()
