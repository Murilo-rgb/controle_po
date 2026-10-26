#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Servidor local do "Controle de PO's".

Fluxo:
  - Serve o index.html em http://localhost:8000
  - GET  /api/pos     -> le db_po.controle_po e devolve JSON para a tela
  - POST /api/import  -> le o .xlsx exportado do SAP, mapeia as colunas,
                         grava em db_po.controle_po e devolve os dados atualizados
  - POST /api/clear   -> limpa a tabela (botao "Limpar tudo")

Nao precisa instalar nada: usa apenas stdlib + psycopg2 + openpyxl.
Iniciar manualmente:  python server.py
"""
import os
import re
import io
import csv
import sys
import glob
import json
import base64
import unicodedata
from datetime import datetime, date
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

import psycopg2
import psycopg2.extras
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# pasta onde os arquivos importados sao salvos/lidos
FILES_DIR = os.path.join(BASE_DIR, "Files")

# ---- configuracao do banco (pode sobrescrever por variaveis de ambiente) ----
DB = dict(
    host=os.environ.get("DB_PGC_HOST", "10.230.43.181"),
    port=int(os.environ.get("DB_PGC_PORT", "5432")),
    user=os.environ.get("DB_PGC_USER", "jose_faria"),
    password=os.environ.get("DB_PGC_PASSWORD", "vXsEha3PYB"),
    dbname=os.environ.get("DB_PGC_DATABASE", "db_gerencial"),
)

SCHEMA = "db_po"
TABLE = "controle_po"
# tabela de vinculo PO <-> arquivo importado (todas as colunas sao TEXT)
TABLE_ARQUIVO = "controle_po_arquivo"

# ---- mapeamento: CABEÇALHO NORMALIZADO da planilha -> CAMPO do banco ----
FIELD_MAP = {
    "documento de compras": "documento_de_compras",
    "item": "item",
    "contrato basico": "contrato_basico",
    "data do documento": "data_documento",
    "preco liquido": "preco_liquido",
    "a ser fornecido valor": "a_ser_fornecido_valor",
    "ainda a faturar valor": "ainda_a_faturar_valor",
    "codigo de eliminacao": "codigo_eliminacao",
    "codigo de imposto": "codigo_imposto",
    "tipo doc compras": "tipo_doc_compras",
    "ctg doc compras": "ctg_doc_compras",
    "fornecedor centro fornecedor": "fornecedor_centro_fornecedor",
    "organizacao compras": "organizacao_compras",
    "grupo compradores": "grupo_compradores",
    "material": "material",
    "texto breve": "texto_breve",
    "grupo mercadorias": "grupo_mercadorias",
    "grupo de mercadorias": "grupo_mercadorias",
    "historico pedido docum solrem": "historico_pedido_docum_sol_rem",
    "categoria do item": "categoria_item",
    "categoria do item_2": "categoria_item_2",
    "ctg classcont": "ctg_class_cont",
    "numero acompanhamento": "numero_acompanhamento",
    "centro": "centro",
    "deposito": "deposito",
    "qtd pedido": "qtd_pedido",
    "um pedido": "um_pedido",
    "qtd na unid gestestoq": "qtd_unid_gest_estoq",
    "unidade manutencao estoque": "unidade_manutencao_estoque",
    "preco liquido": "preco_liquido",
    "moeda": "moeda",
    "unidade preco": "unidade_preco",
    "contrato basico": "contrato_basico",
    "it contrato superior": "it_contrato_superior",
    "valor fixado": "valor_fixado",
    "valor global pendente": "valor_global_pendente",
    "valor pendente": "valor_pendente",
    "valor solicitado": "valor_solicitado",
    "quantidade prevista": "quantidade_prevista",
    "qtd prev pendente": "qtd_prev_pendente",
    "qtd solicit anterior": "qtd_solicit_anterior",
    "inicio per validade": "inicio_per_validade",
    "fim da validade": "fim_validade",
    "a ser fornecida quantidade": "a_ser_fornecida_quantidade",
    "a ser fornecido valor": "a_ser_fornecido_valor",
    "ainda a faturar quantidade": "ainda_a_faturar_quantidade",
    "ainda a faturar valor": "ainda_a_faturar_valor",
    "prz apresent cotacao": "prz_apresent_cotacao",
    "origem": "origem",
    "rfq coletiva": "rfq_coletiva",
    "num itens": "num_itens",
    "codigo de controle": "codigo_controle",
    "reg info compras": "reg_info_compras",
    "numero pacote": "numero_pacote",
    "grupo liberacao": "grupo_liberacao",
    "estrat de liberacao": "estrat_de_liberacao",
    "estado de liberacao": "estado_liberacao",
    "codigo de liberacao": "codigo_liberacao",
    "dep saida do estq": "dep_saida_do_estq",
    "nome de fornecedor": "nome_fornecedor",
    "unid prec pedido": "unid_prec_pedido",
    "domicilio fiscal": "domicilio_fiscal",
    "valor liquido pedido": "valor_liquido_pedido",
    "incompleto": "incompleto",
    "qtd notificada": "qtd_notificada",
    "numero da hierarquia": "numero_hierarquia",
    "numero externo ordenacao": "numero_externo_ordenacao",
    "ctg externa hierarquia": "ctg_externa_hierarquia",
    "urgencia necessidade": "urgencia_necessidade",
    "prior necessario": "prior_neces",
    "numero de documento configurável": "numero_documento_configuravel",
}


def get_conn():
    return psycopg2.connect(connect_timeout=8, **DB)


def ensure_table():
    con = get_conn()
    cur = con.cursor()
    cur.execute("CREATE SCHEMA IF NOT EXISTS %s" % SCHEMA)
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS {s}.{t} (
            id                   BIGSERIAL PRIMARY KEY,
            documento_de_compras TEXT,
            item                 TEXT,
            contrato_basico      TEXT,
            data_documento       DATE,
            preco_liquido        NUMERIC(18,2),
            a_ser_fornecido_valor NUMERIC(18,2),
            ainda_a_faturar_valor NUMERIC(18,2),
            nome_do_servico      TEXT DEFAULT '-',
            competencia          TEXT DEFAULT '-',
            arquivo_origem       TEXT,
            importado_em         TIMESTAMP WITHOUT TIME ZONE DEFAULT now(),
            CONSTRAINT uq_controle_po UNIQUE (documento_de_compras, item, contrato_basico)
        )
        """.format(s=SCHEMA, t=TABLE)
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS {s}.controle_po_aux (
            id                   BIGSERIAL PRIMARY KEY,
            nome_do_projeto      TEXT NOT NULL,
            contrato_basico      TEXT NOT NULL,
            importado_em         TIMESTAMP WITHOUT TIME ZONE DEFAULT now(),
            CONSTRAINT uq_controle_po_aux UNIQUE (nome_do_projeto, contrato_basico)
        )
        """.format(s=SCHEMA)
    )
    # blindagem: garante que a coluna de eliminacao exista mesmo se a
    # tabela foi criada fora desta aplicacao
    cur.execute(
        "ALTER TABLE {s}.{t} ADD COLUMN IF NOT EXISTS codigo_eliminacao TEXT".format(
            s=SCHEMA, t=TABLE
        )
    )
    # ---- tabela de vinculo PO <-> arquivo importado ----
    # Guarda o HISTORICO de quais arquivos contem cada PO.
    # A gravacao e sempre ON CONFLICT DO NOTHING: nunca sobrescreve nem apaga.
    # Todas as colunas sao TEXT (inclusive importado_em), por requisito.
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS {s}.{t} (
            documento_de_compras TEXT,
            item                 TEXT,
            contrato_basico      TEXT,
            arquivo_origem       TEXT,
            importado_em         TEXT DEFAULT to_char(now(), 'YYYY-MM-DD HH24:MI:SS'),
            CONSTRAINT pk_controle_po_arquivo
                UNIQUE (documento_de_compras, item, contrato_basico, arquivo_origem)
        )
        """.format(s=SCHEMA, t=TABLE_ARQUIVO)
    )
    # migracao: se a tabela ja existia com importado_em TIMESTAMP, converte
    # para TEXT preservando os valores (formato 'YYYY-MM-DD HH24:MI:SS'),
    # que ordenam corretamente como texto.
    cur.execute(
        """
        SELECT data_type FROM information_schema.columns
        WHERE table_schema = %s AND table_name = %s AND column_name = 'importado_em'
        """,
        (SCHEMA, TABLE_ARQUIVO),
    )
    row = cur.fetchone()
    if row and row[0].lower() != "text":
        cur.execute(
            """
            ALTER TABLE {s}.{t}
                ALTER COLUMN importado_em TYPE TEXT
                USING to_char(importado_em, 'YYYY-MM-DD HH24:MI:SS')
            """.format(s=SCHEMA, t=TABLE_ARQUIVO)
        )
        cur.execute(
            """
            ALTER TABLE {s}.{t}
                ALTER COLUMN importado_em
                SET DEFAULT to_char(now(), 'YYYY-MM-DD HH24:MI:SS')
            """.format(s=SCHEMA, t=TABLE_ARQUIVO)
        )
    # retroalimentacao: registra o vinculo das POs que ja existem no banco
    # (coluna arquivo_origem) na nova tabela de vinculo. Roda uma unica vez
    # com efeito, pois nas execucoes seguintes nao ha nada novo a copiar.
    cur.execute(
        """
        INSERT INTO {s}.{t} (documento_de_compras, item, contrato_basico, arquivo_origem)
        SELECT documento_de_compras, item, contrato_basico, arquivo_origem
        FROM {s}.{c}
        WHERE arquivo_origem IS NOT NULL AND trim(arquivo_origem) <> ''
        ON CONFLICT DO NOTHING
        """.format(s=SCHEMA, t=TABLE_ARQUIVO, c=TABLE)
    )
    con.commit()
    cur.close()
    con.close()


# ---------- normalizacao / parsing ----------
def norm(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()
    s = re.sub(r"[()]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s in ("", "-"):
        return None
    if "," in s:  # formato BR: 1.234,56
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if s in ("", "-"):
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _rows_from_xlsx_bytes(data):
    """Retorna (header_idx, col_map, linhas) a partir de bytes de um .xlsx."""
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            primeiras = list(ws.iter_rows(min_row=1, max_row=10, values_only=True))
            for hi, row in enumerate(primeiras):
                tmp = {}
                for idx, cell in enumerate(row):
                    if cell is not None and str(cell).strip():
                        tmp[norm(str(cell).strip())] = (idx, str(cell).strip())
                if len(tmp) >= 5:
                    linhas = list(ws.iter_rows(min_row=hi + 2, values_only=True))
                    return hi, tmp, linhas
        return None, None, None
    finally:
        wb.close()


def _rows_from_csv_bytes(data):
    """Retorna (header_idx, col_map, linhas) a partir de bytes de um .csv.

    col_map mapeia o CABEÇALHO NORMALIZADO -> (INDICE_DA_COLUNA, CABEÇALHO_RO),
    para TODOS os cabeçalhos do arquivo (nao so os de FIELD_MAP). O cabecalho
    cru eh guardado pra servir de chave no colunas_extra sem depender da
    normalizacao.
    """
    text = data.decode("utf-8-sig", errors="replace")
    amostra = text[:8192]
    try:
        delim = csv.Sniffer().sniff(amostra, delimiters=";,\t").delimiter
    except Exception:
        delim = ";" if amostra.count(";") >= amostra.count(",") else ","
    linhas = list(csv.reader(io.StringIO(text), delimiter=delim))
    for hi in range(min(10, len(linhas))):
        tmp = {}
        for idx, cell in enumerate(linhas[hi]):
            cell_str = "" if cell is None else str(cell).strip()
            if cell_str:
                tmp[norm(cell_str)] = (idx, cell_str)
        if any(k in FIELD_MAP for k in tmp):
            return hi, tmp, linhas[hi + 1:]
    return None, None, None


def _records_from_rows(col_map, rows):
    recs = []
    PRIMARIOS = set(FIELD_MAP)   # chaves normalizadas dos campos estruturais

    for row in rows:
        def val(header):
            item = col_map.get(header)
            if item is None:
                return None
            idx, _ = item
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        def get_str(header):
            v = val(header)
            return None if v is None else str(v).strip()

        doc = get_str("documento de compras")
        if doc is None or doc == "":
            continue

        rec = {}
        for norm_key, db_col in FIELD_MAP.items():
            rec[db_col] = get_str(norm_key)

        # Campos que não vêm do arquivo, mas precisam existir no banco
        rec.setdefault("nome_do_servico", "-")
        rec.setdefault("competencia", "-")
        rec.setdefault("arquivo_origem", os.path.basename(""))

        recs.append(rec)
    return recs


def extract_records_bytes(name, data):
    """Extrai registros de um arquivo (bytes) .csv ou .xlsx."""
    ext = os.path.splitext(name)[1].lower()
    if ext == ".csv":
        _, col_map, rows = _rows_from_csv_bytes(data)
    else:
        _, col_map, rows = _rows_from_xlsx_bytes(data)
    if not col_map:
        raise ValueError("Colunas esperadas nao encontradas no arquivo.")
    return _records_from_rows(col_map, rows)


def import_file(path):
    with open(path, "rb") as f:
        data = f.read()
    recs = extract_records_bytes(path, data)
    orig = os.path.basename(path)
    for r in recs:
        r["arquivo_origem"] = orig
    con = get_conn()
    cur = con.cursor()
    if recs:
        cols = [
            "documento_de_compras", "item", "contrato_basico", "data_documento",
            "preco_liquido", "a_ser_fornecido_valor", "ainda_a_faturar_valor",
            "codigo_eliminacao", "codigo_imposto", "tipo_doc_compras", "ctg_doc_compras",
            "fornecedor_centro_fornecedor", "organizacao_compras", "grupo_compradores",
            "historico_pedido_docum_sol_rem", "material", "texto_breve", "grupo_mercadorias",
            "categoria_item", "categoria_item_2",
            "ctg_class_cont", "numero_acompanhamento", "centro", "deposito",
            "qtd_pedido", "um_pedido", "qtd_unid_gest_estoq",
            "unidade_manutencao_estoque", "preco_liquido", "moeda", "unidade_preco",
            "contrato_basico", "it_contrato_superior", "valor_fixado",
            "valor_global_pendente", "valor_pendente", "valor_solicitado",
            "quantidade_prevista", "qtd_prev_pendente", "qtd_solicit_anterior",
            "inicio_per_validade", "fim_validade", "a_ser_fornecida_quantidade",
            "a_ser_fornecido_valor", "ainda_a_faturar_quantidade",
            "ainda_a_faturar_valor", "prz_apresent_cotacao", "origem",
            "rfq_coletiva", "num_itens", "codigo_controle", "reg_info_compras",
            "numero_pacote", "grupo_liberacao", "estrat_de_liberacao",
            "estado_liberacao", "codigo_liberacao", "dep_saida_do_estq",
            "nome_fornecedor", "unid_prec_pedido", "domicilio_fiscal",
            "valor_liquido_pedido", "incompleto", "qtd_notificada",
            "numero_hierarquia", "numero_externo_ordenacao",
            "ctg_externa_hierarquia", "urgencia_necessidade", "prior_neces",
            "numero_documento_configuravel", "nome_do_servico", "competencia",
            "arquivo_origem"
        ]
        seen = set()
        unique_cols = []
        for c in cols:
            if c not in seen:
                seen.add(c)
                unique_cols.append(c)
        chunks = []
        for r in recs:
            row = tuple(r.get(c) for c in unique_cols)
            chunks.append(row)
        insert_sql = (
            """
            INSERT INTO {s}.{t}
                ({columns})
            VALUES %s
            ON CONFLICT (documento_de_compras, item, contrato_basico) DO UPDATE SET
                {update_cols}
            """.format(
                s=SCHEMA,
                t=TABLE,
                columns=", ".join(unique_cols),
                update_cols=", ".join([f"{c} = EXCLUDED.{c}" for c in unique_cols])
            )
        )
        psycopg2.extras.execute_values(cur, insert_sql, chunks, page_size=1000)
        vinculo_sql = (
            """
            INSERT INTO {s}.{ta}
                (documento_de_compras, item, contrato_basico, arquivo_origem)
            VALUES %s
            ON CONFLICT DO NOTHING
            """.format(s=SCHEMA, ta=TABLE_ARQUIVO)
        )
        vinculos = [
            (r["documento_de_compras"], r["item"], r["contrato_basico"], orig)
            for r in recs
        ]
        if vinculos:
            psycopg2.extras.execute_values(cur, vinculo_sql, vinculos, page_size=1000)
        con.commit()

    cur.execute("SELECT COUNT(*) FROM {s}.{t}".format(s=SCHEMA, t=TABLE))
    total = cur.fetchone()[0]
    cur.close()
    con.close()
    return {"inserted": len(recs), "file": orig, "total": total}

def _to_cents(v):
    if v is None:
        return 0
    return int(round(float(v) * 100))


def _fmt_data(v):
    """Formata a coluna data_documento (TEXT) como DD/MM/AAAA.
    Aceita 'AAAA-MM-DD' (formato ISO gravado na importacao) e 'DD/MM/AAAA'."""
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%Y")
    s = str(v).strip()
    if not s:
        return ""
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return "%s/%s/%s" % (m.group(3), m.group(2), m.group(1))
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})", s)
    if m:
        return s
    return s


def fetch_arquivos():
    """Lista os arquivos (bases) que ja foram importados, com a contagem de POs.

    A contagem usa a tabela de vinculo, que guarda o historico: uma PO
    compartilhada por dois arquivos conta em ambos.
    """
    con = get_conn()
    cur = con.cursor()
    try:
        cur.execute(
            """
            SELECT v.arquivo_origem, COUNT(*) AS qtd, MAX(v.importado_em) AS ult
            FROM {s}.{t} v
            WHERE v.arquivo_origem IS NOT NULL AND trim(v.arquivo_origem) <> ''
            GROUP BY v.arquivo_origem
            ORDER BY ult DESC, v.arquivo_origem
            """.format(s=SCHEMA, t=TABLE_ARQUIVO)
        )
        return [{"arquivo": r[0], "count": r[1]} for r in cur.fetchall()]
    finally:
        cur.close()
        con.close()


def fetch_arquivos_por_po():
    """Mapa {doc|item|contrato} -> [lista de arquivos que contem a PO]."""
    con = get_conn()
    cur = con.cursor()
    try:
        cur.execute(
            """
            SELECT documento_de_compras, item, contrato_basico, arquivo_origem
            FROM {s}.{t}
            WHERE arquivo_origem IS NOT NULL AND trim(arquivo_origem) <> ''
            """.format(s=SCHEMA, t=TABLE_ARQUIVO)
        )
        mapa = {}
        for doc, item, contrato, arq in cur.fetchall():
            chave = "%s|%s|%s" % (doc or "", item or "", contrato or "")
            mapa.setdefault(chave, []).append(arq)
        return mapa
    finally:
        cur.close()
        con.close()




def fetch_pos(origem=None):
    """Carrega as POs do banco.

    Quando 'origem' e informada, devolve apenas as POs que pertencem a aquele
    arquivo (base) importado. As edicoes manuais (nome_do_servico /
    competencia) sao preservadas, porque vivem na tabela controle_po.
    """
    con = get_conn()
    cur = con.cursor()
    filtro = ""
    params = []
    if origem:
        filtro = (
            " AND EXISTS (SELECT 1 FROM {s}.{ta} v "
            "WHERE v.documento_de_compras = cp.documento_de_compras "
            "AND v.item = cp.item "
            "AND v.contrato_basico = cp.contrato_basico "
            "AND v.arquivo_origem = %s)"
        ).format(s=SCHEMA, ta=TABLE_ARQUIVO)
        params.append(origem)
    # Regra de eliminacao: traz somente as linhas em que codigo_eliminacao
    # NAO contem a letra L. Aplicado aqui, no SQL, para que o filtro de mes
    # (e qualquer outro recorte da tela) ja considere essa classificacao.
    # NULL e tratado como "sem L" (COALESCE) para nao perder linhas.
    filtro += " AND COALESCE(cp.codigo_eliminacao, '') NOT ILIKE '%%L%%'"
    cur.execute(
        """
        SELECT DISTINCT ON (cp.documento_de_compras, cp.item, cp.contrato_basico)
               cp.documento_de_compras, cp.item, cp.contrato_basico, cp.data_documento,
               cp.preco_liquido, cp.a_ser_fornecido_valor, cp.ainda_a_faturar_valor,
               cp.nome_do_servico, cp.competencia,
               aux.nome_do_projeto, cp.codigo_eliminacao
        FROM {s}.{t} cp
        LEFT JOIN {s}.controle_po_aux aux
            ON cp.contrato_basico = aux.contrato_basico
        WHERE substring(cp.data_documento FROM 1 FOR 4) = '2026' {filtro}
        ORDER BY cp.documento_de_compras, cp.item, cp.contrato_basico, aux.nome_do_projeto
        """.format(s=SCHEMA, t=TABLE, filtro=filtro),
        params,
    )
    rows = cur.fetchall()
    cur.close()
    con.close()

    mapa_arq = fetch_arquivos_por_po()
    pos = []
    for r in rows:
        chave = "%s|%s|%s" % (r[0] or "", r[1] or "", r[2] or "")
        pos.append(
            {
                "doc": r[0] or "",
                "item": r[1] or "",
                "contrato": r[2] or "",
                "data": _fmt_data(r[3]),
                "preco": _to_cents(r[4]),
                "aSerFornecido": _to_cents(r[5]),
                "aindaAFaturar": _to_cents(r[6]),
                "nomeServico": r[7] if r[7] is not None else "-",
                "competencia": r[8] if r[8] is not None else "-",
                "projeto": r[9] if r[9] is not None else "",
                # codigo de eliminacao do SAP: 'L' = PO cancelada,
                # 'S' = eliminacao parcial, '' = sem eliminacao
                "eliminacao": (r[10] or "").strip(),
                "arquivos": sorted(mapa_arq.get(chave, [])),
            }
        )
    return pos


def fetch_projetos():
    """
    Retorna lista de projetos com nome e contrato basico da tabela auxiliar.
    A contagem de POs vinculadas sera feita no frontend, analisando estado.pos.
    """
    con = get_conn()
    cur = con.cursor()
    cur.execute(
        """
        SELECT DISTINCT nome_do_projeto, contrato_basico
        FROM %s.controle_po_aux
        WHERE nome_do_projeto IS NOT NULL AND trim(nome_do_projeto) <> ''
        ORDER BY nome_do_projeto
        """ % SCHEMA
    )
    rows = cur.fetchall()
    cur.close()
    con.close()
    return [
        {"nome": r[0], "contrato": r[1] or ""}
        for r in rows if r[0]
    ]


def update_po(doc, item, contrato, nome_servico, data_inicio, data_fim):
    """Atualiza nome do servico e competencia (range) de uma PO.

    Persiste no banco. nome_servico vazio vira '-'.
    competencia vira 'inicio - fim' (ex: 01/10/2026 - 31/10/2026) ou '-' quando limpo.
    Retorna o numero de linhas atualizadas.
    """
    nome_servico = (nome_servico or "").strip() or "-"
    ini = (data_inicio or "").strip()
    fim = (data_fim or "").strip()
    if ini and fim:
        competencia = "{} - {}".format(ini, fim)
    else:
        competencia = "-"
    con = get_conn()
    cur = con.cursor()
    cur.execute(
        "UPDATE {s}.{t} SET nome_do_servico = %s, competencia = %s "
        "WHERE documento_de_compras = %s AND item = %s AND contrato_basico = %s".format(
            s=SCHEMA, t=TABLE
        ),
        (nome_servico, competencia, doc, item, contrato),
    )
    con.commit()
    updated = cur.rowcount
    cur.close()
    con.close()
    print("[update] doc=%s item=%s contrato=%s | competencia=%s | linhas atualizadas: %s" % (
        doc, item, contrato, competencia, updated))
    return updated


def find_latest_xlsx():
    files = [f for f in glob.glob(os.path.join(FILES_DIR, "*.xlsx"))
             if not os.path.basename(f).startswith("~$")]
    if not files:
        return None
    exp = [f for f in files if os.path.basename(f).upper().startswith("EXPORT")]
    pool = exp if exp else files
    return max(pool, key=os.path.getmtime)


class Handler(BaseHTTPRequestHandler):
    server_version = "ControlePO/1.0"

    def log_message(self, fmt, *args):
        sys.stderr.write("%s - [%s] %s\n" % (self.client_address[0], self.log_date_time_string(), fmt % args))

    def _json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _file(self, path, ctype):
        try:
            with open(path, "rb") as f:
                data = f.read()
        except FileNotFoundError:
            self.send_error(404, "Not Found")
            return
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        u = urlparse(self.path)
        if u.path == "/api/pos":
            try:
                origem = parse_qs(u.query).get("origem", [None])[0]
                pos = fetch_pos(origem if origem else None)
                self._json({"ok": True, "count": len(pos), "pos": pos})
            except Exception as e:
                self._json({"ok": False, "error": "%s: %s" % (type(e).__name__, e)}, 500)
        elif u.path == "/api/arquivos":
            try:
                arquivos = fetch_arquivos()
                self._json({"ok": True, "arquivos": arquivos})
            except Exception as e:
                self._json({"ok": False, "error": "%s: %s" % (type(e).__name__, e)}, 500)
        elif u.path == "/api/info":
            try:
                con = get_conn()
                cur = con.cursor()
                cur.execute("SELECT COUNT(*) FROM {s}.{t}".format(s=SCHEMA, t=TABLE))
                total = cur.fetchone()[0]
                # mes vigente: ultimo mes com dados na coluna data_documento
                # coluna eh TEXT (ex: '2026-09-09'), entao faz cast ::date antes do to_char
                cur.execute(
                    "SELECT to_char(MAX(data_documento::date), 'MM/YYYY') "
                    "FROM {s}.{t} WHERE data_documento IS NOT NULL".format(
                        s=SCHEMA, t=TABLE
                    )
                )
                mes_vigente_row = cur.fetchone()
                mes_vigente = mes_vigente_row[0] if mes_vigente_row and mes_vigente_row[0] else None
                cur.close()
                con.close()
                self._json({
                    "ok": True,
                    "host": DB["host"],
                    "port": DB["port"],
                    "database": DB["dbname"],
                    "schema": SCHEMA,
                    "table": TABLE,
                    "count": total,
                    "mes_vigente": mes_vigente,
                })
            except Exception as e:
                self._json({"ok": False, "error": "%s: %s" % (type(e).__name__, e)}, 500)
        elif u.path == "/api/projetos":
            try:
                projetos = fetch_projetos()
                self._json({"ok": True, "projetos": projetos})
            except Exception as e:
                self._json({"ok": False, "error": "%s: %s" % (type(e).__name__, e)}, 500)
        elif u.path in ("/", "/index.html"):
            self._file(os.path.join(BASE_DIR, "index.html"), "text/html; charset=utf-8")
        elif u.path == "/favicon.ico":
            self.send_response(204)
            self.end_headers()
        else:
            self.send_error(404, "Not Found")

    def do_POST(self):
        u = urlparse(self.path)
        length = int(self.headers.get("Content-Length", 0) or 0)
        body = self.rfile.read(length) if length else b""
        if u.path == "/api/import":
            try:
                # Tenta ler arquivo enviado via corpo JSON (base64) ou por query ?file=
                qs = parse_qs(u.query)
                fname = qs.get("file", [None])[0]
                saved_name = None
                if body:
                    try:
                        payload = json.loads(body.decode("utf-8"))
                        data_b64 = payload.get("data_base64") or payload.get("data")
                        fname = payload.get("filename") or fname
                        if data_b64 and fname:
                            os.makedirs(FILES_DIR, exist_ok=True)
                            file_data = base64.b64decode(data_b64)
                            saved_path = os.path.join(FILES_DIR, fname)
                            with open(saved_path, "wb") as f:
                                f.write(file_data)
                            saved_name = fname
                    except Exception:
                        pass
                target = saved_name or fname or find_latest_xlsx()
                if not target:
                    self._json({"ok": False, "error": "Nenhum arquivo .xlsx encontrado na pasta Files."}, 404)
                    return
                full = target if os.path.isabs(target) else os.path.join(FILES_DIR, target)
                if not os.path.exists(full):
                    self._json({"ok": False, "error": "Arquivo nao encontrado: %s" % full}, 404)
                    return
                res = import_file(full)
                pos = fetch_pos()
                res.update({"ok": True, "pos": pos, "count": len(pos)})
                self._json(res)
            except Exception as e:
                import traceback
                traceback.print_exc()
                self._json({"ok": False, "error": "%s: %s" % (type(e).__name__, e)}, 500)
        elif u.path == "/api/clear":
            try:
                con = get_conn()
                cur = con.cursor()
                cur.execute("TRUNCATE TABLE {s}.{t}".format(s=SCHEMA, t=TABLE))
                cur.execute("TRUNCATE TABLE {s}.{ta}".format(s=SCHEMA, ta=TABLE_ARQUIVO))
                con.commit()
                cur.close()
                con.close()
                self._json({"ok": True, "cleared": True})
            except Exception as e:
                self._json({"ok": False, "error": "%s: %s" % (type(e).__name__, e)}, 500)
        elif u.path == "/api/update":
            try:
                payload = json.loads(body.decode("utf-8")) if body else {}
                doc = (payload.get("doc") or "").strip()
                if not doc:
                    self._json({"ok": False, "error": "Documento não informado."}, 400)
                    return
                updated = update_po(
                    doc,
                    (payload.get("item") or "").strip(),
                    (payload.get("contrato") or "").strip(),
                    payload.get("nome_servico"),
                    payload.get("data_inicio"),
                    payload.get("data_fim"),
                )
                pos = fetch_pos()
                self._json({"ok": True, "updated": updated, "pos": pos, "count": len(pos)})
            except Exception as e:
                import traceback
                traceback.print_exc()
                self._json({"ok": False, "error": "%s: %s" % (type(e).__name__, e)}, 500)
        elif u.path == "/api/projetos":
            try:
                payload = json.loads(body.decode("utf-8")) if body else {}
                nome = (payload.get("nome") or "").strip()
                contrato = (payload.get("contrato") or "").strip()
                if not nome or not contrato:
                    self._json({"ok": False, "error": "Nome do projeto e contrato são obrigatórios."}, 400)
                    return
                con = get_conn()
                cur = con.cursor()
                cur.execute(
                    """
                    INSERT INTO {s}.controle_po_aux (nome_do_projeto, contrato_basico)
                    VALUES (%s, %s)
                    """.format(s=SCHEMA),
                    (nome, contrato),
                )
                con.commit()
                cur.close()
                con.close()
                self._json({"ok": True, "projeto": {"nome": nome, "contrato": contrato}})
            except Exception as e:
                import traceback
                traceback.print_exc()
                self._json({"ok": False, "error": "%s: %s" % (type(e).__name__, e)}, 500)
        else:
            self.send_error(404, "Not Found")

    def do_DELETE(self):
        u = urlparse(self.path)
        length = int(self.headers.get("Content-Length", 0) or 0)
        body = self.rfile.read(length) if length else b""
        if u.path == "/api/projetos":
            try:
                payload = json.loads(body.decode("utf-8")) if body else {}
                nome = (payload.get("nome") or "").strip()
                contrato = (payload.get("contrato") or "").strip()
                if not nome or not contrato:
                    self._json({"ok": False, "error": "Nome e contrato são obrigatórios."}, 400)
                    return
                con = get_conn()
                cur = con.cursor()
                cur.execute(
                    """
                    DELETE FROM {s}.controle_po_aux
                    WHERE nome_do_projeto = %s AND contrato_basico = %s
                    """.format(s=SCHEMA),
                    (nome, contrato),
                )
                deleted = cur.rowcount
                con.commit()
                cur.close()
                con.close()
                if deleted == 0:
                    self._json({"ok": False, "error": "Projeto não encontrado na tabela auxiliar."}, 404)
                    return
                self._json({"ok": True, "deleted": deleted, "projeto": {"nome": nome, "contrato": contrato}})
            except Exception as e:
                import traceback
                traceback.print_exc()
                self._json({"ok": False, "error": "%s: %s" % (type(e).__name__, e)}, 500)
        else:
            self.send_error(404, "Not Found")


def main():
    port = int(os.environ.get("PORT", "8000"))
    try:
        os.makedirs(FILES_DIR, exist_ok=True)
        ensure_table()
        print("[ok] schema/tabela prontos: %s.%s" % (SCHEMA, TABLE))
    except Exception as e:
        print("[aviso] nao foi possivel preparar o banco agora: %s" % e)
    srv = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print("Servidor em http://localhost:%d  (Ctrl+C para parar)" % port)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        srv.server_close()


if __name__ == "__main__":
    main()
