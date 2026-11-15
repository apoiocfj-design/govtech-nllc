"""
GovTech NLLC - Motor de Importação CNBS (CATMAT / CATSER) e Parser de Palavras-Chave
====================================================================================
Importa os arquivos oficiais `catser.csv` (anexado pelo usuário) e `catmat.csv`
(Repositório CNBS / Consulta Detalhada Compras.gov.br) para um banco SQLite FTS5
de alta velocidade, permitindo:
  1. Busca instantânea por palavras-chave (sem itens pré-carregados de nenhum órgão).
  2. Filtro por Tipo (MATERIAL = CATMAT, SERVICO = CATSER, ou TODOS) e Status Ativo.
  3. Parser inteligente de texto/demanda livre (extrai itens de listas coladas do DFD/TR
     e sugere automaticamente o código CATMAT/CATSER correspondente).
"""

from __future__ import annotations

import csv
import os
import re
import sqlite3
import time
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional

BASE_DIR = Path(__file__).resolve().parent.parent.parent
DB_DIR = BASE_DIR / "database"
CATSER_CSV_PATH = DB_DIR / "catser.csv"
CATMAT_CSV_PATH = DB_DIR / "catmat.csv"
SQLITE_DB_PATH = DB_DIR / "cnbs_catalog.db"

STOPWORDS_PT = {
    "a", "as", "o", "os", "de", "da", "do", "das", "dos", "em", "no", "na",
    "nos", "nas", "para", "por", "com", "sem", "sob", "sobre", "um", "uma",
    "uns", "umas", "e", "ou", "tipo", "modelo", "marca", "uso", "geral",
    "item", "itens", "material", "servico", "servicos", "un", "und", "unid",
    "unidade", "unidades", "par", "pares", "cx", "caixa", "pct", "pacote",
}

# Equivalências técnicas comuns no vocabulário do CNBS / Compras.gov.br
SYNONYM_EQUIVALENTS: Dict[str, List[str]] = {
    "ETILOMETRO": ["ETILOMETRO", "BAFOMETRO"],
    "BAFOMETRO": ["BAFOMETRO", "ETILOMETRO"],
    "NOBREAK": ["NOBREAK", "ININTERRUPTA"],
    "PENDRIVE": ["PENDRIVE", "FLASH"],
    "NOTEBOOK": ["NOTEBOOK", "PORTATIL"],
    "LAPTOP": ["NOTEBOOK", "PORTATIL"],
}


def strip_accents(text: str) -> str:
    """Remove acentos e converte para maiúsculas para indexação e busca FTS5."""
    if not text:
        return ""
    nfkd = unicodedata.normalize("NFKD", str(text))
    return "".join(c for c in nfkd if not unicodedata.combining(c)).upper()


def tokenize_keywords(query: str) -> List[str]:
    """Extrai palavras-chave significativas de uma frase ou descrição de item."""
    normalized = strip_accents(query)
    raw_tokens = re.findall(r"[A-Z0-9]{2,}", normalized)
    tokens = [t for t in raw_tokens if t.lower() not in STOPWORDS_PT]
    return tokens if tokens else raw_tokens


class CNBSCatalogEngine:
    """
    Gerencia o banco local SQLite FTS5 alimentado pelas planilhas oficiais
    catser.csv e catmat.csv do Catálogo Nacional de Bens e Serviços (CNBS).
    """

    def __init__(self, db_path: Path = SQLITE_DB_PATH) -> None:
        self.db_path = db_path
        self.ensure_database_built()

    def ensure_database_built(self, force_rebuild: bool = False) -> Dict[str, Any]:
        """Cria e popula o banco SQLite FTS5 caso ainda não exista."""
        gz_path = self.db_path.parent / (self.db_path.name + ".gz")
        if not self.db_path.exists() and gz_path.exists():
            try:
                import gzip, shutil
                with gzip.open(gz_path, "rb") as f_in, open(self.db_path, "wb") as f_out:
                    shutil.copyfileobj(f_in, f_out)
            except Exception as e:
                print(f"Erro ao descompactar banco gz: {e}")

        if self.db_path.exists() and not force_rebuild:
            try:
                conn = sqlite3.connect(str(self.db_path))
                cur = conn.cursor()
                cur.execute("SELECT COUNT(*) FROM catalog_items")
                total = cur.fetchone()[0]
                conn.close()
                if total > 5000:
                    return {"status": "ready", "total_items": total}
            except Exception:
                pass

        return self.build_database()

    def build_database(self) -> Dict[str, Any]:
        """Importa catser.csv e catmat.csv (ativos) para SQLite + FTS5."""
        t0 = time.time()
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        if self.db_path.exists():
            os.remove(self.db_path)

        conn = sqlite3.connect(str(self.db_path))
        cur = conn.cursor()
        cur.execute("PRAGMA synchronous = OFF")
        cur.execute("PRAGMA journal_mode = MEMORY")

        cur.execute(
            """
            CREATE TABLE catalog_items (
                item_type TEXT NOT NULL,
                codigo INTEGER NOT NULL,
                nome_pdm_ou_classe TEXT,
                descricao TEXT NOT NULL,
                nome_grupo TEXT,
                nome_classe TEXT,
                ncm TEXT,
                ativo INTEGER NOT NULL DEFAULT 1,
                PRIMARY KEY (item_type, codigo)
            )
            """
        )

        cur.execute(
            """
            CREATE VIRTUAL TABLE catalog_fts USING fts5(
                item_type UNINDEXED,
                codigo UNINDEXED,
                texto_busca,
                tokenize = 'unicode61 remove_diacritics 2'
            )
            """
        )

        catser_count = 0
        catmat_count = 0

        # 1. Importar CATSER (catser.csv anexado pelo usuário)
        if CATSER_CSV_PATH.exists():
            with open(CATSER_CSV_PATH, mode="r", encoding="utf-8-sig", errors="replace") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader, None)
                batch_items = []
                batch_fts = []
                for row in reader:
                    if len(row) < 7:
                        continue
                    cod_grupo, nome_grupo, cod_classe, nome_classe, cod_serv, nome_serv, status_serv = row[:7]
                    ativo = 1 if status_serv.strip().lower() == "true" else 0
                    if not ativo:
                        continue
                    try:
                        codigo_int = int(cod_serv.strip())
                    except ValueError:
                        continue
                    desc_clean = " ".join(nome_serv.split())
                    classe_clean = " ".join(nome_classe.split())
                    grupo_clean = " ".join(nome_grupo.split())
                    search_blob = f"{codigo_int} {desc_clean} {classe_clean} {grupo_clean}"

                    batch_items.append(
                        ("SERVICO", codigo_int, classe_clean, desc_clean, grupo_clean, classe_clean, "", ativo)
                    )
                    batch_fts.append(("SERVICO", codigo_int, search_blob))
                    catser_count += 1

                cur.executemany(
                    "INSERT OR REPLACE INTO catalog_items VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    batch_items,
                )
                cur.executemany(
                    "INSERT INTO catalog_fts (item_type, codigo, texto_busca) VALUES (?, ?, ?)",
                    batch_fts,
                )
                conn.commit()

        # 2. Importar CATMAT (catmat.csv oficial do Repositório CNBS)
        if CATMAT_CSV_PATH.exists():
            with open(CATMAT_CSV_PATH, mode="r", encoding="utf-8-sig", errors="replace") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader, None)
                batch_items = []
                batch_fts = []
                for row in reader:
                    if len(row) < 12:
                        continue
                    status_item = row[11].strip()
                    if status_item != "True":
                        continue
                    try:
                        codigo_int = int(row[6].strip())
                    except ValueError:
                        continue
                    nome_grupo = row[1].strip()
                    nome_classe = row[3].strip()
                    nome_pdm = row[5].strip()
                    descricao = row[7].strip()
                    ncm = row[8].strip()
                    search_blob = f"{codigo_int} {nome_pdm} {descricao} {nome_classe}"

                    batch_items.append(
                        ("MATERIAL", codigo_int, nome_pdm, descricao, nome_grupo, nome_classe, ncm, 1)
                    )
                    batch_fts.append(("MATERIAL", codigo_int, search_blob))
                    catmat_count += 1

                    if len(batch_items) >= 25000:
                        cur.executemany(
                            "INSERT OR REPLACE INTO catalog_items VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                            batch_items,
                        )
                        cur.executemany(
                            "INSERT INTO catalog_fts (item_type, codigo, texto_busca) VALUES (?, ?, ?)",
                            batch_fts,
                        )
                        conn.commit()
                        batch_items.clear()
                        batch_fts.clear()

                if batch_items:
                    cur.executemany(
                        "INSERT OR REPLACE INTO catalog_items VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                        batch_items,
                    )
                    cur.executemany(
                        "INSERT INTO catalog_fts (item_type, codigo, texto_busca) VALUES (?, ?, ?)",
                        batch_fts,
                    )
                    conn.commit()

        cur.execute("CREATE INDEX IF NOT EXISTS idx_catalog_codigo ON catalog_items(codigo)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_catalog_type ON catalog_items(item_type)")
        conn.commit()
        conn.close()

        elapsed = round(time.time() - t0, 2)
        return {
            "status": "built",
            "catser_active_imported": catser_count,
            "catmat_active_imported": catmat_count,
            "total_items": catser_count + catmat_count,
            "elapsed_seconds": elapsed,
        }

    def get_stats(self) -> Dict[str, Any]:
        """Retorna estatísticas dos catálogos CATMAT e CATSER carregados."""
        conn = sqlite3.connect(str(self.db_path))
        cur = conn.cursor()
        cur.execute("SELECT item_type, COUNT(*) FROM catalog_items GROUP BY item_type")
        counts = dict(cur.fetchall())
        conn.close()
        return {
            "catmat_ativos": counts.get("MATERIAL", 0),
            "catser_ativos": counts.get("SERVICO", 0),
            "total_ativos": sum(counts.values()),
            "fonte_oficial": "Repositório CNBS Compras.gov.br (catmat.csv + catser.csv)",
        }

    def search(
        self,
        query: str,
        item_type: str = "ALL",
        limit: int = 25,
    ) -> Dict[str, Any]:
        """
        Realiza o parsing de palavras-chave e pesquisa na base unificada CATMAT + CATSER.
        Aceita tanto códigos numéricos quanto termos livres (ex: 'bocal etilometro',
        'manutencao ar condicionado', 'notebook i7', 'extintor co2').
        """
        q_clean = (query or "").strip()
        if not q_clean:
            return {"query": query, "tokens": [], "total_found": 0, "results": []}

        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        item_type_upper = (item_type or "ALL").upper()
        results: List[Dict[str, Any]] = []

        # Caso 1: Busca direta por código numérico exato
        if q_clean.isdigit():
            code_val = int(q_clean)
            if item_type_upper in ("MATERIAL", "SERVICO"):
                cur.execute(
                    "SELECT * FROM catalog_items WHERE codigo = ? AND item_type = ? LIMIT ?",
                    (code_val, item_type_upper, limit),
                )
            else:
                cur.execute(
                    "SELECT * FROM catalog_items WHERE codigo = ? LIMIT ?",
                    (code_val, limit),
                )
            for r in cur.fetchall():
                results.append(dict(r))

        # Caso 2: Parser de palavras-chave com FTS5 (AND progressivo com equivalências + fallback OR)
        tokens = tokenize_keywords(q_clean)
        if tokens and len(results) < limit:
            and_clauses = []
            for t in tokens:
                syns = SYNONYM_EQUIVALENTS.get(t, [t])
                if len(syns) == 1:
                    and_clauses.append(f'"{t}"*')
                else:
                    or_group = " OR ".join(f'"{s}"*' for s in syns)
                    and_clauses.append(f"({or_group})")
            fts_and_query = " AND ".join(and_clauses)
            results.extend(self._execute_fts_query(cur, fts_and_query, item_type_upper, limit, tokens))

            # Se retornou poucos itens e há múltiplos tokens, complementa com combinação parcial
            if len(results) < 5 and len(tokens) > 1:
                seen_keys = {(r["item_type"], r["codigo"]) for r in results}
                fts_or_query = " OR ".join(and_clauses[:5])
                extra = self._execute_fts_query(cur, fts_or_query, item_type_upper, limit, tokens)
                for item in extra:
                    key = (item["item_type"], item["codigo"])
                    if key not in seen_keys:
                        results.append(item)
                        seen_keys.add(key)
                        if len(results) >= limit:
                            break

        conn.close()
        return {
            "query": q_clean,
            "tokens_parsed": tokens,
            "item_type_filter": item_type_upper,
            "total_found": len(results),
            "results": results[:limit],
        }

    def _execute_fts_query(
        self,
        cur: sqlite3.Cursor,
        fts_expr: str,
        item_type_upper: str,
        limit: int,
        tokens: List[str],
    ) -> List[Dict[str, Any]]:
        base_sql = """
            SELECT
                c.item_type,
                c.codigo,
                c.nome_pdm_ou_classe,
                c.descricao,
                c.nome_grupo,
                c.nome_classe,
                c.ncm,
                bm25(catalog_fts) AS rank_score
            FROM catalog_fts f
            JOIN catalog_items c
              ON c.item_type = f.item_type AND c.codigo = CAST(f.codigo AS INTEGER)
            WHERE catalog_fts MATCH ?
        """
        params: List[Any] = [fts_expr]
        if item_type_upper in ("MATERIAL", "SERVICO"):
            base_sql += " AND c.item_type = ?"
            params.append(item_type_upper)

        rows: List[Dict[str, Any]] = []
        try:
            # 1º Passo: Busca códigos modernos na faixa de alta liquidez do Compras.gov.br (450000..625000) e CATSER
            modern_sql = (
                base_sql
                + " AND (c.item_type = 'SERVICO' OR (c.codigo >= 450000 AND c.codigo <= 625000))"
                + " ORDER BY c.codigo DESC LIMIT ?"
            )
            cur.execute(modern_sql, [*params, max(limit * 8, 160)])
            rows = [dict(r) for r in cur.fetchall()]

            if len(rows) < limit:
                seen = {(r["item_type"], r["codigo"]) for r in rows}
                fallback_sql = base_sql + " ORDER BY rank_score LIMIT ?"
                cur.execute(fallback_sql, [*params, max(limit * 6, 120)])
                for r in cur.fetchall():
                    d = dict(r)
                    k = (d["item_type"], d["codigo"])
                    if k not in seen:
                        rows.append(d)
                        seen.add(k)
        except sqlite3.OperationalError:
            return []

        first_tok = tokens[0] if tokens else ""

        def token_matches_text(tok: str, text_norm: str) -> bool:
            for s in SYNONYM_EQUIVALENTS.get(tok, [tok]):
                if s in text_norm:
                    return True
            return False

        def score_row(item: Dict[str, Any]) -> tuple:
            pdm_norm = strip_accents(item.get("nome_pdm_ou_classe", ""))
            desc_norm = strip_accents(item.get("descricao", ""))
            code_num = int(item.get("codigo") or 0)
            is_prime_modern = 1 if (item.get("item_type") == "SERVICO" or 450000 <= code_num <= 622000) else 0
            matched_tokens = sum(1 for t in tokens if token_matches_text(t, desc_norm) or token_matches_text(t, pdm_norm))
            pdm_starts_first = 1 if pdm_norm.startswith(first_tok) else 0
            desc_starts_first = 1 if desc_norm.startswith(first_tok) else 0
            pdm_words = pdm_norm.split()
            first_is_main_noun = 1 if (pdm_words and pdm_words[0] == first_tok) else 0
            return (
                -matched_tokens,
                -is_prime_modern,
                -first_is_main_noun,
                -pdm_starts_first,
                -desc_starts_first,
                item.get("rank_score", 0.0),
            )

        rows.sort(key=score_row)
        for r in rows:
            r.pop("rank_score", None)
        return rows[:limit]

    def parse_free_demand_text(self, raw_text: str) -> List[Dict[str, Any]]:
        """
        Recebe um texto livre colado pelo usuário (ex: lista de itens de um DFD/TR)
        e extrai automaticamente cada linha, identificando palavras-chave, unidade,
        quantidade estimada e sugerindo os códigos CATMAT/CATSER correspondentes.
        """
        lines = [ln.strip() for ln in (raw_text or "").splitlines() if len(ln.strip()) >= 3]
        parsed_items: List[Dict[str, Any]] = []

        for idx, line in enumerate(lines, start=1):
            # Extrai quantidade se houver (ex: "800 UN - Bocal..." ou "15 pares Luva...")
            qty_match = re.search(
                r"\b(\d+(?:[\.,]\d+)?)\s*(UN|UND|UNID|PAR|PARES|CX|CAIXA|PCT|KG|M|L|SERV|MES|MESES)?\b",
                line,
                re.IGNORECASE,
            )
            qty = 1
            unit = "UN"
            cleaned_desc = line
            if qty_match:
                try:
                    qty = max(1, int(float(qty_match.group(1).replace(",", "."))))
                except ValueError:
                    qty = 1
                if qty_match.group(2):
                    unit = qty_match.group(2).upper()
                # Remove numeração inicial de lista tipo "1." ou "Item 01 -"
                cleaned_desc = re.sub(r"^(?:item\s*)?\d+[\.\)\-\s]+", "", line, flags=re.IGNORECASE).strip()

            search_res = self.search(cleaned_desc, item_type="ALL", limit=5)
            parsed_items.append(
                {
                    "line_number": idx,
                    "original_text": line,
                    "extracted_keywords": search_res.get("tokens_parsed", []),
                    "suggested_quantity": qty,
                    "suggested_unit": unit,
                    "matches": search_res.get("results", []),
                }
            )

        return parsed_items


# Links Oficiais Extraídos da Página de Consulta Detalhada do Compras.gov.br
CONSULTA_DETALHADA_LINKS: List[Dict[str, Any]] = [
    {
        "categoria": "Catálogos e Dados Abertos (CNBS / APIs)",
        "itens": [
            {
                "titulo": "Busca Web Oficial CNBS (CATMAT / CATSER)",
                "descricao": "Interface oficial do Compras.gov.br para navegação hierárquica de materiais e serviços.",
                "url": "https://catalogo.compras.gov.br/cnbs-web/busca",
            },
            {
                "titulo": "Repositório Oficial de Planilhas CSV (CATMAT e CATSER)",
                "descricao": "Download direto de catmat.csv e catser.csv atualizados semanalmente (sábados às 6h).",
                "url": "https://repositorio.dados.gov.br/seges/comprasgov/catalogo_cnbs/",
            },
            {
                "titulo": "API de Dados Abertos Compras.gov.br (Swagger OAS 3.1)",
                "descricao": "Documentação interativa dos endpoints de Pesquisa de Preços, Material, Serviço, UASG e Contratos.",
                "url": "https://dadosabertos.compras.gov.br/swagger-ui/index.html",
            },
            {
                "titulo": "Portal Nacional de Contratações Públicas (PNCP)",
                "descricao": "Consulta oficial de editais, avisos de dispensa, atas de registro de preços e contratos na Lei 14.133/2021.",
                "url": "https://www.gov.br/pncp/",
            },
            {
                "titulo": "Painel de Compras da Administração Pública Federal",
                "descricao": "Painel BI de inteligência e estatísticas de gastos públicos federais.",
                "url": "http://paineldecompras.economia.gov.br",
            },
        ],
    },
    {
        "categoria": "Habilitação de Fornecedores e Certidões (SICAF / TCU / RFB)",
        "itens": [
            {
                "titulo": "SICAF - Consultar CRC de Fornecedor",
                "descricao": "Emissão e verificação do Certificado de Registro Cadastral no SICAF.",
                "url": "https://www3.comprasnet.gov.br/sicaf-web/public/pages/consultas/consultarCRC.jsf",
            },
            {
                "titulo": "SICAF - Restrição de Contratar com a Administração Pública",
                "descricao": "Consulta pública de impedimentos de licitar, suspensões e inidoneidades no SICAF.",
                "url": "https://www3.comprasnet.gov.br/sicaf-web/public/pages/consultas/consultarRestricaoContratarAdministracaoPublica.jsf",
            },
            {
                "titulo": "SICAF - Linha de Fornecimento por Material/Serviço",
                "descricao": "Localiza empresas e fornecedores cadastrados no SICAF que fornecem determinado item ou classe.",
                "url": "https://www3.comprasnet.gov.br/sicaf-web/public/pages/consultas/consultarLinhaFornecimento.jsf",
            },
            {
                "titulo": "Certidão de Regularidade Fiscal RFB / PGFN (Receita Federal)",
                "descricao": "Emissão de Certidão Negativa/Positiva com Efeitos de Negativa de Débitos Federais e Dívida Ativa.",
                "url": "https://servicos.receita.fazenda.gov.br/Servicos/certidao/CNDConjuntaInter/InformaNICertidao.asp?tipo=2",
            },
            {
                "titulo": "Certidão Negativa de Débitos Trabalhistas (CNDT - TST)",
                "descricao": "Verificação obrigatória de regularidade trabalhista perante o Tribunal Superior do Trabalho.",
                "url": "https://www.tst.jus.br/certidao",
            },
            {
                "titulo": "Certificado de Regularidade do FGTS (CRF - Caixa)",
                "descricao": "Consulta de regularidade do empregador perante o Fundo de Garantia do Tempo de Serviço.",
                "url": "https://consulta-crf.caixa.gov.br/consultacrf/pages/consultaEmpregador.jsf",
            },
        ],
    },
    {
        "categoria": "Consultas de Licitações, Pregões, Atas SRP e Contratos",
        "itens": [
            {
                "titulo": "Gestão e Consulta de Atas de Registro de Preços (SRP)",
                "descricao": "Consulta de Atas de Registro de Preços vigentes para carona/adesão ou pesquisa de preços.",
                "url": "https://www2.comprasnet.gov.br/siasgnet-atasrp/public/principal.do",
            },
            {
                "titulo": "Intenção de Registro de Preços (IRP)",
                "descricao": "Consulta de IRPs abertas para participação de órgãos interessados.",
                "url": "https://www2.comprasnet.gov.br/siasgnet-irp/consultarIRPComprasNet.do?method=iniciar",
            },
            {
                "titulo": "Acompanhamento Público de Compras / Pregões (Comprasnet Web)",
                "descricao": "Consulta detalhada de sessões públicas, propostas, lances e atas na Lei 14.133/2021.",
                "url": "https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras",
            },
            {
                "titulo": "Pesquisa Textual de Editais e Avisos de Licitação",
                "descricao": "Busca textual dentro de editais publicados no sistema Comprasnet.",
                "url": "http://comprasnet.gov.br/ConsultaLicitacoes/ConsLicitacao_texto.asp",
            },
            {
                "titulo": "Consulta de UASGs (Unidades Administrativas de Serviços Gerais)",
                "descricao": "Catálogo de códigos UASG e órgãos vinculados ao SISG.",
                "url": "http://comprasnet.gov.br/livre/uasg/index.htm",
            },
            {
                "titulo": "Caderno de Logística - Pesquisa de Preços (SEGES)",
                "descricao": "Manual técnico oficial sobre aplicação prática da IN SEGES/ME nº 65/2021.",
                "url": "https://www.gov.br/compras/pt-br/agente-publico/cadernos-de-logistica/midia/temporario-caderno-de-pesquisa-de-precos-2023_final-1.pdf",
            },
        ],
    },
]


if __name__ == "__main__":
    engine = CNBSCatalogEngine()
    stats = engine.get_stats()
    print("CNBS Stats:", stats)
    for termo in ["bocal etilometro", "manutencao ar condicionado", "extintor co2", "software java"]:
        res = engine.search(termo, limit=3)
        print(f"\nBusca: '{termo}' (Tokens: {res['tokens_parsed']}) -> {res['total_found']} resultados")
        for r in res["results"]:
            print(f"  [{r['item_type']} {r['codigo']}] {r['nome_pdm_ou_classe']} | {r['descricao'][:90]}")
