# -*- coding: utf-8 -*-
"""
Testes Unitários da Camada de Consolidação em Lote, Nota Técnica AGU, Mapa Geral
com Todas as Empresas, Fallback 12M->24M e Pacote ZIP Condensado por Categoria (IN 65/2021).
"""

import io
import os
import sys
import time
import unittest
import zipfile
from unittest.mock import MagicMock

import docx
import openpyxl

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "core")))

from core.agu_batch_reports import (
    brl_por_extenso,
    build_agu_nota_tecnica_docx_bytes,
    build_agu_nota_tecnica_pdf_bytes,
)
from core.gov_api_client import PriceResearchController
from core.job_manager import BatchJobManager
from core.report_generators import (
    build_batch_zip_bytes,
    build_batch_zip_to_file,
    build_consolidated_in65_docx_bytes,
    build_consolidated_in65_excel_bytes,
    build_consolidated_in65_pdf_bytes,
    build_consolidated_suppliers_dossier_pdf_bytes,
)
from core.stat_engine import RawQuoteInput


def _make_sample_item(i: int, is_24m: bool = False, cat_name: str = "MATERIAL DE EXPEDIENTE") -> dict:
    return {
        "item_index": i,
        "code": 600000 + i,
        "type": "MATERIAL",
        "unit": "UN",
        "qty": 5,
        "search_window_months": 24 if is_24m else 12,
        "is_extended_24m": is_24m,
        "report": {
            "catmat_catser_code": 600000 + i,
            "item_description": f"MATERIAL DE CONSUMO ITEM TESTE {i}",
            "target_supply_unit": "UN",
            "nome_classe": cat_name,
            "nome_pdm": f"PDM {cat_name}",
            "search_window_months": 24 if is_24m else 12,
            "is_extended_24m": is_24m,
            "final_estimated_unit_price": 50.0 + i,
            "final_metric_used": "MEDIA" if i % 2 == 0 else "MEDIANA",
            "adopted_parameter": "MEDIA" if i % 2 == 0 else "MEDIANA",
            "is_homogeneous": (i % 2 == 0),
            "valid_quotes": [
                {
                    "unit_price": 50.0 + i,
                    "supplier_cnpj": f"12.345.678/0001-{(i % 89) + 10}",
                    "supplier_name": f"FORNECEDOR ALFA {i} LTDA",
                    "buyer_uasg": "782340",
                    "buyer_org_name": "COMANDO DA MARINHA",
                    "buyer_uf": "RJ",
                    "esfera": "Federal",
                    "brand": "MARCA PRO",
                    "supply_unit": "UN",
                    "quantity_homologated": 100,
                    "purchase_date": "2025-02-10" if is_24m else "2026-08-10",
                    "purchase_number_year": f"000{i:02d}/2026",
                    "pncp_or_compras_uri": f"https://pncp.gov.br/app/editais/00394502000144/2026/{i:06d}",
                    "source_inciso": "Inciso I - Compras.gov.br",
                    "is_extended_24m": is_24m,
                },
                {
                    "unit_price": 52.0 + i,
                    "supplier_cnpj": f"98.765.432/0001-{(i % 89) + 10}",
                    "supplier_name": f"DISTRIBUIDORA BETA {i} S/A",
                    "buyer_uasg": "160518",
                    "buyer_org_name": "COMANDO DO EXERCITO",
                    "buyer_uf": "SP",
                    "esfera": "Federal",
                    "brand": "MARCA PLUS",
                    "supply_unit": "UN",
                    "quantity_homologated": 50,
                    "purchase_date": "2025-03-12" if is_24m else "2026-09-12",
                    "purchase_number_year": f"001{i:02d}/2026",
                    "pncp_or_compras_uri": f"https://pncp.gov.br/app/editais/00394452000103/2026/{i:06d}",
                    "source_inciso": "Inciso I - Compras.gov.br",
                    "is_extended_24m": is_24m,
                },
            ],
            "purged_quotes": [
                {
                    "unit_price": 5000.0 + i,
                    "supplier_cnpj": "11.111.111/0001-11",
                    "supplier_name": "EMPRESA OUTLIER LTDA",
                    "buyer_uasg": "999999",
                    "buyer_org_name": "ORGAO TESTE",
                    "buyer_uf": "MG",
                    "supply_unit": "UN",
                    "purchase_date": "2026-05-01",
                    "purchase_number_year": "99999/2026",
                    "purge_status": "EXPURGADA_ACIMA_LIMITE",
                    "purge_justification": "Sobrepreço acima de 400% da mediana bruta.",
                }
            ],
            "sanitized_metrics": {
                "sample_size": 2,
                "min_price": 50.0 + i,
                "max_price": 52.0 + i,
                "mean": 51.0 + i,
                "median": 51.0 + i,
                "cv_percent": 2.8 if i % 2 == 0 else 28.5,
                "std_deviation": 1.41,
            },
        },
    }


class TestBatchReports(unittest.TestCase):

    def setUp(self):
        cats = ["MATERIAL DE EXPEDIENTE", "EQUIPAMENTOS DE INFORMATICA", "MATERIAL HOSPITALAR"]
        self.sample_items = [
            _make_sample_item(i, is_24m=(i > 16), cat_name=cats[i % len(cats)])
            for i in range(1, 22)
        ]

    def test_brl_por_extenso(self):
        self.assertEqual(brl_por_extenso(0), "zero reais")
        self.assertEqual(brl_por_extenso(1.0), "um real")
        self.assertEqual(brl_por_extenso(1250.50), "mil, duzentos e cinquenta reais e cinquenta centavos")
        self.assertEqual(brl_por_extenso(1000000.00), "um milhão de reais")

    def test_consolidated_pdf_and_docx_generation(self):
        pdf_bytes = build_consolidated_in65_pdf_bytes(self.sample_items)
        self.assertTrue(pdf_bytes.startswith(b"%PDF"))
        docx_bytes = build_consolidated_in65_docx_bytes(self.sample_items)
        self.assertTrue(docx_bytes.startswith(b"PK"))

    def test_consolidated_excel_has_all_companies_sheet_and_expurgos(self):
        xlsx_bytes = build_consolidated_in65_excel_bytes(self.sample_items)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        self.assertIn("Resumo Geral do Lote", wb.sheetnames)
        self.assertIn("Mapa Geral - Todas Cotações", wb.sheetnames)
        self.assertIn("Expurgos e Inconsistências", wb.sheetnames)
        self.assertIn("Resumo por Categoria", wb.sheetnames)

        ws_all = wb["Mapa Geral - Todas Cotações"]
        # 21 items * (2 valid + 1 purged) = 63 quote rows + 3 header rows = 66 rows
        self.assertEqual(ws_all.max_row, 66)
        # Verify individual item sheets also exist since 21 <= 30
        item_tabs = [s for s in wb.sheetnames if s.startswith("Item_")]
        self.assertEqual(len(item_tabs), 21)

    def test_agu_nota_tecnica_docx_and_pdf_follow_official_model(self):
        nt_docx = build_agu_nota_tecnica_docx_bytes(self.sample_items, om_name="CAPITANIA FLUVIAL")
        doc = docx.Document(io.BytesIO(nt_docx))
        full_text = "\n".join(p.text for p in doc.paragraphs)
        for section_hdr in [
            "NOTA TÉCNICA nº",
            "I - OBJETO DA CONTRATAÇÃO",
            "II - FONTES CONSULTADAS",
            "III - SÉRIE DE PREÇOS COLETADOS",
            "IV - METODOLOGIA PARA OBTENÇÃO DO PREÇO ESTIMADO",
            "V - MEMÓRIA DE CÁLCULO E CONCLUSÃO",
            "VI - IDENTIFICAÇÃO DOS AGENTES RESPONSÁVEIS PELA PESQUISA DE PREÇOS",
            "reais",
        ]:
            self.assertIn(section_hdr, full_text)
        # 4 tables: Consulta Direta, Categorias, Expurgos (§ 4.2), Memória (§ 5.1) + Assinaturas
        self.assertGreaterEqual(len(doc.tables), 5)

        nt_pdf = build_agu_nota_tecnica_pdf_bytes(self.sample_items, om_name="CAPITANIA FLUVIAL")
        self.assertTrue(nt_pdf.startswith(b"%PDF"))
        self.assertGreater(len(nt_pdf), 2000)

    def test_fallback_12m_to_24m_automatic_controller(self):
        ctrl = PriceResearchController()
        mock_quote_24m = RawQuoteInput(
            id_compra_item="7823400500001202500001",
            purchase_number_year="00001/2025",
            buyer_uasg="782340",
            buyer_org_name="CAPITANIA TESTE",
            buyer_uf="RJ",
            supplier_cnpj="12.345.678/0001-95",
            supplier_name="FORNECEDOR 24M LTDA",
            supply_unit="UN",
            quantity_homologated=10,
            unit_price=120.0,
            purchase_date="2025-01-15",
            pncp_or_compras_uri="https://pncp.gov.br/app/editais/00394502000144/2025/000001",
            codigo_classe="6515",
            nome_classe="EQUIPAMENTOS MEDICOS",
            is_extended_24m=True,
        )

        def fake_fetch(catalog_code, item_type="MATERIAL", months_back=12, uf=None, uasg=None):
            if months_back <= 12:
                return "", []
            return "ITEM RARO DE BAIXA LIQUIDEZ", [mock_quote_24m]

        ctrl.api_client.fetch_material_or_service_prices = MagicMock(side_effect=fake_fetch)
        ctrl.api_client.fetch_contract_item_prices = MagicMock(return_value=("", []))
        ctrl.api_client.fetch_active_arps = MagicMock(return_value=[])

        res = ctrl.execute_in65_research(catalog_code=999111, target_supply_unit="UN", months_back=12, allow_24m_fallback=True)
        self.assertEqual(res["status"], "SUCCESS")
        rep = res["report"]
        self.assertTrue(rep["is_extended_24m"])
        self.assertEqual(rep["search_window_months"], 24)
        self.assertEqual(rep["nome_classe"], "EQUIPAMENTOS MEDICOS")
        self.assertIn("13 a 24 MESES", rep["fallback_24m_note"])

    def test_batch_zip_separates_12m_and_24m_folders_and_groups_by_category(self):
        zip_bytes = build_batch_zip_bytes(self.sample_items, extract_gov_docs=False, mode="auto")
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            names = zf.namelist()
            self.assertIn("00_Mapa_Comparativo_Geral_Completo.xlsx", names)
            self.assertIn("01_Nota_Tecnica_AGU_IN65_Pesquisa_Precos.docx", names)
            self.assertIn("01_Nota_Tecnica_AGU_IN65_Pesquisa_Precos.pdf", names)
            self.assertIn("02_RaioX_Certidoes_Empresas_Consolidado.pdf", names)
            self.assertIn("02_RaioX_Certidoes_Empresas_Consolidado.docx", names)
            self.assertIn("03_Justificativa_Consolidada_Art6_Par5.docx", names)
            self.assertIn("03_Justificativa_Consolidada_Art6_Par5.pdf", names)
            self.assertIn("04_Comprovantes_PNCP_Contratacoes_QRCode.pdf", names)
            self.assertIn("00_MANUAL_DIDATICO_DO_PROCESSO_LEIA_ME.pdf", names)
            self.assertIn("00_LEIA_ME_MANUAL_DIDATICO.txt", names)

            self.assertTrue(zf.read("04_Comprovantes_PNCP_Contratacoes_QRCode.pdf").startswith(b"%PDF"))
            self.assertTrue(zf.read("00_MANUAL_DIDATICO_DO_PROCESSO_LEIA_ME.pdf").startswith(b"%PDF"))
            self.assertTrue(zf.read("02_RaioX_Certidoes_Empresas_Consolidado.docx").startswith(b"PK"))
            self.assertIn("PASSO 1", zf.read("00_LEIA_ME_MANUAL_DIDATICO.txt").decode("utf-8"))

            # Separate folders for <=12M vs 13-24M, including Relatorios_CATMAT_CATSER for every item!
            folder_12m = [n for n in names if n.startswith("01_Pesquisas_Ate_12_Meses_IN65/")]
            folder_24m = [n for n in names if n.startswith("02_Pesquisas_13_a_24_Meses_Excepcional/")]
            self.assertTrue(any("Relatorio_Detalhado_Completo_12M.pdf" in n for n in folder_12m))
            self.assertTrue(any("/Por_Categoria/" in n for n in folder_12m))
            catmat_12m = [n for n in folder_12m if "/Relatorios_CATMAT_CATSER/" in n and n.endswith(".pdf")]
            self.assertEqual(len(catmat_12m), 16)
            self.assertTrue(any("Relatorio_Detalhado_Completo_13_a_24M.pdf" in n for n in folder_24m))
            self.assertTrue(any("/Por_Categoria/" in n for n in folder_24m))
            catmat_24m = [n for n in folder_24m if "/Relatorios_CATMAT_CATSER/" in n and n.endswith(".pdf")]
            self.assertEqual(len(catmat_24m), 5)

    def test_pncp_direct_urls_never_use_generic_search_query(self):
        from core.gov_api_client import GovProcurementAPIClient
        u_pncp = GovProcurementAPIClient.build_audit_uri(
            "75121206000702026", pncp_control_number="00394502000144-1-010123/2026"
        )
        self.assertEqual(u_pncp, "https://pncp.gov.br/app/editais/00394502000144/2026/10123")
        u_fallback = GovProcurementAPIClient.build_audit_uri("75121206000702026")
        self.assertEqual(
            u_fallback,
            "https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras/acompanhamento-compra?compra=75121206000702026",
        )
        self.assertNotIn("?q=", u_fallback)

    def test_150_items_large_volume_zip_fast_and_complete(self):
        cats = [
            "MATERIAL DE EXPEDIENTE",
            "EQUIPAMENTOS DE INFORMATICA",
            "MATERIAL HOSPITALAR",
            "FERRAMENTAS MANUAIS",
            "MATERIAL ELETRICO",
            "LIMPEZA E HIGIENE",
        ]
        items_150 = [
            _make_sample_item(i, is_24m=(i > 130), cat_name=cats[i % len(cats)])
            for i in range(1, 151)
        ]
        t0 = time.perf_counter()
        zip_bytes = build_batch_zip_bytes(items_150, mode="auto")
        elapsed = time.perf_counter() - t0
        self.assertLess(elapsed, 60.0, f"ZIP de 150 itens demorou {elapsed:.2f}s (limite 60s)")
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            names = zf.namelist()
            self.assertIn("00_Mapa_Comparativo_Geral_Completo.xlsx", names)
            self.assertIn("01_Nota_Tecnica_AGU_IN65_Pesquisa_Precos.docx", names)
            self.assertIn("02_RaioX_Certidoes_Empresas_Consolidado.docx", names)
            self.assertIn("00_MANUAL_DIDATICO_DO_PROCESSO_LEIA_ME.pdf", names)
            all_catmat_pdfs = [n for n in names if "/Relatorios_CATMAT_CATSER/" in n and n.endswith(".pdf")]
            self.assertEqual(len(all_catmat_pdfs), 150)

    def test_async_job_manager_streams_zip_with_progress(self):
        jm = BatchJobManager(max_workers=1)
        items_subset = self.sample_items[:5]

        def _worker(out_path, cb):
            return build_batch_zip_to_file(items_subset, out_path, mode="auto", progress_callback=cb)

        job = jm.submit_file_job(
            kind="batch_zip",
            filename="teste.zip",
            content_type="application/zip",
            total_items=len(items_subset),
            worker_fn=_worker,
        )
        jid = job["job_id"]
        deadline = time.time() + 15.0
        st = None
        while time.time() < deadline:
            st = jm.get_status(jid)
            if st["status"] in ("done", "error"):
                break
            time.sleep(0.1)
        self.assertIsNotNone(st)
        self.assertEqual(st["status"], "done")
        self.assertEqual(st["percent"], 100)
        dl = jm.get_download_info(jid)
        self.assertIsNotNone(dl)
        self.assertGreater(dl["file_size"], 2000)
        try:
            os.remove(dl["file_path"])
        except OSError:
            pass


if __name__ == "__main__":
    unittest.main()
