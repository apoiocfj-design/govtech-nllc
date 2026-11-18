# -*- coding: utf-8 -*-
"""
Testes Unitários da Camada de Consolidação em Lote, Pacote ZIP e Jobs Assíncronos (IN 65/2021)
"""

import io
import os
import sys
import time
import unittest
import zipfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "core")))

from core.report_generators import (
    build_consolidated_in65_pdf_bytes,
    build_consolidated_in65_docx_bytes,
    build_consolidated_in65_excel_bytes,
    build_consolidated_suppliers_dossier_pdf_bytes,
    build_batch_zip_to_file,
    build_batch_zip_bytes,
)
from core.job_manager import BatchJobManager


class TestBatchReports(unittest.TestCase):

    def setUp(self):
        self.sample_items = []
        for i in range(1, 22):
            self.sample_items.append({
                "item_index": i,
                "code": 600000 + i,
                "type": "MATERIAL",
                "unit": "UN",
                "qty": 5,
                "report": {
                    "catmat_catser_code": 600000 + i,
                    "item_description": f"MATERIAL DE CONSUMO ITEM TESTE {i}",
                    "target_supply_unit": "UN",
                    "final_estimated_unit_price": 50.0 + i,
                    "final_metric_used": "MEDIANA",
                    "adopted_parameter": "MEDIANA",
                    "is_homogeneous": True,
                    "valid_quotes": [
                        {"unit_price": 50.0 + i, "supplier_cnpj": "12345678000195", "supplier_name": "FORNECEDOR ALFA"}
                    ],
                    "sanitized_metrics": {"mean": 50.0 + i, "median": 50.0 + i, "cv_percent": 0.0, "std_deviation": 0.0},
                },
            })

    def test_consolidated_pdf_generation(self):
        pdf_bytes = build_consolidated_in65_pdf_bytes(self.sample_items)
        self.assertIsInstance(pdf_bytes, bytes)
        self.assertGreater(len(pdf_bytes), 1000)
        self.assertTrue(pdf_bytes.startswith(b"%PDF"))

    def test_consolidated_docx_generation(self):
        docx_bytes = build_consolidated_in65_docx_bytes(self.sample_items)
        self.assertIsInstance(docx_bytes, bytes)
        self.assertGreater(len(docx_bytes), 1000)
        self.assertTrue(docx_bytes.startswith(b"PK"))

    def test_consolidated_excel_generation(self):
        xlsx_bytes = build_consolidated_in65_excel_bytes(self.sample_items)
        self.assertIsInstance(xlsx_bytes, bytes)
        self.assertGreater(len(xlsx_bytes), 1000)
        self.assertTrue(xlsx_bytes.startswith(b"PK"))

    def test_consolidated_suppliers_dossier_pdf(self):
        pdf_bytes = build_consolidated_suppliers_dossier_pdf_bytes(self.sample_items)
        self.assertTrue(pdf_bytes.startswith(b"%PDF"))
        self.assertGreater(len(pdf_bytes), 1000)

    def test_batch_zip_large_batch_contains_all_master_and_individual_pdfs(self):
        # 21 items (> 15 items): deve conter todos os 4 arquivos mestres + README + 21 PDFs individuais!
        zip_bytes = build_batch_zip_bytes(self.sample_items, extract_gov_docs=False, mode="auto")
        self.assertIsInstance(zip_bytes, bytes)
        self.assertGreater(len(zip_bytes), 5000)

        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            names = zf.namelist()
            self.assertIn("00_Mapa_Comparativo_Consolidado_Lote.xlsx", names)
            self.assertIn("01_Caderno_Consolidado_Precos_Lote.pdf", names)
            self.assertIn("02_Caderno_Consolidado_Precos_Lote.docx", names)
            self.assertIn("03_Dossie_Consolidado_Certidoes_Fornecedores.pdf", names)
            self.assertIn("00_LEIA_ME_ESTRUTURA_DO_LOTE.txt", names)
            indiv_pdfs = [n for n in names if n.startswith("04_Relatorios_Individuais_PDF/") and n.endswith(".pdf")]
            self.assertEqual(len(indiv_pdfs), 21)

    def test_async_job_manager_streams_zip_with_progress(self):
        jm = BatchJobManager(max_workers=1)
        items_subset = self.sample_items[:5]

        def _worker(out_path, cb):
            return build_batch_zip_to_file(items_subset, out_path, mode="consolidated", progress_callback=cb)

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
