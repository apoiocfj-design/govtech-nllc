# -*- coding: utf-8 -*-
"""
Testes Unitários da Camada de Consolidação em Lote e Pacote ZIP (IN 65/2021)
Valida a geração de Cadernos Consolidados em PDF, Word (.DOCX) e pacote ZIP de lote amplo.
"""

import unittest
import zipfile
import io
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "core")))

from core.report_generators import (
    build_consolidated_in65_pdf_bytes,
    build_consolidated_in65_docx_bytes,
    build_consolidated_in65_excel_bytes,
    build_batch_zip_bytes,
)


class TestBatchReports(unittest.TestCase):

    def setUp(self):
        self.sample_items = []
        for i in range(1, 25):
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
                    "is_homogeneous": True,
                    "valid_quotes": [
                        {"price": 50.0 + i, "supplier_cnpj": "00000000000191", "supplier_name": "FORNECEDOR ALFA"}
                    ],
                    "sanitized_metrics": {"mean": 50.0 + i, "median": 50.0 + i, "cv_percent": 0.0},
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
        # PK header of zip/docx
        self.assertTrue(docx_bytes.startswith(b"PK"))

    def test_consolidated_excel_generation(self):
        xlsx_bytes = build_consolidated_in65_excel_bytes(self.sample_items)
        self.assertIsInstance(xlsx_bytes, bytes)
        self.assertGreater(len(xlsx_bytes), 1000)
        self.assertTrue(xlsx_bytes.startswith(b"PK"))

    def test_batch_zip_large_batch(self):
        # 24 items (> 15 items) should generate master files + top individual folders + README
        zip_bytes = build_batch_zip_bytes(self.sample_items, extract_gov_docs=False)
        self.assertIsInstance(zip_bytes, bytes)
        self.assertGreater(len(zip_bytes), 5000)

        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            names = zf.namelist()
            self.assertIn("00_Mapa_Comparativo_Consolidado_Lote.xlsx", names)
            self.assertIn("01_Caderno_Consolidado_Precos_Lote.pdf", names)
            self.assertIn("02_Caderno_Consolidado_Precos_Lote.docx", names)
            self.assertIn("00_LEIA_ME_LOTE_AMPLO.txt", names)


if __name__ == "__main__":
    unittest.main()
