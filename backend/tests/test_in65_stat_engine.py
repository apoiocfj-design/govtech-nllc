# -*- coding: utf-8 -*-
"""
SUÍTE DE TESTES UNITÁRIOS PYTHON (IN SEGES/ME Nº 65/2021)
Validação do Motor Estatístico backend/core/stat_engine.py
"""

import os
import sys
import unittest

# Adiciona backend/core ao path
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
BACKEND_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
CORE_DIR = os.path.join(BACKEND_DIR, "core")
if CORE_DIR not in sys.path:
    sys.path.insert(0, CORE_DIR)

from stat_engine import IN65StatisticalEngine, RawQuoteInput


class TestIN65StatisticalEngine(unittest.TestCase):
    def setUp(self):
        self.engine = IN65StatisticalEngine(
            cv_threshold_percent=25.0,
            std_dev_multiplier=1.0,
            infeasible_median_ratio=0.25,
            excessive_median_ratio=4.0
        )

    def test_caso1_amostra_homogenea_adota_media(self):
        """Amostra com CV <= 25% deve adotar obrigatoriamente a MÉDIA ARITMÉTICA."""
        quotes = [
            RawQuoteInput(id_compra_item="1", purchase_number_year="1/2026", buyer_uasg="123456", buyer_org_name="Org 1", buyer_uf="RJ", supplier_cnpj="111", supplier_name="Fornecedor 1", supply_unit="UN", quantity_homologated=10, unit_price=10.0, purchase_date="2026-01-01", pncp_or_compras_uri="https://pncp.gov.br"),
            RawQuoteInput(id_compra_item="2", purchase_number_year="2/2026", buyer_uasg="123456", buyer_org_name="Org 2", buyer_uf="SP", supplier_cnpj="222", supplier_name="Fornecedor 2", supply_unit="UN", quantity_homologated=10, unit_price=10.5, purchase_date="2026-01-02", pncp_or_compras_uri="https://pncp.gov.br"),
            RawQuoteInput(id_compra_item="3", purchase_number_year="3/2026", buyer_uasg="123456", buyer_org_name="Org 3", buyer_uf="MG", supplier_cnpj="333", supplier_name="Fornecedor 3", supply_unit="UN", quantity_homologated=10, unit_price=9.8, purchase_date="2026-01-03", pncp_or_compras_uri="https://pncp.gov.br"),
            RawQuoteInput(id_compra_item="4", purchase_number_year="4/2026", buyer_uasg="123456", buyer_org_name="Org 4", buyer_uf="PR", supplier_cnpj="444", supplier_name="Fornecedor 4", supply_unit="UN", quantity_homologated=10, unit_price=10.2, purchase_date="2026-01-04", pncp_or_compras_uri="https://pncp.gov.br"),
        ]

        report = self.engine.process_quotes(
            catmat_code=12345,
            item_description="CANETA ESFEROGRAFICA",
            target_unit="UN",
            quotes=quotes
        )

        self.assertTrue(report.is_homogeneous)
        self.assertLessEqual(report.sanitized_metrics.cv_percent, 25.0)
        self.assertEqual(report.adopted_parameter, "MEDIA")
        self.assertEqual(report.final_estimated_unit_price, 10.13)
        self.assertTrue(report.has_minimum_three_quotes)
        self.assertEqual(len(report.valid_quotes), 4)

    def test_caso2_amostra_heterogenea_adota_mediana(self):
        """Amostra com CV > 25% (n=3) deve adotar obrigatoriamente a MEDIANA."""
        quotes = [
            RawQuoteInput(id_compra_item="1", purchase_number_year="1/2026", buyer_uasg="123456", buyer_org_name="Org 1", buyer_uf="RJ", supplier_cnpj="111", supplier_name="Fornecedor 1", supply_unit="UN", quantity_homologated=10, unit_price=10.0, purchase_date="2026-01-01", pncp_or_compras_uri="https://pncp.gov.br"),
            RawQuoteInput(id_compra_item="2", purchase_number_year="2/2026", buyer_uasg="123456", buyer_org_name="Org 2", buyer_uf="SP", supplier_cnpj="222", supplier_name="Fornecedor 2", supply_unit="UN", quantity_homologated=10, unit_price=15.0, purchase_date="2026-01-02", pncp_or_compras_uri="https://pncp.gov.br"),
            RawQuoteInput(id_compra_item="3", purchase_number_year="3/2026", buyer_uasg="123456", buyer_org_name="Org 3", buyer_uf="MG", supplier_cnpj="333", supplier_name="Fornecedor 3", supply_unit="UN", quantity_homologated=10, unit_price=20.0, purchase_date="2026-01-03", pncp_or_compras_uri="https://pncp.gov.br"),
        ]

        report = self.engine.process_quotes(
            catmat_code=12345,
            item_description="ITEM COM DISPERSAO RESIDUAL",
            target_unit="UN",
            quotes=quotes
        )

        self.assertFalse(report.is_homogeneous)
        self.assertGreater(report.sanitized_metrics.cv_percent, 25.0)
        self.assertEqual(report.adopted_parameter, "MEDIANA")
        self.assertEqual(report.final_estimated_unit_price, 15.0)
        self.assertTrue(report.has_minimum_three_quotes)

    def test_caso3_expurgo_outliers_inexequivel_e_sobrepreco(self):
        """Valores discrepantes devem ser expurgados com purge_status correspondente."""
        quotes = [
            RawQuoteInput(id_compra_item="1", purchase_number_year="1/2026", buyer_uasg="123456", buyer_org_name="Org 1", buyer_uf="RJ", supplier_cnpj="111", supplier_name="Fornecedor 1", supply_unit="UN", quantity_homologated=10, unit_price=10.0, purchase_date="2026-01-01", pncp_or_compras_uri="https://pncp.gov.br"),
            RawQuoteInput(id_compra_item="2", purchase_number_year="2/2026", buyer_uasg="123456", buyer_org_name="Org 2", buyer_uf="SP", supplier_cnpj="222", supplier_name="Fornecedor 2", supply_unit="UN", quantity_homologated=10, unit_price=10.5, purchase_date="2026-01-02", pncp_or_compras_uri="https://pncp.gov.br"),
            RawQuoteInput(id_compra_item="3", purchase_number_year="3/2026", buyer_uasg="123456", buyer_org_name="Org 3", buyer_uf="MG", supplier_cnpj="333", supplier_name="Fornecedor 3", supply_unit="UN", quantity_homologated=10, unit_price=9.8, purchase_date="2026-01-03", pncp_or_compras_uri="https://pncp.gov.br"),
            RawQuoteInput(id_compra_item="4", purchase_number_year="4/2026", buyer_uasg="123456", buyer_org_name="Org 4", buyer_uf="RS", supplier_cnpj="444", supplier_name="Fornecedor 4", supply_unit="UN", quantity_homologated=10, unit_price=100.0, purchase_date="2026-01-04", pncp_or_compras_uri="https://pncp.gov.br"),  # Outlier > 4x
            RawQuoteInput(id_compra_item="5", purchase_number_year="5/2026", buyer_uasg="123456", buyer_org_name="Org 5", buyer_uf="BA", supplier_cnpj="555", supplier_name="Fornecedor 5", supply_unit="UN", quantity_homologated=10, unit_price=0.50, purchase_date="2026-01-05", pncp_or_compras_uri="https://pncp.gov.br"),  # Outlier < 0.25x
        ]

        report = self.engine.process_quotes(
            catmat_code=12345,
            item_description="TESTE OUTLIERS",
            target_unit="UN",
            quotes=quotes
        )

        purged_ids = [q.id_compra_item for q in report.purged_quotes]
        self.assertIn("4", purged_ids)
        self.assertIn("5", purged_ids)

        q4 = next(q for q in report.purged_quotes if q.id_compra_item == "4")
        self.assertTrue(q4.is_outlier)
        self.assertEqual(q4.purge_status, "EXPURGADA_ACIMA_LIMITE")

        q5 = next(q for q in report.purged_quotes if q.id_compra_item == "5")
        self.assertTrue(q5.is_outlier)
        self.assertEqual(q5.purge_status, "EXPURGADA_INEXEQUIVEL")

        self.assertEqual(len(report.valid_quotes), 3)

    def test_caso4_alerta_cesta_restrita_abaixo_tres_cotacoes(self):
        """Menos de 3 cotações válidas deve acionar has_minimum_three_quotes=False."""
        quotes = [
            RawQuoteInput(id_compra_item="1", purchase_number_year="1/2026", buyer_uasg="123456", buyer_org_name="Org 1", buyer_uf="RJ", supplier_cnpj="111", supplier_name="Fornecedor 1", supply_unit="UN", quantity_homologated=10, unit_price=10.0, purchase_date="2026-01-01", pncp_or_compras_uri="https://pncp.gov.br"),
            RawQuoteInput(id_compra_item="2", purchase_number_year="2/2026", buyer_uasg="123456", buyer_org_name="Org 2", buyer_uf="SP", supplier_cnpj="222", supplier_name="Fornecedor 2", supply_unit="UN", quantity_homologated=10, unit_price=10.5, purchase_date="2026-01-02", pncp_or_compras_uri="https://pncp.gov.br"),
        ]

        report = self.engine.process_quotes(
            catmat_code=12345,
            item_description="CESTA RESTRITA",
            target_unit="UN",
            quotes=quotes
        )

        self.assertFalse(report.has_minimum_three_quotes)
        self.assertIn("ATENÇÃO AUDITORIA", report.legal_methodology_justification)


if __name__ == "__main__":
    unittest.main()
