# -*- coding: utf-8 -*-
"""
Testes Unitários e de Integração — Motor Master de Auditoria Documental (NLLC / AGU)
Valida:
1. Análise sem mutação automática (permissão obrigatória do usuário).
2. Cirurgia XML in-place em .DOCX original preservando 100% cabeçalho, rodapé, fontes e tabelas.
3. Marcação em AMARELO para alterado/inserido e VERMELHO TACHADO para retirado (parágrafos e tabelas).
4. Conversão de alta fidelidade PDF -> DOCX respeitando alterações permitidas/recusadas.
"""

import io
import os
import sys
import unittest

import docx
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt, RGBColor

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "core")))

from core.audited_document_engine import (
    DocumentLayoutParser,
    DocumentLegalAuditor,
    build_audited_document_docx_bytes,
    build_audited_document_pdf_bytes,
    document_session_store,
)


class TestAuditedDocumentEngine(unittest.TestCase):

    def _create_sample_docx(self) -> bytes:
        d = docx.Document()
        d.sections[0].header.paragraphs[0].text = "CABECALHO ORIGINAL MARINHA DO BRASIL"
        d.sections[0].footer.paragraphs[0].text = "RODAPE ORIGINAL DO DOCUMENTO"

        p0 = d.add_paragraph()
        r0 = p0.add_run("1. OBJETO E FUNDAMENTAÇÃO")
        r0.font.name = "Georgia"
        r0.font.size = Pt(12)
        r0.font.bold = True

        p1 = d.add_paragraph()
        r1 = p1.add_run("1.1. O presente processo segue a Lei nº 8.666/1993 e admite empresa EIRELI.")
        r1.font.name = "Georgia"
        r1.font.size = Pt(11)

        tbl = d.add_table(rows=2, cols=2)
        tbl.cell(0, 0).text = "NORMA"
        tbl.cell(0, 1).text = "FUNDAMENTAÇÃO"
        tbl.cell(1, 0).text = "Dispensa"
        tbl.cell(1, 1).text = "Enquadramento no art. 24, inciso II da Lei 8.666."

        buf = io.BytesIO()
        d.save(buf)
        return buf.getvalue()

    def test_permission_gate_blocks_unapproved_mutations(self):
        orig_bytes = self._create_sample_docx()
        ds = DocumentLayoutParser.parse_docx(orig_bytes, "minuta_teste.docx")
        audited = DocumentLegalAuditor.audit_and_diff(ds, require_permission=True)

        findings = audited["findings"]
        self.assertGreaterEqual(len(findings), 3)
        # Nenhum finding deve vir aprovado automaticamente
        self.assertTrue(all(not f["approved"] for f in findings))

        # Se o usuário não aprovar NENHUMA alteração, o DOCX gerado deve permanecer 100% original
        resolved, actions = DocumentLegalAuditor.apply_approved_decisions(audited, [])
        self.assertEqual(len(actions), 0)
        out_bytes = build_audited_document_docx_bytes(resolved, original_bytes=orig_bytes, applied_actions=actions)
        d_out = docx.Document(io.BytesIO(out_bytes))
        self.assertIn("8.666/1993", d_out.paragraphs[1].text)
        yellow_runs = [r for r in d_out.paragraphs[1].runs if r.font.highlight_color == WD_COLOR_INDEX.YELLOW]
        self.assertEqual(len(yellow_runs), 0)

    def test_surgical_docx_preserves_layout_and_applies_yellow_red(self):
        orig_bytes = self._create_sample_docx()
        ds = DocumentLayoutParser.parse_docx(orig_bytes, "minuta_teste.docx")
        audited = DocumentLegalAuditor.audit_and_diff(ds, require_permission=True)

        # Aprova alterações em parágrafo e tabela; recusa omissões
        decisions = [
            {"id": f["id"], "approved": f["target_type"] != "append_clause"}
            for f in audited["findings"]
        ]
        resolved, actions = DocumentLegalAuditor.apply_approved_decisions(audited, decisions)
        self.assertEqual(len(actions), 2)

        out_bytes = build_audited_document_docx_bytes(resolved, original_bytes=orig_bytes, applied_actions=actions)
        d_out = docx.Document(io.BytesIO(out_bytes))

        # 1. Cabeçalho e Rodapé originais intactos
        self.assertEqual(d_out.sections[0].header.paragraphs[0].text, "CABECALHO ORIGINAL MARINHA DO BRASIL")
        self.assertEqual(d_out.sections[0].footer.paragraphs[0].text, "RODAPE ORIGINAL DO DOCUMENTO")

        # 2. Fonte Georgia preservada + destaque amarelo + vermelho tachado no parágrafo 1.1
        p_target = d_out.paragraphs[1]
        fonts = {r.font.name for r in p_target.runs if r.text.strip()}
        self.assertIn("Georgia", fonts)
        yellow = [r.text for r in p_target.runs if r.font.highlight_color == WD_COLOR_INDEX.YELLOW]
        red = [r.text for r in p_target.runs if r.font.strike and r.font.color and r.font.color.rgb == RGBColor(220, 38, 38)]
        self.assertTrue(any("14.133/2021" in t for t in yellow))
        self.assertTrue(any("SLU" in t for t in yellow))
        self.assertTrue(any("8.666/1993" in t for t in red))
        self.assertTrue(any("EIRELI" in t for t in red))

        # 3. Tabela original preservada com amarelo/vermelho dentro da célula (1, 1)
        cell_p = d_out.tables[0].cell(1, 1).paragraphs[0]
        tbl_yellow = [r.text for r in cell_p.runs if r.font.highlight_color == WD_COLOR_INDEX.YELLOW]
        tbl_red = [r.text for r in cell_p.runs if r.font.strike and r.font.color and r.font.color.rgb == RGBColor(220, 38, 38)]
        self.assertTrue(any("75," in t or "14.133/2021" in t for t in tbl_yellow))
        self.assertTrue(any("24," in t or "8.666" in t for t in tbl_red))

    def test_session_store_and_pdf_export(self):
        orig_bytes = self._create_sample_docx()
        ds = DocumentLayoutParser.parse_docx(orig_bytes, "minuta_sessao.docx")
        audited = DocumentLegalAuditor.audit_and_diff(ds, require_permission=True)
        sid = document_session_store.save(orig_bytes, "minuta_sessao.docx", "docx", audited)
        loaded = document_session_store.get(sid)
        self.assertIsNotNone(loaded)
        self.assertEqual(loaded["original_bytes"], orig_bytes)

        decisions = [{"id": f["id"], "approved": True} for f in audited["findings"]]
        resolved, _ = DocumentLegalAuditor.apply_approved_decisions(loaded["doc_structure"], decisions)
        pdf_bytes = build_audited_document_pdf_bytes(resolved)
        self.assertTrue(pdf_bytes.startswith(b"%PDF"))
        self.assertGreater(len(pdf_bytes), 2000)


if __name__ == "__main__":
    unittest.main()
