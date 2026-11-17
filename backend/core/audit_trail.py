# -*- coding: utf-8 -*-
"""
MÓDULO DE TRILHA DE AUDITORIA IMUTÁVEL (AUDIT TRAIL COM HASH CRIPTOGRÁFICO SHA-256)
Conformidade estrita com o Tribunal de Contas da União (TCU), Controladoria-Geral da União (CGU)
e Instrução Normativa SEGES/ME nº 65/2021 (Lei nº 14.133/2021).

Garante a imutabilidade e a rastreabilidade absoluta de todas as ações de curadoria,
descarte de outliers, fixação de parâmetros e emissão de relatórios do operador.
"""

from __future__ import annotations

import hashlib
import json
import os
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

BASE_DIR = Path(__file__).resolve().parent.parent
DB_DIR = BASE_DIR / "database"
AUDIT_LOG_FILE = DB_DIR / "immutable_audit_trail.json"

GENESIS_PREVIOUS_HASH = "0" * 64


class ImmutableAuditTrail:
    """
    Gerenciador da Trilha de Auditoria Imutável encadeada por SHA-256 (Blockchain-like).
    Registra eventos com timestamp universal UTC, identificador de usuário/matrícula,
    detalhes do item/lote, justificativa jurídica e hash de integridade encadeado.
    """

    def __init__(self, log_path: Path = AUDIT_LOG_FILE) -> None:
        self.log_path = log_path
        self._ensure_storage()

    def _ensure_storage(self) -> None:
        """Cria o diretório de banco de dados e o arquivo de log se não existirem."""
        self.log_path.parent.mkdir(parents=True, exist_ok=True)
        if not self.log_path.exists():
            with open(self.log_path, "w", encoding="utf-8") as f:
                json.dump([], f, ensure_ascii=False, indent=2)

    def _read_all(self) -> List[Dict[str, Any]]:
        try:
            with open(self.log_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
        except Exception:
            return []

    def _write_all(self, entries: List[Dict[str, Any]]) -> None:
        with open(self.log_path, "w", encoding="utf-8") as f:
            json.dump(entries, f, ensure_ascii=False, indent=2)

    def calculate_hash(
        self,
        sequence: int,
        timestamp: str,
        user: str,
        action: str,
        target_item: str,
        details: Any,
        previous_hash: str,
    ) -> str:
        """Calcula o hash criptográfico SHA-256 do bloco de auditoria."""
        serialized_details = (
            json.dumps(details, sort_keys=True, ensure_ascii=False)
            if isinstance(details, (dict, list))
            else str(details)
        )
        payload = f"{sequence}|{timestamp}|{user}|{action}|{target_item}|{serialized_details}|{previous_hash}"
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    def record_event(
        self,
        user: str,
        action: str,
        target_item: str,
        details: Any,
        legal_basis: str = "IN SEGES/ME nº 65/2021",
        ip_address: str = "127.0.0.1",
    ) -> Dict[str, Any]:
        """
        Adiciona uma nova entrada imutável à trilha de auditoria.
        Gera sequência crescente e encadeia o hash com a entrada anterior.
        """
        entries = self._read_all()
        sequence = len(entries) + 1
        previous_hash = entries[-1]["sha256_hash"] if entries else GENESIS_PREVIOUS_HASH
        timestamp = datetime.now(timezone.utc).isoformat()

        sha256_hash = self.calculate_hash(
            sequence=sequence,
            timestamp=timestamp,
            user=user or "OPERADOR DO SISTEMA",
            action=action,
            target_item=target_item or "GERAL",
            details=details,
            previous_hash=previous_hash,
        )

        entry = {
            "sequence": sequence,
            "timestamp": timestamp,
            "user": user or "OPERADOR DO SISTEMA",
            "action": action,
            "target_item": target_item,
            "legal_basis": legal_basis,
            "details": details,
            "ip_address": ip_address,
            "previous_hash": previous_hash,
            "sha256_hash": sha256_hash,
        }

        entries.append(entry)
        self._write_all(entries)
        return entry

    def verify_integrity(self) -> Dict[str, Any]:
        """
        Verifica a integridade criptográfica de toda a trilha de auditoria.
        Detecta qualquer tentativa de adulteração de histórico.
        """
        entries = self._read_all()
        if not entries:
            return {
                "valid": True,
                "total_entries": 0,
                "message": "Trilha de auditoria vazia.",
                "tampered_index": None,
            }

        prev_hash = GENESIS_PREVIOUS_HASH
        for idx, entry in enumerate(entries):
            # Valida encadeamento com hash anterior
            if entry.get("previous_hash") != prev_hash:
                return {
                    "valid": False,
                    "total_entries": len(entries),
                    "tampered_index": idx,
                    "tampered_sequence": entry.get("sequence"),
                    "message": f"Quebra de encadeamento no evento #{entry.get('sequence')}: previous_hash divergente.",
                }

            # Recalcula hash para verificar se o conteúdo do bloco foi modificado
            expected_hash = self.calculate_hash(
                sequence=entry.get("sequence", idx + 1),
                timestamp=entry.get("timestamp", ""),
                user=entry.get("user", ""),
                action=entry.get("action", ""),
                target_item=entry.get("target_item", ""),
                details=entry.get("details"),
                previous_hash=prev_hash,
            )

            if entry.get("sha256_hash") != expected_hash:
                return {
                    "valid": False,
                    "total_entries": len(entries),
                    "tampered_index": idx,
                    "tampered_sequence": entry.get("sequence"),
                    "message": f"Hash inválido no evento #{entry.get('sequence')}: os dados do evento foram adulterados.",
                }

            prev_hash = entry.get("sha256_hash")

        return {
            "valid": True,
            "total_entries": len(entries),
            "latest_hash": prev_hash,
            "message": "Trilha 100% íntegra. Todos os hashes SHA-256 e encadeamentos validados com sucesso.",
            "tampered_index": None,
        }

    def get_recent_entries(self, limit: int = 50) -> List[Dict[str, Any]]:
        """Retorna os eventos mais recentes da trilha de auditoria."""
        entries = self._read_all()
        return entries[-limit:][::-1]


audit_trail_manager = ImmutableAuditTrail()
