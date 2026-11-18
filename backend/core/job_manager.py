# -*- coding: utf-8 -*-
"""
GERENCIADOR DE FILA ASSÍNCRONA DE JOBS E STREAMING EM DISCO (ENTERPRISE)
========================================================================
Elimina completamente erros HTTP 504 (Gateway Timeout) e OOM (Error 137)
em lotes grandes (50, 150, 300+ itens) no Render / Cloud Run:
1. POST /api/v1/jobs/batch-zip -> retorna 202 Accepted em < 50ms com job_id.
2. Worker dedicado executa em background escrevendo o ZIP direto em disco.
3. GET /api/v1/jobs/<job_id>/status -> progresso real (percentual, item atual, etapa).
4. GET /api/v1/jobs/<job_id>/download -> streaming em chunks de 64 KB direto do disco.
"""

from __future__ import annotations
import os
import tempfile
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Callable, Dict, Optional


class BatchJobManager:
    """Gerenciador thread-safe de tarefas pesadas em background com limpeza automática (TTL)."""

    def __init__(self, max_workers: int = 2, ttl_seconds: int = 1500, max_jobs: int = 20):
        self._lock = threading.Lock()
        self._jobs: Dict[str, Dict[str, Any]] = {}
        self._executor = ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="govtech-job")
        self._ttl_seconds = ttl_seconds
        self._max_jobs = max_jobs

    def _cleanup_expired_locked(self) -> None:
        now = time.time()
        to_delete = []
        for jid, info in self._jobs.items():
            age = now - info.get("updated_at", now)
            if info.get("status") in ("done", "error") and age > self._ttl_seconds:
                to_delete.append(jid)

        # Se ainda exceder max_jobs, remove os concluídos mais antigos
        if len(self._jobs) - len(to_delete) > self._max_jobs:
            finished = [
                (info.get("updated_at", 0), jid)
                for jid, info in self._jobs.items()
                if jid not in to_delete and info.get("status") in ("done", "error")
            ]
            finished.sort()
            excess = (len(self._jobs) - len(to_delete)) - self._max_jobs
            for _, jid in finished[:max(0, excess)]:
                to_delete.append(jid)

        for jid in to_delete:
            info = self._jobs.pop(jid, None)
            if info and info.get("file_path"):
                try:
                    if os.path.exists(info["file_path"]):
                        os.remove(info["file_path"])
                except OSError:
                    pass

    def submit_file_job(
        self,
        kind: str,
        filename: str,
        content_type: str,
        total_items: int,
        worker_fn: Callable[[str, Callable[[int, int, str], None]], str],
    ) -> Dict[str, Any]:
        job_id = f"job_{uuid.uuid4().hex[:16]}"
        fd, tmp_path = tempfile.mkstemp(prefix=f"govtech_{kind}_", suffix=".tmp")
        os.close(fd)

        now = time.time()
        record: Dict[str, Any] = {
            "job_id": job_id,
            "kind": kind,
            "status": "queued",
            "percent": 0,
            "done_steps": 0,
            "total_steps": max(1, total_items),
            "total_items": total_items,
            "stage": "Na fila de processamento assíncrono...",
            "created_at": now,
            "updated_at": now,
            "finished_at": None,
            "file_path": tmp_path,
            "filename": filename,
            "content_type": content_type,
            "file_size": 0,
            "error": None,
        }

        with self._lock:
            self._cleanup_expired_locked()
            self._jobs[job_id] = record

        def _progress_cb(done: int, total: int, stage: str) -> None:
            pct = int(round((done / max(1, total)) * 100)) if total > 0 else 0
            pct = max(0, min(99, pct))
            with self._lock:
                j = self._jobs.get(job_id)
                if j and j["status"] in ("queued", "running"):
                    j["status"] = "running"
                    j["done_steps"] = done
                    j["total_steps"] = max(1, total)
                    j["percent"] = pct
                    j["stage"] = stage
                    j["updated_at"] = time.time()

        def _run() -> None:
            with self._lock:
                j = self._jobs.get(job_id)
                if j:
                    j["status"] = "running"
                    j["stage"] = "Iniciando motor de consolidação em alta performance..."
                    j["updated_at"] = time.time()
            try:
                out_path = worker_fn(tmp_path, _progress_cb)
                fsize = os.path.getsize(out_path) if os.path.exists(out_path) else 0
                with self._lock:
                    j = self._jobs.get(job_id)
                    if j:
                        j["status"] = "done"
                        j["percent"] = 100
                        j["done_steps"] = j["total_steps"]
                        j["stage"] = "Arquivo pronto para download!"
                        j["file_path"] = out_path
                        j["file_size"] = fsize
                        j["finished_at"] = time.time()
                        j["updated_at"] = time.time()
            except Exception as exc:
                try:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)
                except OSError:
                    pass
                with self._lock:
                    j = self._jobs.get(job_id)
                    if j:
                        j["status"] = "error"
                        j["error"] = str(exc)
                        j["stage"] = f"Erro no processamento: {exc}"
                        j["finished_at"] = time.time()
                        j["updated_at"] = time.time()

        self._executor.submit(_run)
        return self.get_status(job_id) or {}

    def get_status(self, job_id: str) -> Optional[Dict[str, Any]]:
        with self._lock:
            j = self._jobs.get(job_id)
            if not j:
                return None
            elapsed = round((j["finished_at"] or time.time()) - j["created_at"], 2)
            return {
                "job_id": j["job_id"],
                "kind": j["kind"],
                "status": j["status"],
                "percent": j["percent"],
                "done_steps": j["done_steps"],
                "total_steps": j["total_steps"],
                "total_items": j["total_items"],
                "stage": j["stage"],
                "filename": j["filename"],
                "file_size": j["file_size"],
                "elapsed_sec": elapsed,
                "error": j["error"],
            }

    def get_download_info(self, job_id: str) -> Optional[Dict[str, Any]]:
        with self._lock:
            j = self._jobs.get(job_id)
            if not j or j["status"] != "done":
                return None
            if not j.get("file_path") or not os.path.exists(j["file_path"]):
                return None
            j["updated_at"] = time.time()
            return {
                "file_path": j["file_path"],
                "filename": j["filename"],
                "content_type": j["content_type"],
                "file_size": os.path.getsize(j["file_path"]),
            }


# Instância singleton global do gerenciador de jobs
batch_job_manager = BatchJobManager(max_workers=2, ttl_seconds=1500, max_jobs=20)
