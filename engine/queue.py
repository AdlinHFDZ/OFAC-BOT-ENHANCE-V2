# engine/queue.py
"""
Job Queue – runs multiple ExtractionJob instances sequentially in a background thread.
Communicates with the GUI via callbacks.
Now supports a per‑job output folder.
"""

import threading
from dataclasses import dataclass, field
from typing import List, Optional, Callable, Dict

from engine.job import ExtractionJob


@dataclass
class QueuedJob:
    """Holds the configuration for one scan job."""
    company_code: str
    passwords: List[str]
    email_received_date: str
    file_names: List[str]
    input_folder: str
    mapping_overrides: Optional[Dict] = None
    output_folder: Optional[str] = None      # <-- NEW: per‑job output folder
    status: str = 'waiting'
    error_message: str = ''


class JobQueue:
    def __init__(
        self,
        on_log: Optional[Callable[[str], None]] = None,
        on_progress: Optional[Callable[[int, int], None]] = None,
        on_job_start: Optional[Callable[['QueuedJob'], None]] = None,
        on_job_finish: Optional[Callable[['QueuedJob'], None]] = None,
        on_queue_empty: Optional[Callable[[], None]] = None,
    ):
        self._queue: List[QueuedJob] = []
        self._lock = threading.Lock()
        self._current_job: Optional[QueuedJob] = None
        self._stop_requested = False
        self._running = False
        self._thread: Optional[threading.Thread] = None

        # UI callbacks
        self.on_log = on_log
        self.on_progress = on_progress
        self.on_job_start = on_job_start
        self.on_job_finish = on_job_finish
        self.on_queue_empty = on_queue_empty

    def add_job(self, job: QueuedJob):
        with self._lock:
            self._queue.append(job)
        if self.on_log:
            self.on_log(f"Added to queue: {job.company_code} ({len(job.file_names)} files)")

    def remove_job(self, index: int):
        with self._lock:
            if 0 <= index < len(self._queue):
                removed = self._queue.pop(index)
                if self.on_log:
                    self.on_log(f"Removed from queue: {removed.company_code}")
                return
        raise IndexError("Invalid queue index")

    def clear(self):
        with self._lock:
            self._queue.clear()

    def stop(self):
        """Request stop of current job and clear remaining queue."""
        self._stop_requested = True
        with self._lock:
            self._queue.clear()

    def start(self):
        """Start processing the queue in a background thread if not already running."""
        if self._running:
            return
        self._stop_requested = False
        self._running = True
        self._thread = threading.Thread(target=self._process_queue, daemon=True)
        self._thread.start()

    def _process_queue(self):
        while not self._stop_requested:
            # Get next job
            with self._lock:
                if not self._queue:
                    break
                job = self._queue.pop(0)
                job.status = 'running'
                self._current_job = job

            if self.on_job_start:
                self.on_job_start(job)

            try:
                extraction_job = ExtractionJob(
                    input_folder=job.input_folder,
                    company_code=job.company_code,
                    passwords=job.passwords,
                    email_received_date=job.email_received_date,
                    file_names=job.file_names,
                    mapping_overrides=job.mapping_overrides,
                    output_root=job.output_folder,     # <-- pass user‑configured folder
                )

                success = extraction_job.run(
                    progress_callback=self.on_log,
                    stop_flag=lambda: self._stop_requested,
                    progress_update=self.on_progress,
                )

                if self._stop_requested:
                    job.status = 'cancelled'
                    if self.on_log:
                        self.on_log(f"Job cancelled: {job.company_code}")
                elif success:
                    job.status = 'completed'
                else:
                    job.status = 'error'
                    job.error_message = 'Extraction returned False'
            except Exception as e:
                job.status = 'error'
                job.error_message = str(e)
                if self.on_log:
                    self.on_log(f"Job failed: {job.company_code} – {e}")

            if self.on_job_finish:
                self.on_job_finish(job)

            self._current_job = None

        self._running = False
        if self.on_queue_empty:
            self.on_queue_empty()

    def get_queue_snapshot(self) -> List[QueuedJob]:
        with self._lock:
            return list(self._queue)