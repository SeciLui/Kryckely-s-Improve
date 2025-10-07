#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Lesson Scribe – journal de leçons avec audio et transcription."""

from __future__ import annotations

import datetime
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid
import wave
from pathlib import Path
from queue import Empty, Queue
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

from dotenv import load_dotenv

try:  # Chargement facultatif pour l'enregistrement audio
    import sounddevice as _sd  # type: ignore
    import numpy as _np  # type: ignore
except Exception:  # pragma: no cover - dépendances optionnelles
    _sd = None  # type: ignore
    _np = None  # type: ignore


APP_NAME = "Lesson Scribe"
WORKSPACE_VERSION = 1
TRANSCRIPT_HEADER = "\n\n--- Transcription Vibe ---\n"
HELP_TEXT = (
    "Bienvenue dans Lesson Scribe !\n\n"
    "1. Ouvre ou crée un workspace.\n"
    "2. Ajoute une leçon : date, horaires, durée et journal libre.\n"
    "3. Optionnel : attache un audio ou enregistre directement.\n"
    "4. Lesson Scribe peut transcrire automatiquement l’audio (Vibe).\n"
    "5. Utilise le bouton ‘Copier le prompt d’analyse’ pour réviser."
)


# ---------------------------------------------------------------------------
# Utilitaires
# ---------------------------------------------------------------------------


def _load_env_from_config() -> None:
    env_path = os.environ.get("LESSON_SCRIBE_ENV_FILE")
    if env_path:
        load_dotenv(env_path, override=False)
    else:
        load_dotenv(override=False)


def _parse_default_time_range(raw_value: str | None) -> tuple[str | None, str | None]:
    if not raw_value:
        return None, None
    candidate = raw_value.strip()
    if not candidate:
        return None, None
    if "-" in candidate:
        start_text, end_text = (part.strip() for part in candidate.split("-", 1))
    else:
        start_text, end_text = candidate, ""
    start = start_text if parse_hhmm(start_text) is not None else None
    end = end_text if parse_hhmm(end_text) is not None else None
    return start, end


def safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except Exception:
        return int(default)


def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except Exception:
        return float(default)


def parse_hhmm(text: str | None) -> int | None:
    if not text:
        return None
    match = re.match(r"^(\d{1,2}):(\d{2})$", text.strip())
    if not match:
        return None
    hours, minutes = int(match.group(1)), int(match.group(2))
    if not (0 <= hours <= 23 and 0 <= minutes <= 59):
        return None
    return hours * 60 + minutes


def minutes_from_times(start: str | None, end: str | None) -> int | None:
    begin = parse_hhmm(start)
    finish = parse_hhmm(end)
    if begin is None or finish is None:
        return None
    if finish < begin:
        finish += 24 * 60
    return max(0, finish - begin)


def format_minutes(value: int | None) -> str:
    if value is None:
        return ""
    return str(int(value))


def human_datetime() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def ensure_directory(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


class Lesson:
    """Données métier d’une leçon."""

    def __init__(self, *, lesson_id: str | None = None) -> None:
        self.lesson_id = lesson_id or str(uuid.uuid4())
        self.title: str = ""
        self.date: str = ""
        self.start: str | None = None
        self.end: str | None = None
        self.minutes: int = 0
        self.journal: str = ""
        self.audio_path: str | None = None
        self.journal_path: str | None = None
        self.transcript_path: str | None = None
        self.created_at: str = human_datetime()
        self.updated_at: str = human_datetime()
        # Champs temporaires utilisés pendant l’édition (non sauvegardés)
        self.audio_source_path: str | None = None
        self.audio_source_is_temp: bool = False
        self.audio_removed: bool = False

    # ------------------------------------------------------------------
    # Sérialisation
    # ------------------------------------------------------------------

    def as_dict(self) -> dict[str, Any]:
        return {
            "lesson_id": self.lesson_id,
            "title": self.title,
            "date": self.date,
            "start": self.start,
            "end": self.end,
            "minutes": self.minutes,
            "audio_path": self.audio_path,
            "journal_path": self.journal_path,
            "transcript_path": self.transcript_path,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
        }

    @classmethod
    def from_payload(cls, payload: dict[str, Any], *, journal: str = "") -> "Lesson":
        lesson = cls(lesson_id=payload.get("lesson_id"))
        lesson.title = str(payload.get("title") or "").strip()
        lesson.date = str(payload.get("date") or "").strip()
        lesson.start = payload.get("start") or None
        lesson.end = payload.get("end") or None
        lesson.minutes = safe_int(payload.get("minutes") or 0)
        lesson.audio_path = payload.get("audio_path") or None
        lesson.journal_path = payload.get("journal_path") or None
        lesson.transcript_path = payload.get("transcript_path") or None
        lesson.created_at = payload.get("created_at") or lesson.created_at
        lesson.updated_at = payload.get("updated_at") or lesson.updated_at
        lesson.journal = journal or ""
        lesson.audio_source_path = None
        lesson.audio_source_is_temp = False
        lesson.audio_removed = False
        return lesson


# ---------------------------------------------------------------------------
# Dialogue d’édition de leçon
# ---------------------------------------------------------------------------


class LessonDialog(tk.Toplevel):
    """Fenêtre modale pour créer ou éditer une leçon."""

    def __init__(self, master: tk.Misc, initial: Lesson | None = None):
        super().__init__(master)
        self.title("Leçon")
        self.resizable(True, True)
        self.transient(master)
        self.grab_set()
        self.result: Lesson | None = None

        initial = initial or Lesson()

        self.lesson_id = initial.lesson_id
        self.initial_audio_path = initial.audio_path or ""
        self.audio_source_path: str | None = None
        self.audio_cleared = False
        self._audio_source_is_temp = False

        self.var_title = tk.StringVar(value=initial.title)
        self.var_date = tk.StringVar(value=initial.date)
        self.var_start = tk.StringVar(value=initial.start or "")
        self.var_end = tk.StringVar(value=initial.end or "")
        self.var_minutes = tk.StringVar(value=format_minutes(initial.minutes))

        audio_display = self.initial_audio_path or "Aucun fichier audio"
        self.var_audio_display = tk.StringVar(value=audio_display)
        self.var_record_button = tk.StringVar(value="Enregistrer…")
        self.var_record_status = tk.StringVar(value="Aucun enregistrement en cours.")

        self._recording_thread: threading.Thread | None = None
        self._recording_stop_event: threading.Event | None = None
        self._recording_finished_event: threading.Event | None = None
        self._recording_frames: list[Any] = []
        self._recording_error: Exception | None = None
        self._recording_active = False
        self._recording_start_time: float | None = None
        self._recording_samplerate = 44100
        self._recording_channels = 1
        self._recording_poll_job: str | None = None
        self._record_timer_job: str | None = None
        self._recording_status_message: str | None = None
        self._temp_recordings: set[str] = set()

        container = ttk.Frame(self, padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(6, weight=1)

        row = 0
        ttk.Label(container, text="Titre de la leçon").grid(row=row, column=0, sticky="w")
        ttk.Entry(container, textvariable=self.var_title).grid(row=row, column=1, sticky="we", padx=6, pady=2)

        row += 1
        ttk.Label(container, text="Date (YYYY-MM-DD)").grid(row=row, column=0, sticky="w")
        ttk.Entry(container, textvariable=self.var_date, width=18).grid(row=row, column=1, sticky="w", padx=6, pady=2)

        row += 1
        ttk.Label(container, text="Début (HH:MM)").grid(row=row, column=0, sticky="w")
        entry_start = ttk.Entry(container, textvariable=self.var_start, width=10)
        entry_start.grid(row=row, column=1, sticky="w", padx=6, pady=2)

        row += 1
        ttk.Label(container, text="Fin (HH:MM)").grid(row=row, column=0, sticky="w")
        entry_end = ttk.Entry(container, textvariable=self.var_end, width=10)
        entry_end.grid(row=row, column=1, sticky="w", padx=6, pady=2)

        row += 1
        ttk.Label(container, text="Durée (minutes)").grid(row=row, column=0, sticky="w")
        entry_minutes = ttk.Entry(container, textvariable=self.var_minutes, width=10)
        entry_minutes.grid(row=row, column=1, sticky="w", padx=6, pady=2)

        def refresh_minutes(*_args: Any) -> None:
            computed = minutes_from_times(self.var_start.get(), self.var_end.get())
            if computed is not None:
                self.var_minutes.set(str(int(computed)))

        entry_start.bind("<FocusOut>", refresh_minutes)
        entry_end.bind("<FocusOut>", refresh_minutes)

        row += 1
        ttk.Label(container, text="Journal").grid(row=row, column=0, sticky="nw", pady=(6, 0))
        self.txt_journal = tk.Text(container, wrap="word", height=12)
        self.txt_journal.grid(row=row, column=1, sticky="nsew", padx=6, pady=(2, 0))
        self.txt_journal.insert("1.0", initial.journal or "")

        row += 1
        audio_box = ttk.LabelFrame(container, text="Audio (optionnel)")
        audio_box.grid(row=row, column=0, columnspan=2, sticky="we", padx=0, pady=(10, 4))
        audio_box.columnconfigure(0, weight=1)
        audio_box.columnconfigure(1, weight=0)
        audio_box.columnconfigure(2, weight=0)

        ttk.Label(audio_box, textvariable=self.var_audio_display, wraplength=380).grid(
            row=0, column=0, columnspan=3, sticky="we", padx=4, pady=(4, 2)
        )
        ttk.Button(audio_box, text="Choisir un fichier…", command=self.select_audio_file).grid(
            row=1, column=0, sticky="w", padx=4, pady=(0, 4)
        )
        ttk.Button(audio_box, text="Effacer", command=self.clear_audio_file).grid(
            row=1, column=1, sticky="w", padx=4, pady=(0, 4)
        )
        ttk.Button(audio_box, textvariable=self.var_record_button, command=self.toggle_audio_recording).grid(
            row=1, column=2, sticky="e", padx=4, pady=(0, 4)
        )
        ttk.Label(
            audio_box,
            textvariable=self.var_record_status,
            wraplength=380,
            foreground="#555555",
        ).grid(row=2, column=0, columnspan=3, sticky="we", padx=4, pady=(0, 4))

        buttons = ttk.Frame(container)
        buttons.grid(row=row + 1, column=0, columnspan=2, sticky="e", pady=(6, 0))
        ttk.Button(buttons, text="Annuler", command=self._on_cancel).grid(row=0, column=0, padx=6)
        ttk.Button(buttons, text="Enregistrer", command=self.on_save).grid(row=0, column=1, padx=6)

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.wait_visibility()
        self.focus_set()

    # ------------------------------------------------------------------
    # Gestion audio (sélection, effacement, enregistrement)
    # ------------------------------------------------------------------

    def select_audio_file(self) -> None:
        self.stop_audio_recording(keep_result=False, show_message=False)
        path = filedialog.askopenfilename(
            parent=self,
            title="Sélectionner un fichier audio",
            filetypes=[
                ("Fichiers audio", "*.wav *.mp3 *.m4a *.aac *.flac *.ogg"),
                ("Tous les fichiers", "*.*"),
            ],
        )
        if not path:
            return
        if self._audio_source_is_temp and self.audio_source_path:
            self._discard_temp_recording(self.audio_source_path)
        self.audio_source_path = path
        self.audio_cleared = False
        self._audio_source_is_temp = False
        self.var_audio_display.set(path)

    def clear_audio_file(self) -> None:
        self.stop_audio_recording(keep_result=False, show_message=False)
        if self._audio_source_is_temp and self.audio_source_path:
            self._discard_temp_recording(self.audio_source_path)
        self.audio_source_path = None
        self.audio_cleared = True
        self.initial_audio_path = ""
        self._audio_source_is_temp = False
        self.var_audio_display.set("Aucun fichier audio")

    def toggle_audio_recording(self) -> None:
        if self._recording_active:
            self.stop_audio_recording()
        else:
            self.start_audio_recording()

    def start_audio_recording(self) -> None:
        if self._recording_active:
            return
        if _sd is None or _np is None:
            messagebox.showwarning(
                "Audio",
                "L’enregistrement nécessite les bibliothèques 'sounddevice' et 'numpy'.\n"
                "Installe-les avec `pip install sounddevice numpy`.",
                parent=self,
            )
            return

        if self._audio_source_is_temp and self.audio_source_path:
            self._discard_temp_recording(self.audio_source_path)
            self.audio_source_path = None
            self._audio_source_is_temp = False
            self.var_audio_display.set("Aucun fichier audio")

        self._recording_stop_event = threading.Event()
        self._recording_finished_event = threading.Event()
        self._recording_frames = []
        self._recording_error = None
        self._recording_active = True
        self._recording_start_time = time.time()
        self._recording_status_message = None

        samplerate = 44100
        try:
            default_input = getattr(_sd.default, "device", (None, None))[0]
            device_info = _sd.query_devices(default_input, "input")
        except Exception:
            try:
                device_info = _sd.query_devices(None, "input")
            except Exception:
                device_info = {}
        if isinstance(device_info, dict):
            default_samplerate = device_info.get("default_samplerate")
            if default_samplerate:
                try:
                    samplerate = int(default_samplerate)
                except (TypeError, ValueError):
                    samplerate = 44100
        if samplerate <= 0:
            samplerate = 44100
        self._recording_samplerate = samplerate

        self.var_record_button.set("Arrêter")
        self.var_record_status.set("Enregistrement en cours… 0s")

        def worker() -> None:
            try:
                def callback(indata, frames, time_info, status):
                    if status and str(status).strip():
                        self._recording_status_message = str(status)
                    self._recording_frames.append(indata.copy())
                    if self._recording_stop_event and self._recording_stop_event.is_set():
                        raise _sd.CallbackStop()

                with _sd.InputStream(
                    samplerate=self._recording_samplerate,
                    channels=self._recording_channels,
                    dtype="float32",
                    callback=callback,
                ):
                    if self._recording_stop_event:
                        self._recording_stop_event.wait()
            except Exception as exc:  # pragma: no cover - dépend matériel audio
                self._recording_error = exc
            finally:
                if self._recording_finished_event:
                    self._recording_finished_event.set()

        self._recording_thread = threading.Thread(target=worker, daemon=True)
        self._recording_thread.start()
        self._record_timer_job = self.after(200, self._update_recording_timer)
        self._recording_poll_job = self.after(300, self._poll_recording_error)

    def stop_audio_recording(
        self,
        *,
        keep_result: bool = True,
        show_message: bool = True,
    ) -> None:
        if not self._recording_active and not self._recording_thread:
            return
        if self._recording_stop_event and not self._recording_stop_event.is_set():
            self._recording_stop_event.set()
        if self._recording_finished_event:
            self._recording_finished_event.wait(timeout=5)
        if self._recording_thread and self._recording_thread.is_alive():
            self._recording_thread.join(timeout=5)
        self._recording_thread = None
        self._recording_active = False

        if self._record_timer_job:
            try:
                self.after_cancel(self._record_timer_job)
            except Exception:
                pass
            self._record_timer_job = None
        if self._recording_poll_job:
            try:
                self.after_cancel(self._recording_poll_job)
            except Exception:
                pass
            self._recording_poll_job = None

        self.var_record_button.set("Enregistrer…")

        error = self._recording_error
        self._recording_error = None

        if not keep_result:
            self._recording_frames = []
            if error and show_message:
                messagebox.showwarning("Audio", f"Enregistrement interrompu : {error}", parent=self)
            self.var_record_status.set("Enregistrement annulé.")
            return

        if error:
            self._recording_frames = []
            self.var_record_status.set("Erreur lors de l’enregistrement.")
            if show_message:
                messagebox.showwarning("Audio", f"Erreur lors de l’enregistrement : {error}", parent=self)
            return

        frames = list(self._recording_frames)
        self._recording_frames = []
        if not frames:
            self.var_record_status.set("Aucune donnée audio capturée.")
            if show_message:
                messagebox.showwarning("Audio", "Aucune donnée audio capturée.", parent=self)
            return

        if _np is None:
            self.var_record_status.set("Modules audio indisponibles.")
            return

        data = _np.concatenate(frames, axis=0)
        scaled = _np.int16(data * 32767)
        temp_dir = Path(self._create_temp_directory())
        ensure_directory(temp_dir)
        filename = temp_dir / f"lesson_{self.lesson_id}.wav"
        with wave.open(str(filename), "wb") as wav_file:
            wav_file.setnchannels(self._recording_channels)
            wav_file.setsampwidth(2)
            wav_file.setframerate(self._recording_samplerate)
            wav_file.writeframes(scaled.tobytes())

        self.audio_source_path = str(filename)
        self._temp_recordings.add(self.audio_source_path)
        self._audio_source_is_temp = True
        self.audio_cleared = False
        self.var_audio_display.set(self.audio_source_path)

        status = "Enregistrement terminé."
        if self._recording_status_message:
            status = self._recording_status_message
        self.var_record_status.set(status)

    def _update_recording_timer(self) -> None:
        if not self._recording_active or self._recording_start_time is None:
            return
        elapsed = int(time.time() - self._recording_start_time)
        self.var_record_status.set(f"Enregistrement en cours… {elapsed}s")
        self._record_timer_job = self.after(200, self._update_recording_timer)

    def _poll_recording_error(self) -> None:
        if self._recording_error:
            self.stop_audio_recording(keep_result=False, show_message=True)
            return
        self._recording_poll_job = self.after(300, self._poll_recording_error)

    def _discard_temp_recording(self, path: str) -> None:
        try:
            if os.path.isfile(path):
                os.remove(path)
        except OSError:
            pass
        self._temp_recordings.discard(path)

    def _create_temp_directory(self) -> str:
        base = Path(Path.home(), ".lesson_scribe", "recordings")
        ensure_directory(base)
        return str(base)

    def destroy(self) -> None:  # type: ignore[override]
        self.stop_audio_recording(keep_result=False, show_message=False)
        for recording in list(self._temp_recordings):
            self._discard_temp_recording(recording)
        super().destroy()

    def _on_cancel(self) -> None:
        self.result = None
        self.destroy()

    def on_save(self) -> None:
        lesson = Lesson(lesson_id=self.lesson_id)
        lesson.title = self.var_title.get().strip()
        lesson.date = self.var_date.get().strip()
        lesson.start = self.var_start.get().strip() or None
        lesson.end = self.var_end.get().strip() or None
        lesson.minutes = safe_int(self.var_minutes.get() or 0)
        lesson.journal = self.txt_journal.get("1.0", "end").strip()
        lesson.audio_path = self.initial_audio_path or None
        lesson.updated_at = human_datetime()

        # Stocker les métadonnées temporaires sur l'instance pour l’appelant.
        lesson.audio_source_path = self.audio_source_path  # type: ignore[attr-defined]
        lesson.audio_removed = self.audio_cleared  # type: ignore[attr-defined]
        lesson.audio_source_is_temp = self._audio_source_is_temp  # type: ignore[attr-defined]

        self.result = lesson
        self.destroy()


# ---------------------------------------------------------------------------
# Application principale
# ---------------------------------------------------------------------------


class LessonScribeApp(tk.Tk):
    """Fenêtre principale de Lesson Scribe."""

    def __init__(self) -> None:
        _load_env_from_config()
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1280x800")
        self.minsize(1100, 720)

        self.lessons: list[Lesson] = []
        self.current_index: int | None = None
        self.current_workspace: Path | None = None
        self.metadata_path: Path | None = None
        self.loading = False

        self._autosave_in_progress = False
        self._autosave_pending = False
        self._autosave_pending_force = False
        self._autosave_pending_show_info = False

        self._transcription_jobs: dict[str, dict[str, Any]] = {}
        self._transcription_queue: "Queue[tuple[str, str, Any]]" = Queue()
        self._active_transcription: dict[str, Any] | None = None

        default_title = os.environ.get("LESSON_SCRIBE_DEFAULT_TITLE") or os.environ.get("LESSON_DEFAULT_TITLE")
        self.default_title_template = default_title.strip() if default_title and default_title.strip() else None

        raw_time = os.environ.get("LESSON_SCRIBE_DEFAULT_TIME") or os.environ.get("LESSON_DEFAULT_TIME")
        start_time, end_time = _parse_default_time_range(raw_time)
        self.default_start_time = start_time
        self.default_end_time = end_time

        self.default_workspace_var: str | None = None
        self.default_workspace: Path | None = None
        for var_name in ("LESSON_SCRIBE_DEFAULT_WORKSPACE", "LESSON_DEFAULT_WORKSPACE"):
            raw_value = os.environ.get(var_name)
            if raw_value:
                self.default_workspace_var = var_name
                self.default_workspace = Path(os.path.expanduser(raw_value.strip()))
                break
        self._default_workspace_checked = False

        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self._build_layout()
        self.select_lesson(None)
        self.after(100, self.ensure_initial_workspace)
        self.after(200, self._poll_transcription_queue)

    def _build_default_lesson(self) -> Lesson:
        lesson = Lesson()
        if self.default_start_time:
            lesson.start = self.default_start_time
        if self.default_end_time:
            lesson.end = self.default_end_time
        if lesson.start and lesson.end:
            computed = minutes_from_times(lesson.start, lesson.end)
            if computed is not None:
                lesson.minutes = computed
        return lesson

    def _apply_lesson_defaults(self, lesson: Lesson) -> Lesson:
        if not lesson.title and self.default_title_template:
            date_label = lesson.date or datetime.date.today().isoformat()
            lesson.title = f"{self.default_title_template} {date_label}".strip()
        if lesson.minutes <= 0 and lesson.start and lesson.end:
            computed = minutes_from_times(lesson.start, lesson.end)
            if computed is not None:
                lesson.minutes = computed
        return lesson

    # ------------------------------------------------------------------
    # Construction UI
    # ------------------------------------------------------------------

    def _build_layout(self) -> None:
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)
        self.columnconfigure(2, weight=0)

        self._build_menubar()
        self._build_sidebar()
        self._build_center()
        self._build_help_panel()

    def _build_menubar(self) -> None:
        menubar = tk.Menu(self)

        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Ouvrir un workspace…", command=self.on_open_workspace)
        filemenu.add_command(label="Exporter le workspace…", command=self.on_export_workspace)
        filemenu.add_separator()
        filemenu.add_command(label="Quitter", command=self.destroy)
        menubar.add_cascade(label="Fichier", menu=filemenu)

        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="À propos", command=self.on_about)
        menubar.add_cascade(label="Aide", menu=helpmenu)

        self.config(menu=menubar)

    def _build_sidebar(self) -> None:
        frame = ttk.Frame(self, padding=8)
        frame.grid(row=0, column=0, sticky="ns")
        frame.rowconfigure(1, weight=1)

        ttk.Label(frame, text="Leçons", font=("TkDefaultFont", 11, "bold")).grid(row=0, column=0, sticky="w")
        self.lesson_list = tk.Listbox(frame, width=36, height=28, exportselection=False)
        self.lesson_list.grid(row=1, column=0, sticky="nswe", pady=(6, 6))
        self.lesson_list.bind("<<ListboxSelect>>", self.on_select_list)

        btns = ttk.Frame(frame)
        btns.grid(row=2, column=0, sticky="we", pady=(6, 0))
        ttk.Button(btns, text="Ajouter", command=self.on_add_lesson).grid(row=0, column=0, padx=2)
        ttk.Button(btns, text="Éditer", command=self.on_edit_lesson).grid(row=0, column=1, padx=2)
        ttk.Button(btns, text="Supprimer", command=self.on_delete_lesson).grid(row=0, column=2, padx=2)
        ttk.Button(btns, text="Copier prompt d’analyse", command=self.on_copy_analysis_prompt).grid(
            row=1, column=0, columnspan=3, pady=(6, 0)
        )

    def _build_center(self) -> None:
        frame = ttk.Frame(self, padding=8)
        frame.grid(row=0, column=1, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        info_frame = ttk.LabelFrame(frame, text="Détails de la leçon")
        info_frame.grid(row=0, column=0, sticky="we")
        for col in range(4):
            info_frame.columnconfigure(col, weight=1)

        self.var_detail_title = tk.StringVar()
        self.var_detail_date = tk.StringVar()
        self.var_detail_time = tk.StringVar()
        self.var_detail_minutes = tk.StringVar()
        self.var_detail_audio = tk.StringVar()

        ttk.Label(info_frame, text="Titre").grid(row=0, column=0, sticky="w", padx=6, pady=(4, 0))
        ttk.Label(info_frame, textvariable=self.var_detail_title, font=("TkDefaultFont", 10, "bold")).grid(
            row=0, column=1, columnspan=3, sticky="w", padx=6, pady=(4, 0)
        )

        ttk.Label(info_frame, text="Date").grid(row=1, column=0, sticky="w", padx=6)
        ttk.Label(info_frame, textvariable=self.var_detail_date).grid(row=1, column=1, sticky="w", padx=6)

        ttk.Label(info_frame, text="Horaire").grid(row=1, column=2, sticky="w", padx=6)
        ttk.Label(info_frame, textvariable=self.var_detail_time).grid(row=1, column=3, sticky="w", padx=6)

        ttk.Label(info_frame, text="Durée").grid(row=2, column=0, sticky="w", padx=6, pady=(0, 4))
        ttk.Label(info_frame, textvariable=self.var_detail_minutes).grid(row=2, column=1, sticky="w", padx=6, pady=(0, 4))

        ttk.Label(info_frame, text="Audio").grid(row=2, column=2, sticky="w", padx=6, pady=(0, 4))
        ttk.Label(info_frame, textvariable=self.var_detail_audio).grid(row=2, column=3, sticky="w", padx=6, pady=(0, 4))

        journal_frame = ttk.LabelFrame(frame, text="Journal")
        journal_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 6))
        journal_frame.columnconfigure(0, weight=1)
        journal_frame.rowconfigure(0, weight=1)

        self.txt_journal_display = tk.Text(journal_frame, wrap="word", state="disabled", height=16)
        self.txt_journal_display.grid(row=0, column=0, sticky="nsew")

        journal_scroll = ttk.Scrollbar(journal_frame, orient="vertical", command=self.txt_journal_display.yview)
        journal_scroll.grid(row=0, column=1, sticky="ns")
        self.txt_journal_display.configure(yscrollcommand=journal_scroll.set)

        actions_frame = ttk.Frame(frame)
        actions_frame.grid(row=2, column=0, sticky="we")
        ttk.Button(actions_frame, text="Éditer la leçon", command=self.on_edit_lesson).grid(row=0, column=0, padx=2)
        ttk.Button(actions_frame, text="Ouvrir le dossier", command=self.on_open_lesson_folder).grid(row=0, column=1, padx=2)

        trans_frame = ttk.LabelFrame(frame, text="Transcription audio (Vibe)")
        trans_frame.grid(row=3, column=0, sticky="we", pady=(10, 0))
        trans_frame.columnconfigure(0, weight=1)
        trans_frame.columnconfigure(1, weight=1)
        trans_frame.columnconfigure(2, weight=0)

        self.transcription_contact_var = tk.StringVar(value="Aucune transcription en cours.")
        self.transcription_status_var = tk.StringVar(value="")
        self.transcription_progress_var = tk.DoubleVar(value=0.0)
        self.transcription_progressbar = ttk.Progressbar(
            trans_frame,
            orient="horizontal",
            mode="determinate",
            variable=self.transcription_progress_var,
            maximum=100.0,
        )
        self.transcription_progressbar.grid(row=0, column=0, columnspan=2, sticky="we", padx=4, pady=(4, 2))
        self.transcription_cancel_btn = ttk.Button(
            trans_frame,
            text="Annuler",
            state="disabled",
            command=self.on_cancel_transcription,
        )
        self.transcription_cancel_btn.grid(row=0, column=2, sticky="e", padx=4, pady=(4, 2))
        self.transcription_label = ttk.Label(trans_frame, textvariable=self.transcription_contact_var)
        self.transcription_label.grid(row=1, column=0, sticky="w", padx=4, pady=(0, 4))
        self.transcription_state_label = ttk.Label(
            trans_frame,
            textvariable=self.transcription_status_var,
            justify="right",
        )
        self.transcription_state_label.grid(row=1, column=1, columnspan=2, sticky="e", padx=4, pady=(0, 4))

    def _build_help_panel(self) -> None:
        frame = ttk.Frame(self, padding=8)
        frame.grid(row=0, column=2, sticky="ns")
        frame.rowconfigure(0, weight=1)
        ttk.Label(frame, text="Aide", font=("TkDefaultFont", 11, "bold")).grid(row=0, column=0, sticky="nw")
        text = tk.Text(frame, width=38, height=32, wrap="word", state="normal")
        text.grid(row=1, column=0, sticky="ns")
        text.insert("1.0", HELP_TEXT)
        text.configure(state="disabled")
        self.help_widget = text

    # ------------------------------------------------------------------
    # Gestion liste / sélection
    # ------------------------------------------------------------------

    def refresh_lesson_list(self) -> None:
        self.lesson_list.delete(0, "end")
        for lesson in self.lessons:
            label = self._format_lesson_label(lesson)
            self.lesson_list.insert("end", label)
        if self.current_index is not None and 0 <= self.current_index < len(self.lessons):
            try:
                self.lesson_list.select_set(self.current_index)
                self.lesson_list.see(self.current_index)
            except tk.TclError:
                pass

    def _format_lesson_label(self, lesson: Lesson) -> str:
        date = lesson.date or "?"
        title = lesson.title or lesson.lesson_id[:8]
        minutes = lesson.minutes
        minutes_part = f" – {minutes} min" if minutes else ""
        return f"{date} – {title}{minutes_part}"

    def on_select_list(self, _event: tk.Event[Any]) -> None:  # pragma: no cover - interaction utilisateur
        selection = self.lesson_list.curselection()
        if not selection:
            self.select_lesson(None)
            return
        self.select_lesson(int(selection[0]))

    def select_lesson(self, index: int | None) -> None:
        self.current_index = index
        if index is None or not (0 <= index < len(self.lessons)):
            self.var_detail_title.set("Aucune leçon sélectionnée")
            self.var_detail_date.set("")
            self.var_detail_time.set("")
            self.var_detail_minutes.set("")
            self.var_detail_audio.set("")
            self._set_journal_display("")
            return

        lesson = self.lessons[index]
        self.var_detail_title.set(lesson.title or "(Sans titre)")
        self.var_detail_date.set(lesson.date or "—")
        times = []
        if lesson.start:
            times.append(lesson.start)
        if lesson.end:
            times.append(lesson.end)
        self.var_detail_time.set(" → ".join(times) if times else "—")
        self.var_detail_minutes.set(f"{lesson.minutes} min" if lesson.minutes else "—")
        self.var_detail_audio.set(lesson.audio_path or "—")
        self._set_journal_display(lesson.journal)

    def _set_journal_display(self, text: str) -> None:
        self.txt_journal_display.configure(state="normal")
        self.txt_journal_display.delete("1.0", "end")
        self.txt_journal_display.insert("1.0", text or "")
        self.txt_journal_display.configure(state="disabled")

    # ------------------------------------------------------------------
    # Actions UI
    # ------------------------------------------------------------------

    def on_add_lesson(self) -> None:
        dialog = LessonDialog(self, initial=self._build_default_lesson())
        self.wait_window(dialog)
        result = dialog.result
        if not result:
            return
        result = self._apply_lesson_defaults(result)
        lesson = self._prepare_lesson_assets(result)
        self.lessons.append(lesson)
        self.lessons.sort(key=lambda l: (l.date or "", l.start or ""))
        self.refresh_lesson_list()
        self.select_lesson(self.lessons.index(lesson))
        self.autosave()
        self._start_transcription_if_needed(lesson)

    def on_edit_lesson(self) -> None:
        if self.current_index is None or not (0 <= self.current_index < len(self.lessons)):
            messagebox.showinfo("Leçon", "Sélectionne une leçon à éditer.", parent=self)
            return
        current = self.lessons[self.current_index]
        dialog = LessonDialog(self, initial=current)
        self.wait_window(dialog)
        result = dialog.result
        if not result:
            return
        result = self._apply_lesson_defaults(result)
        lesson = self._prepare_lesson_assets(result, existing=current)
        lesson.updated_at = human_datetime()
        self.lessons[self.current_index] = lesson
        self.lessons.sort(key=lambda l: (l.date or "", l.start or ""))
        new_index = self.lessons.index(lesson)
        self.refresh_lesson_list()
        self.select_lesson(new_index)
        self.autosave()
        self._start_transcription_if_needed(lesson)

    def on_delete_lesson(self) -> None:
        if self.current_index is None or not (0 <= self.current_index < len(self.lessons)):
            messagebox.showinfo("Leçon", "Sélectionne une leçon à supprimer.", parent=self)
            return
        lesson = self.lessons[self.current_index]
        if not messagebox.askyesno(
            "Suppression",
            f"Supprimer la leçon « {lesson.title or lesson.date or lesson.lesson_id[:8]} » ?",
            parent=self,
        ):
            return
        self._cancel_transcription_job(lesson.lesson_id)
        self.lessons.pop(self.current_index)
        self.refresh_lesson_list()
        self.select_lesson(None if not self.lessons else min(self.current_index, len(self.lessons) - 1))
        self.autosave()

    def on_copy_analysis_prompt(self) -> None:
        if not self.lessons:
            messagebox.showinfo("Copier", "Aucune leçon disponible.", parent=self)
            return
        entries = sorted(self.lessons, key=lambda l: (l.date or "", l.start or ""))
        lines: list[str] = [
            "Tu es mon coach d’apprentissage.",
            "Analyse ces leçons et propose-moi une synthèse et un plan de révision.",
            "—",
        ]
        for lesson in entries:
            header = f"Leçon {lesson.date or '?'} — {lesson.title or lesson.lesson_id[:8]}"
            duration = f" ({lesson.minutes} min)" if lesson.minutes else ""
            hours = []
            if lesson.start:
                hours.append(lesson.start)
            if lesson.end:
                hours.append(lesson.end)
            if hours:
                header += f" [{hours[0]} → {hours[1] if len(hours) > 1 else ''}]"
            header += duration
            lines.append(header)
            lines.append(lesson.journal or "(journal vide)")
            lines.append("---")
        prompt = "\n".join(lines)
        self.clipboard_clear()
        self.clipboard_append(prompt)
        messagebox.showinfo("Copié", "Prompt copié dans le presse-papiers.", parent=self)

    def on_open_lesson_folder(self) -> None:
        if self.current_workspace is None:
            messagebox.showinfo("Workspace", "Ouvre d’abord un workspace.", parent=self)
            return
        if self.current_index is None or not (0 <= self.current_index < len(self.lessons)):
            messagebox.showinfo("Leçon", "Sélectionne une leçon.", parent=self)
            return
        lesson = self.lessons[self.current_index]
        if not lesson.journal_path:
            messagebox.showinfo("Leçon", "Cette leçon n’a pas encore été sauvegardée.", parent=self)
            return
        try:
            rel_path = Path(lesson.journal_path).parent
            folder = (self.current_workspace / rel_path).resolve()
            if folder.exists():
                if os.name == "nt":
                    os.startfile(str(folder))  # type: ignore[attr-defined]
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", str(folder)])
                else:
                    subprocess.Popen(["xdg-open", str(folder)])
            else:
                messagebox.showwarning("Leçon", "Dossier introuvable.", parent=self)
        except Exception as exc:  # pragma: no cover - dépend plateforme
            messagebox.showwarning("Leçon", f"Impossible d’ouvrir le dossier : {exc}", parent=self)

    # ------------------------------------------------------------------
    # Workspace
    # ------------------------------------------------------------------

    def ensure_initial_workspace(self) -> None:
        if self.current_workspace:
            return
        if not self._default_workspace_checked:
            self._default_workspace_checked = True
            if self.default_workspace and self.default_workspace.is_dir():
                if self.load_workspace(self.default_workspace, show_info=True):
                    return
            elif self.default_workspace:
                messagebox.showwarning(
                    "Workspace",
                    (
                        "Le dossier indiqué via la variable d’environnement "
                        f"{self.default_workspace_var} est introuvable :\n{self.default_workspace}.\n\n"
                        "Sélectionne un workspace manuellement."
                    ),
                    parent=self,
                )
        answer = messagebox.askyesnocancel(
            "Workspace requis",
            "Pour éviter toute perte, ouvre un workspace existant (Oui) ou crée-en un nouveau (Non).\n"
            "Annuler fermera l’application.",
            parent=self,
        )
        if answer is None:
            self.destroy()
            return
        if answer:
            path = filedialog.askdirectory(title="Ouvrir un workspace Lesson Scribe", mustexist=True)
            if path and self.load_workspace(Path(path)):
                return
        else:
            path = filedialog.askdirectory(title="Choisir un dossier vide pour créer un workspace")
            if path:
                folder = Path(path)
                if folder.exists() and any(folder.iterdir()):
                    proceed = messagebox.askyesno(
                        "Dossier non vide",
                        "Le dossier sélectionné contient déjà des fichiers. Continuer ?",
                        parent=self,
                    )
                    if not proceed:
                        folder = None  # type: ignore[assignment]
                if folder:
                    try:
                        ensure_directory(Path(path))
                    except Exception as exc:
                        messagebox.showerror("Workspace", f"Impossible de préparer le dossier : {exc}", parent=self)
                    else:
                        self.lessons = []
                        self.current_index = None
                        self.refresh_lesson_list()
                        if self.save_workspace(Path(path), show_message=False):
                            messagebox.showinfo("Workspace", f"Nouveau workspace créé :\n{path}", parent=self)
                            return
        self.after(200, self.ensure_initial_workspace)

    def on_open_workspace(self) -> None:
        path = filedialog.askdirectory(title="Ouvrir un workspace Lesson Scribe", mustexist=True)
        if path:
            self.load_workspace(Path(path))

    def load_workspace(self, folder: Path, *, show_info: bool = False) -> bool:
        workspace = folder.resolve()
        metadata_path = workspace / "metadata.json"
        if not workspace.is_dir():
            messagebox.showerror("Workspace", "Le dossier sélectionné est introuvable.", parent=self)
            return False
        if not metadata_path.is_file():
            messagebox.showerror("Workspace", "metadata.json est introuvable dans ce dossier.", parent=self)
            return False
        try:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        except Exception as exc:
            messagebox.showerror("Workspace", f"Lecture impossible : {exc}", parent=self)
            return False

        lessons_entries = metadata.get("lessons", [])
        if lessons_entries and not isinstance(lessons_entries, list):
            messagebox.showerror("Workspace", "Le champ 'lessons' doit être une liste.", parent=self)
            return False

        loaded: list[Lesson] = []
        warnings: list[str] = []
        for entry in lessons_entries:
            if not isinstance(entry, dict):
                warnings.append("Entrée de leçon invalide dans metadata.json")
                continue
            lesson_path = entry.get("lesson_path")
            if not lesson_path:
                warnings.append("Entrée sans 'lesson_path'.")
                continue
            try:
                abs_path = (workspace / lesson_path).resolve()
            except Exception:
                warnings.append(f"Chemin invalide : {lesson_path!r}")
                continue
            if not abs_path.is_file():
                warnings.append(f"Fichier de leçon manquant : {lesson_path}")
                continue
            try:
                payload = json.loads(abs_path.read_text(encoding="utf-8"))
            except Exception as exc:
                warnings.append(f"Lecture impossible de {lesson_path} : {exc}")
                continue
            journal_text = ""
            journal_rel = payload.get("journal_path")
            if journal_rel:
                journal_file = (workspace / journal_rel).resolve()
                if journal_file.is_file():
                    try:
                        journal_text = journal_file.read_text(encoding="utf-8")
                    except Exception as exc:
                        warnings.append(f"Lecture impossible du journal {journal_rel} : {exc}")
                else:
                    warnings.append(f"Journal manquant : {journal_rel}")
            lesson = Lesson.from_payload(payload, journal=journal_text)
            lesson.journal_path = payload.get("journal_path")
            loaded.append(lesson)

        loaded.sort(key=lambda l: (l.date or "", l.start or ""))

        self.lessons = loaded
        self.current_workspace = workspace
        self.metadata_path = metadata_path
        self.refresh_lesson_list()
        self.select_lesson(0 if self.lessons else None)

        if warnings:
            messagebox.showwarning("Import partiel", "\n".join(warnings), parent=self)
        if show_info:
            messagebox.showinfo("Workspace", f"Workspace chargé :\n{workspace}", parent=self)
        return True

    def on_export_workspace(self) -> None:
        if not self.current_workspace:
            messagebox.showinfo("Workspace", "Aucun workspace chargé.", parent=self)
            return
        path = filedialog.askdirectory(title="Choisir un dossier de destination")
        if not path:
            return
        destination = Path(path).resolve()
        if destination == self.current_workspace:
            messagebox.showinfo("Export", "Sélectionne un dossier différent du workspace actuel.", parent=self)
            return
        try:
            ensure_directory(destination)
            archive_name = destination / f"lesson-scribe_{datetime.datetime.now():%Y%m%d_%H%M%S}"
            shutil.make_archive(str(archive_name), "zip", root_dir=self.current_workspace)
        except Exception as exc:
            messagebox.showerror("Export", f"Échec de l’export : {exc}", parent=self)
            return
        messagebox.showinfo("Export", f"Archive créée : {archive_name}.zip", parent=self)

    def save_workspace(self, folder: Path | None = None, *, show_message: bool = False) -> bool:
        if folder is not None:
            self.current_workspace = folder.resolve()
        if self.current_workspace is None:
            return False
        workspace = self.current_workspace
        ensure_directory(workspace)
        self.metadata_path = workspace / "metadata.json"
        metadata = self._build_metadata()
        try:
            self._write_lessons(metadata, workspace)
            self.metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as exc:
            messagebox.showerror("Sauvegarde", f"Impossible d’écrire le workspace : {exc}", parent=self)
            return False
        if show_message:
            messagebox.showinfo("Sauvegarde", f"Workspace sauvegardé :\n{workspace}", parent=self)
        return True

    def _build_metadata(self) -> dict[str, Any]:
        lessons_meta = []
        for lesson in self.lessons:
            lessons_meta.append(
                {
                    "lesson_id": lesson.lesson_id,
                    "lesson_path": f"lessons/{lesson.lesson_id}/lesson.json",
                }
            )
        return {
            "app": APP_NAME,
            "version": WORKSPACE_VERSION,
            "last_updated": human_datetime(),
            "lessons": lessons_meta,
        }

    def _write_lessons(self, metadata: dict[str, Any], workspace: Path) -> None:
        lessons_dir = workspace / "lessons"
        ensure_directory(lessons_dir)
        existing_dirs = {p.name for p in lessons_dir.iterdir() if p.is_dir()}
        expected_dirs: set[str] = set()
        for lesson in self.lessons:
            lesson_dir = lessons_dir / lesson.lesson_id
            ensure_directory(lesson_dir)
            expected_dirs.add(lesson.lesson_id)
            journal_rel = lesson.journal_path or f"lessons/{lesson.lesson_id}/journal.txt"
            journal_abs = workspace / journal_rel
            ensure_directory(journal_abs.parent)
            journal_abs.write_text(lesson.journal or "", encoding="utf-8")
            lesson.journal_path = journal_rel
            payload = lesson.as_dict()
            payload["journal_path"] = journal_rel
            if lesson.transcript_path:
                payload["transcript_path"] = lesson.transcript_path
            lesson_file = lesson_dir / "lesson.json"
            lesson_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        for name in existing_dirs - expected_dirs:
            shutil.rmtree(lessons_dir / name, ignore_errors=True)

    # ------------------------------------------------------------------
    # Autosauvegarde
    # ------------------------------------------------------------------

    def autosave(self, *, force: bool = False, show_info: bool = False) -> None:
        if self._autosave_in_progress:
            self._autosave_pending = True
            self._autosave_pending_force = self._autosave_pending_force or force
            self._autosave_pending_show_info = self._autosave_pending_show_info or show_info
            return
        if not self.current_workspace:
            if force:
                self.ensure_initial_workspace()
            return
        try:
            self._autosave_in_progress = True
            self.save_workspace(show_message=show_info)
        finally:
            self._autosave_in_progress = False
            if self._autosave_pending:
                pending_force = self._autosave_pending_force
                pending_show = self._autosave_pending_show_info
                self._autosave_pending = False
                self._autosave_pending_force = False
                self._autosave_pending_show_info = False
                self.after_idle(lambda: self.autosave(force=pending_force, show_info=pending_show))

    # ------------------------------------------------------------------
    # Préparation des données (journal, audio, transcription)
    # ------------------------------------------------------------------

    def _prepare_lesson_assets(self, lesson: Lesson, existing: Lesson | None = None) -> Lesson:
        workspace = self.current_workspace
        if workspace is None:
            return lesson
        workspace = workspace.resolve()
        lesson_dir = workspace / "lessons" / lesson.lesson_id
        ensure_directory(lesson_dir)

        lesson.journal_path = f"lessons/{lesson.lesson_id}/journal.txt"

        if existing is not None:
            lesson.created_at = existing.created_at
            if not lesson.transcript_path:
                lesson.transcript_path = existing.transcript_path

        audio_source_path: str | None = getattr(lesson, "audio_source_path", None)
        audio_removed: bool = getattr(lesson, "audio_removed", False)
        audio_source_is_temp: bool = getattr(lesson, "audio_source_is_temp", False)

        existing_audio_rel = existing.audio_path if existing else None

        if audio_removed:
            if existing_audio_rel:
                try:
                    audio_abs = (workspace / existing_audio_rel).resolve()
                except Exception:
                    audio_abs = None
                if audio_abs and audio_abs.is_file():
                    try:
                        audio_abs.unlink()
                    except OSError:
                        pass
            lesson.audio_path = None
            lesson.transcript_path = None
            self._cancel_transcription_job(lesson.lesson_id)
        elif audio_source_path:
            src_abs = Path(audio_source_path).resolve()
            if src_abs.is_file():
                try:
                    rel = src_abs.relative_to(workspace)
                except ValueError:
                    dest = lesson_dir / src_abs.name
                    counter = 1
                    while dest.exists():
                        dest = lesson_dir / f"{src_abs.stem}_{counter}{src_abs.suffix}"
                        counter += 1
                    shutil.copy2(src_abs, dest)
                    lesson.audio_path = f"lessons/{lesson.lesson_id}/{dest.name}"
                    if audio_source_is_temp:
                        try:
                            src_abs.unlink()
                        except OSError:
                            pass
                else:
                    lesson.audio_path = str(rel).replace("\\", "/")
            else:
                messagebox.showwarning("Audio", "Le fichier audio sélectionné est introuvable.", parent=self)
        else:
            lesson.audio_path = existing_audio_rel
            if existing and not lesson.transcript_path:
                lesson.transcript_path = existing.transcript_path

        return lesson

    # ------------------------------------------------------------------
    # Transcription (Vibe)
    # ------------------------------------------------------------------

    def on_cancel_transcription(self) -> None:
        job = self._active_transcription or {}
        lesson_id = job.get("lesson_id")
        if lesson_id:
            self._cancel_transcription_job(lesson_id, user_request=True)

    def _start_transcription_if_needed(self, lesson: Lesson) -> None:
        if not self.current_workspace or not lesson.audio_path:
            return
        transcript_rel = lesson.transcript_path or f"lessons/{lesson.lesson_id}/transcript.txt"
        lesson.transcript_path = transcript_rel
        transcript_abs = (self.current_workspace / transcript_rel).resolve()
        ensure_directory(transcript_abs.parent)
        audio_abs = (self.current_workspace / lesson.audio_path).resolve()
        journal_rel = lesson.journal_path or f"lessons/{lesson.lesson_id}/journal.txt"
        journal_abs = (self.current_workspace / journal_rel).resolve()
        ensure_directory(journal_abs.parent)
        if not journal_abs.exists():
            try:
                journal_abs.write_text(lesson.journal or "", encoding="utf-8")
            except Exception:
                pass
        job_started = self._start_transcription_job(
            lesson_id=lesson.lesson_id,
            display_label=self._format_lesson_label(lesson),
            audio_abs=str(audio_abs),
            transcript_rel=transcript_rel,
            transcript_abs=str(transcript_abs),
            journal_abs=str(journal_abs),
        )
        if not job_started:
            lesson.transcript_path = None

    def _resolve_vibe_executable(self) -> str:
        candidate = os.environ.get("VIBE_CLI", "vibe")
        candidate = os.path.expanduser(candidate)
        if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
            return candidate
        return candidate

    def _resolve_vibe_model(self) -> str | None:
        value = os.environ.get("VIBE_MODEL_PATH")
        if value:
            expanded = os.path.expanduser(value)
            return expanded if os.path.isfile(expanded) else None
        return None

    def _start_transcription_job(
        self,
        *,
        lesson_id: str,
        display_label: str,
        audio_abs: str,
        transcript_rel: str,
        transcript_abs: str,
        journal_abs: str,
    ) -> bool:
        vibe_exe = self._resolve_vibe_executable()
        vibe_model = self._resolve_vibe_model()
        if not vibe_model:
            messagebox.showwarning(
                "Vibe",
                "Modèle Whisper introuvable. Définis VIBE_MODEL_PATH vers un fichier .bin.",
                parent=self,
            )
            return False
        if not os.path.isfile(audio_abs):
            messagebox.showwarning("Audio", "Fichier audio introuvable pour la transcription.", parent=self)
            return False

        job = {
            "lesson_id": lesson_id,
            "display_label": display_label,
            "audio_abs": audio_abs,
            "transcript_rel": transcript_rel,
            "transcript_abs": transcript_abs,
            "journal_abs": journal_abs,
            "process": None,
            "cancel_event": threading.Event(),
        }

        worker = threading.Thread(target=self._run_transcription_job, args=(job,), daemon=True)
        worker.start()
        self._transcription_jobs[lesson_id] = job
        self._active_transcription = job
        self._update_transcription_panel(display_label, 0.0, "Initialisation de Vibe…", enable_cancel=True)
        return True

    def _run_transcription_job(self, job: dict[str, Any]) -> None:
        lesson_id = job["lesson_id"]
        display_label = job["display_label"]
        audio_abs = job["audio_abs"]
        transcript_abs = job["transcript_abs"]
        journal_abs = job["journal_abs"]

        vibe_exe = self._resolve_vibe_executable()
        vibe_model = self._resolve_vibe_model()
        if not vibe_model:
            self._transcription_queue.put(("error", lesson_id, "Modèle Vibe introuvable."))
            return

        env = os.environ.copy()
        args = [
            vibe_exe,
            "--model",
            vibe_model,
            "--file",
            audio_abs,
            "--language",
            os.environ.get("VIBE_LANGUAGE", "french"),
        ]
        temperature = os.environ.get("VIBE_TEMPERATURE")
        if temperature:
            args += ["--temperature", temperature]
        threads = os.environ.get("VIBE_THREADS")
        if threads:
            args += ["--threads", threads]

        try:
            process = subprocess.Popen(
                args,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                env=env,
            )
        except Exception as exc:
            self._transcription_queue.put(("error", lesson_id, f"Échec du démarrage de Vibe : {exc}"))
            return

        job["process"] = process
        self._transcription_queue.put(("start", lesson_id, display_label))

        try:
            for line in iter(process.stdout.readline, ""):
                if job["cancel_event"].is_set():
                    break
                line = line.strip()
                if not line:
                    continue
                match = re.search(r"(\d+)%", line)
                if match:
                    progress = safe_float(match.group(1), 0.0)
                    self._transcription_queue.put(("progress", lesson_id, progress))
            process.wait()
        finally:
            job["process"] = None

        if job["cancel_event"].is_set():
            self._transcription_queue.put(("cancelled", lesson_id, display_label))
            return

        if process.returncode != 0:
            tail = process.stdout.read().strip() if process.stdout else ""
            if tail:
                tail = tail.splitlines()[-1]
            self._transcription_queue.put(("error", lesson_id, tail or "Transcription échouée."))
            return

        try:
            content = Path(audio_abs).with_suffix(".txt").read_text(encoding="utf-8")
        except Exception as exc:
            self._transcription_queue.put(("error", lesson_id, f"Lecture transcription impossible : {exc}"))
            return
        Path(transcript_abs).write_text(content, encoding="utf-8")
        payload = {
            "lesson_id": lesson_id,
            "display_label": display_label,
            "transcript_abs": transcript_abs,
            "journal_abs": journal_abs,
            "content": content,
        }
        self._transcription_queue.put(("done", lesson_id, payload))

    def _cancel_transcription_job(self, lesson_id: str, user_request: bool = False) -> None:
        job = self._transcription_jobs.get(lesson_id)
        if not job:
            return
        process = job.get("process")
        job["cancel_event"].set()
        if process and process.poll() is None:
            try:
                process.terminate()
            except Exception:
                pass
        if user_request:
            self._transcription_queue.put(("cancel-requested", lesson_id, job.get("display_label")))

    def _poll_transcription_queue(self) -> None:
        try:
            while True:
                event = self._transcription_queue.get_nowait()
                self._handle_transcription_event(event)
        except Empty:
            pass
        finally:
            self.after(200, self._poll_transcription_queue)

    def _handle_transcription_event(self, event: tuple[str, str, Any]) -> None:
        kind, lesson_id, payload = event
        if kind == "start":
            self._update_transcription_panel(payload, 0.0, "Modèle en cours de chargement…", enable_cancel=True)
        elif kind == "progress":
            status = f"Progression : {payload:.0f}%"
            self._update_transcription_panel(None, payload, status, enable_cancel=True)
        elif kind == "cancel-requested":
            if self._active_transcription and self._active_transcription.get("lesson_id") == lesson_id:
                self._update_transcription_panel(payload, None, "Annulation en cours…", enable_cancel=False)
        elif kind == "cancelled":
            if self._active_transcription and self._active_transcription.get("lesson_id") == lesson_id:
                self._set_transcription_idle("Transcription annulée.")
            self._transcription_jobs.pop(lesson_id, None)
        elif kind == "error":
            if self._active_transcription and self._active_transcription.get("lesson_id") == lesson_id:
                self._set_transcription_idle("Erreur lors de la transcription.")
            messagebox.showerror("Vibe", str(payload) or "Erreur inconnue", parent=self)
            self._transcription_jobs.pop(lesson_id, None)
        elif kind == "done":
            self._apply_transcription_to_lesson(payload)
            if self._active_transcription and self._active_transcription.get("lesson_id") == lesson_id:
                self._set_transcription_idle("Transcription terminée.")
            self._transcription_jobs.pop(lesson_id, None)
        if not self._transcription_jobs:
            self._set_transcription_idle()

    def _update_transcription_panel(
        self,
        label: str | None,
        progress: float | None,
        status: str | None,
        *,
        enable_cancel: bool | None = None,
    ) -> None:
        if label is not None:
            self.transcription_contact_var.set(label)
        if progress is not None:
            self.transcription_progressbar.config(mode="determinate")
            self.transcription_progress_var.set(float(progress))
        if status is not None:
            self.transcription_status_var.set(status)
        if enable_cancel is not None:
            state = "normal" if enable_cancel else "disabled"
            self.transcription_cancel_btn.config(state=state)
        if progress is None:
            self.transcription_progressbar.config(mode="indeterminate")
            self.transcription_progressbar.start(80)
        else:
            self.transcription_progressbar.stop()

    def _set_transcription_idle(self, status: str = "") -> None:
        self._active_transcription = None
        self.transcription_progressbar.stop()
        self.transcription_progressbar.config(mode="determinate")
        self.transcription_progress_var.set(0.0)
        self.transcription_contact_var.set("Aucune transcription en cours.")
        self.transcription_status_var.set(status)
        self.transcription_cancel_btn.config(state="disabled")

    def _apply_transcription_to_lesson(self, payload: dict[str, Any]) -> None:
        lesson_id = payload.get("lesson_id")
        content = payload.get("content") or ""
        transcript_abs = payload.get("transcript_abs")
        journal_abs = payload.get("journal_abs")
        if not lesson_id:
            return
        lesson = next((l for l in self.lessons if l.lesson_id == lesson_id), None)
        if not lesson:
            return
        if not content.strip():
            messagebox.showwarning("Vibe", "La transcription générée est vide.", parent=self)
            return
        if transcript_abs and self.current_workspace:
            rel = os.path.relpath(transcript_abs, self.current_workspace).replace("\\", "/")
            lesson.transcript_path = rel
        if TRANSCRIPT_HEADER not in lesson.journal:
            lesson.journal = (lesson.journal or "").rstrip() + TRANSCRIPT_HEADER + content
        else:
            head, *_ = lesson.journal.split(TRANSCRIPT_HEADER, 1)
            lesson.journal = (head or "").rstrip() + TRANSCRIPT_HEADER + content
        if journal_abs:
            try:
                Path(journal_abs).write_text(lesson.journal, encoding="utf-8")
            except Exception as exc:
                messagebox.showwarning("Vibe", f"Impossible d’écrire le journal : {exc}", parent=self)
        if transcript_abs:
            try:
                Path(transcript_abs).write_text(content, encoding="utf-8")
            except Exception as exc:
                messagebox.showwarning("Vibe", f"Impossible d’écrire la transcription : {exc}", parent=self)
        self.refresh_lesson_list()
        if self.current_index is not None and 0 <= self.current_index < len(self.lessons):
            self.select_lesson(self.current_index)
        self.autosave()

    # ------------------------------------------------------------------
    # Divers
    # ------------------------------------------------------------------

    def on_about(self) -> None:
        messagebox.showinfo(
            "À propos",
            f"{APP_NAME}\nJournal de leçons avec audio et transcription.",
            parent=self,
        )

    def on_close(self) -> None:
        if messagebox.askokcancel("Quitter", "Fermer Lesson Scribe ?", parent=self):
            self.destroy()


def main() -> None:
    _load_env_from_config()
    app = LessonScribeApp()
    app.mainloop()


__all__ = ["LessonScribeApp", "main"]


if __name__ == "__main__":
    main()
