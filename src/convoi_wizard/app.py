#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Convoi Wizard v3 – LOG-FIRST (journal libre) + ANALYSE VIA PATCH
===============================================================

Objectif:
- Tu saisis des contacts avec uniquement: date, heure début/fin (→ minutes auto), et un GRAND TEXTE "journal".
- Tu copies un "Prompt d'analyse" (TOUS les logs) → ChatGPT renvoie un PATCH JSON d'analyse.
- Tu importes ce patch → le soft met à jour TOUT: scores, besoins, need_support, et même des enrichissements par contact.

Schéma JSON exporté (global):
{
  "title": "...", "owner": "...", "last_updated": "...Z",
  "people": [
    {
      "name": "...", "circle": "Noyau|Médiane|Externe",
      "intimacy": 0..10, "reliability": 0..10, "valence": -2..+2,
      "energy_cost": 0..2, "match_besoins": 0..2,
      "needs_nourished": [...],
      "need_support": { "toucher": {"minutes_per_week": 30}, "parler_profond": {"occ_per_week":1} },
      "contacts": [
        {
          "contact_id": "uuid",
          "date":"YYYY-MM-DD","start":"HH:MM","end":"HH:MM","minutes": int,
          "journal": "ton récit libre",
          // champs facultatifs ajoutés PAR ANALYSE (patch):
          // "channel": "presentiel|audio|visio|texte",
          // "needs": [...],
          // "valence": -2..+2,
          // "mood_after": 0..10,
          // "format": "...",
          // "note": "...",
          // "tags": ["ambivalent?","conflit réparé?",...]
        }
      ],
      "ambivalent": bool, "status": "", "notes": ""
    }
  ]
}

Schéma PATCH d'analyse attendu (à importer):
{
  "updates": {
    "intimacy": float,
    "reliability": float,
    "valence": float,
    "energy_cost": float,
    "match_besoins": float,
    "needs_nourished": [ ... ],
    "need_support": { "<need>": {"minutes_per_week": float?, "occ_per_week": float?}, ... },
    "ambivalent": true|false,
    "status": "..."
  },
  "notes_append": "texte à ajouter aux notes",
  "plan": ["rituel 1", "rituel 2", ...],
  "questions": ["question 1", "question 2", ...],
  "contacts_updates": [
    {
      "contact_id": "...",
      // champs facultatifs inférés par ChatGPT à reporter dans le contact correspondant
      "channel": "presentiel|audio|visio|texte",
      "needs": [ ... ],
      "valence": float,
      "mood_after": float,
      "format": "string",
      "note": "string",
      "tags": [ ... ]
    },
    ...
  ]
}

Remarque:
- Le UI n'impose plus valence/besoins/format/etc. → tout peut être inféré/ajouté par le patch d'analyse.
- Tu peux évidemment continuer à ajuster manuellement les sliders globaux si tu le souhaites; le patch pourra les écraser ensuite.

"""

import json, os, datetime, re, uuid, subprocess, sys, shutil, copy, threading, time, tempfile, wave, importlib
from collections import deque
from queue import Queue, Empty
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from dotenv import load_dotenv


def _load_env_from_config():
    env_path = os.environ.get("CONVOI_ENV_FILE")
    if env_path:
        load_dotenv(env_path, override=False)
    else:
        load_dotenv(override=False)


_load_env_from_config()

# Audio helpers --------------------------------------------------------------

def _load_audio_modules():
    sounddevice = importlib.import_module("sounddevice")
    numpy = importlib.import_module("numpy")
    return sounddevice, numpy

# --------------------------- Constantes & aides ---------------------------

DEFAULT_NEEDS = [
    "apaisement", "toucher", "parler_profond",
    "fun", "co_projet", "anglais", "danse"
]

WORKSPACE_VERSION = 1

ACTION_COLOR = {
    "nourrir": "#2ca02c",            # vert
    "renegocier": "#ff7f0e",         # orange
    "reclasser/espacer": "#d62728"   # rouge
}

TRANSCRIPT_HEADER = "\n\n--- Transcription Vibe ---\n"

HELP_TEXT = {
    "name": ("Prénom / Nom court",
             "Prénom (ou pseudo) facile à reconnaître."),
    "circle": ("Cercle HMT",
               "Place selon PROXIMITÉ ÉMOTIONNELLE.\n"
               "• Noyau: 2–5 personnes très proches\n"
               "• Médiane: proches réguliers\n"
               "• Externe: liens faibles/occasionnels"),
    "intimacy": ("Intimité (0–10)",
                 "À quel point tu te sens intime/serein avec cette personne ?"),
    "reliability": ("Fiabilité (0–10)",
                    "Disponible quand ça compte ? Tient ses engagements ?"),
    "valence": ("Valence moyenne (−2..+2)",
                "Comment tu te sens APRÈS en général ? −2 vidé → +2 rechargé"),
    "energy_cost": ("Énergie_coût (0..2)",
                    "0 recharge • 1 neutre • 2 vide souvent"),
    "match_besoins": ("Match_besoins (0..2)",
                      "À quel point cette personne apporte EXACTEMENT ce que tu cherches ?"),
    "needs_nourished": ("Besoins nourris (tags)",
                        "Coche ce que la relation nourrit souvent (ou laisse vide si tu préfères déléguer à l'analyse)."),
    "contacts": ("Journal des contacts",
                 "Ajoute un contact avec Date + Heures + Journal libre (grand texte)."),
    "notes": ("Notes libres",
              "Infos libres (préférences tactiles, formats, etc.).")
}

def safe_float(x, default=0.0):
    try:
        return float(x)
    except Exception:
        return float(default)

def clamp(v, lo, hi):
    v = safe_float(v, lo)
    return max(lo, min(hi, v))

def circle_norm(c):
    c = (c or "").strip().lower()
    if c.startswith("noy"): return "Noyau"
    if c.startswith("méd") or c.startswith("med"): return "Médiane"
    return "Externe"

def map_valence_to_0_10(v):
    try:
        return max(0.0, min(10.0, (float(v) + 2.0) * 2.5))
    except Exception:
        return 5.0

def compute_priority(p):
    # Preview local (l'analyse peut écraser ensuite)
    intim = clamp(p.get("intimacy", 0), 0, 10)
    fiab  = clamp(p.get("reliability", 0), 0, 10)
    val_n = map_valence_to_0_10(clamp(p.get("valence", 0), -2, 2))
    match = clamp(p.get("match_besoins", 0), 0, 2)
    energy= clamp(p.get("energy_cost", 0), 0, 2)
    score = 0.35*intim + 0.35*fiab + 0.20*val_n + 0.10*(match*5.0) - 2.0*energy
    return round(max(0.0, min(10.0, score)), 2)

def label_action(score):
    if score >= 7.5: return "nourrir"
    if score >= 5.0: return "renegocier"
    return "reclasser/espacer"

def parse_date_iso(s):
    try:
        return datetime.date.fromisoformat(s)
    except Exception:
        return None

def parse_hhmm(s):
    if not s: return None
    m = re.match(r"^(\d{1,2}):(\d{2})$", str(s).strip())
    if not m: return None
    h, mi = int(m.group(1)), int(m.group(2))
    if h<0 or h>23 or mi<0 or mi>59: return None
    return h*60 + mi

def minutes_from_times(start_hhmm, end_hhmm):
    a, b = parse_hhmm(start_hhmm), parse_hhmm(end_hhmm)
    if a is None or b is None: return None
    if b < a:  # passage minuit
        b += 24*60
    return max(0, b - a)

# --------------------------- Tooltips ---------------------------

class Tooltip:
    def __init__(self, widget, text, wrap=420):
        self.widget = widget
        self.text = text
        self.wrap = wrap
        self.tip = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        if self.tip: return
        x = self.widget.winfo_rootx()+20
        y = self.widget.winfo_rooty()+20
        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(tw, text=self.text, justify="left", relief="solid", borderwidth=1,
                       bg="#ffffe0", wraplength=self.wrap)
        lbl.pack(ipadx=6, ipady=4)

    def hide(self, event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None

# --------------------------- Dialog Contact (journal libre) ---------------------------

class ContactDialog(tk.Toplevel):
    def __init__(self, master, initial=None):
        super().__init__(master)
        self.title("Ajouter / Éditer un contact (journal)")
        self.resizable(True, True)
        self.result = None

        self.var_date  = tk.StringVar(value=(initial or {}).get("date",""))
        self.var_start = tk.StringVar(value=(initial or {}).get("start",""))
        self.var_end   = tk.StringVar(value=(initial or {}).get("end",""))
        init_min = (initial or {}).get("minutes","")
        self.var_min   = tk.StringVar(value=str(init_min) if init_min not in (None,"") else "")
        init_journal = (initial or {}).get("journal","")
        self.contact_id = (initial or {}).get("contact_id") or str(uuid.uuid4())
        self.initial_audio_path = (initial or {}).get("audio_path") or ""
        self.previous_audio_path = self.initial_audio_path
        self.audio_source_path = None
        self.audio_cleared = False
        self._audio_source_is_temp = False
        self._temp_recordings = set()
        self._temp_recordings_in_use = set()
        self._recording_thread = None
        self._recording_stop_event = None
        self._recording_finished_event = None
        self._recording_frames = []
        self._recording_error = None
        self._recording_modules = None
        self._recording_poll_job = None
        self._record_timer_job = None
        self._recording_active = False
        self._recording_start_time = None
        self._recording_samplerate = 44100
        self._recording_channels = 1
        self._recording_status_message = None
        audio_display = self.initial_audio_path or "Aucun fichier audio sélectionné"
        self.var_audio_display = tk.StringVar(value=audio_display)
        self.var_record_button = tk.StringVar(value="Enregistrer…")
        self.var_record_status = tk.StringVar(value="Aucun enregistrement en cours.")

        frm = ttk.Frame(self, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(5, weight=1)

        row=0
        ttk.Label(frm, text="Date (YYYY-MM-DD)").grid(row=row, column=0, sticky="w")
        e_date = ttk.Entry(frm, textvariable=self.var_date, width=16)
        e_date.grid(row=row, column=1, sticky="w")

        row+=1
        ttk.Label(frm, text="Début (HH:MM)").grid(row=row, column=0, sticky="w")
        e_start = ttk.Entry(frm, textvariable=self.var_start, width=10)
        e_start.grid(row=row, column=1, sticky="w")

        row+=1
        ttk.Label(frm, text="Fin (HH:MM)").grid(row=row, column=0, sticky="w")
        e_end = ttk.Entry(frm, textvariable=self.var_end, width=10)
        e_end.grid(row=row, column=1, sticky="w")

        row+=1
        ttk.Label(frm, text="Minutes (auto si début/fin)").grid(row=row, column=0, sticky="w")
        e_min = ttk.Entry(frm, textvariable=self.var_min, width=10)
        e_min.grid(row=row, column=1, sticky="w")

        def refresh_minutes(*_):
            m = minutes_from_times(self.var_start.get(), self.var_end.get())
            if m is not None:
                self.var_min.set(str(int(m)))
        e_start.bind("<FocusOut>", refresh_minutes)
        e_end.bind("<FocusOut>", refresh_minutes)

        row+=1
        ttk.Label(frm, text="Journal (récit libre)").grid(row=row, column=0, sticky="w", pady=(8,0))
        self.txt_journal = tk.Text(frm, height=12, wrap="word")
        self.txt_journal.grid(row=row+1, column=0, columnspan=2, sticky="nsew", pady=(2,0))
        self.txt_journal.insert("1.0", init_journal)

        row += 2
        audio_box = ttk.LabelFrame(frm, text="Enregistrement audio (optionnel)")
        audio_box.grid(row=row, column=0, columnspan=2, sticky="we", pady=(8,0))
        audio_box.columnconfigure(0, weight=1)
        audio_box.columnconfigure(1, weight=0)
        audio_box.columnconfigure(2, weight=0)
        ttk.Label(audio_box, textvariable=self.var_audio_display, wraplength=380).grid(
            row=0, column=0, columnspan=3, sticky="we", padx=4, pady=(4,2)
        )
        ttk.Button(audio_box, text="Choisir un fichier…", command=self.select_audio_file).grid(
            row=1, column=0, sticky="w", padx=4, pady=(0,4)
        )
        ttk.Button(audio_box, text="Effacer", command=self.clear_audio_file).grid(
            row=1, column=1, sticky="w", padx=4, pady=(0,4)
        )
        ttk.Button(audio_box, textvariable=self.var_record_button, command=self.toggle_audio_recording).grid(
            row=1, column=2, sticky="e", padx=4, pady=(0,4)
        )
        ttk.Label(
            audio_box,
            textvariable=self.var_record_status,
            wraplength=380,
            foreground="#555555",
        ).grid(row=2, column=0, columnspan=3, sticky="we", padx=4, pady=(0,4))

        # Boutons
        btns = ttk.Frame(frm)
        btns.grid(row=row+2, column=0, columnspan=2, pady=(8,0), sticky="e")
        ttk.Button(btns, text="Annuler", command=self.destroy).grid(row=0, column=0, padx=4)
        ttk.Button(btns, text="Enregistrer", command=self.on_save).grid(row=0, column=1, padx=4)

        self.transient(master)
        self.wait_visibility()
        self.grab_set()
        self.focus()

    def select_audio_file(self):
        if self._recording_active:
            self.stop_audio_recording(keep_result=False, show_message=False)
        path = filedialog.askopenfilename(
            title="Sélectionner un fichier audio",
            filetypes=[
                ("Fichiers audio", "*.mp3 *.wav *.m4a *.aac *.flac *.ogg"),
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

    def clear_audio_file(self):
        if self._recording_active:
            self.stop_audio_recording(keep_result=False, show_message=False)
        if self._audio_source_is_temp and self.audio_source_path:
            self._discard_temp_recording(self.audio_source_path)
        self.audio_source_path = None
        self.audio_cleared = True
        self.initial_audio_path = ""
        self._audio_source_is_temp = False
        self.var_audio_display.set("Aucun fichier audio sélectionné")

    def toggle_audio_recording(self):
        if self._recording_active:
            self.stop_audio_recording()
        else:
            self.start_audio_recording()

    def start_audio_recording(self):
        if self._recording_active:
            return
        try:
            sd_mod, np_mod = _load_audio_modules()
        except ModuleNotFoundError as exc:
            missing = getattr(exc, "name", "sounddevice") or "sounddevice"
            messagebox.showwarning(
                "Audio",
                "Enregistrement indisponible : installe les dépendances 'sounddevice' et 'numpy' (ex: pip install sounddevice numpy).",
            )
            self.var_record_status.set(f"Dépendance manquante : {missing}")
            return
        except Exception as exc:
            messagebox.showwarning("Audio", f"Impossible d'initialiser l'enregistrement : {exc}")
            self.var_record_status.set("Échec de l'initialisation de l'audio.")
            return

        if self._audio_source_is_temp and self.audio_source_path:
            self._discard_temp_recording(self.audio_source_path)
            self.audio_source_path = None
            self._audio_source_is_temp = False
            self.var_audio_display.set("Aucun fichier audio sélectionné")

        self._recording_modules = (sd_mod, np_mod)
        self._recording_stop_event = threading.Event()
        self._recording_finished_event = threading.Event()
        self._recording_frames = []
        self._recording_error = None
        self._recording_active = True
        self._recording_start_time = time.time()
        self._recording_status_message = None

        samplerate = 44100
        try:
            default_input = getattr(sd_mod.default, "device", (None, None))[0]
            device_info = sd_mod.query_devices(default_input, "input")
        except Exception:
            try:
                device_info = sd_mod.query_devices(None, "input")
            except Exception:
                device_info = {}
        default_samplerate = device_info.get("default_samplerate") if isinstance(device_info, dict) else None
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

        def worker():
            try:
                def callback(indata, frames, time_info, status):
                    if status and str(status).strip():
                        self._recording_status_message = str(status)
                    self._recording_frames.append(indata.copy())
                    if self._recording_stop_event.is_set():
                        raise sd_mod.CallbackStop()

                with sd_mod.InputStream(
                    samplerate=self._recording_samplerate,
                    channels=self._recording_channels,
                    dtype="float32",
                    callback=callback,
                ):
                    self._recording_stop_event.wait()
            except Exception as exc:
                self._recording_error = exc
            finally:
                self._recording_finished_event.set()

        self._recording_thread = threading.Thread(target=worker, daemon=True)
        self._recording_thread.start()
        self._record_timer_job = self.after(200, self._update_recording_timer)
        self._recording_poll_job = self.after(300, self._poll_recording_error)

    def stop_audio_recording(self, *, keep_result=True, show_message=True, error=None):
        active = self._recording_active or (self._recording_thread is not None)
        if not active:
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

        if error is None:
            error = self._recording_error
        self._recording_error = None
        status_message = self._recording_status_message
        self._recording_status_message = None

        if not keep_result:
            self._recording_frames = []
            if error and show_message:
                messagebox.showwarning("Audio", f"Enregistrement interrompu : {error}")
            self.var_record_status.set("Enregistrement annulé.")
            return

        if error:
            self._recording_frames = []
            self.var_record_status.set("Erreur lors de l'enregistrement.")
            if show_message:
                messagebox.showwarning("Audio", f"Erreur lors de l'enregistrement : {error}")
            return

        frames = list(self._recording_frames)
        self._recording_frames = []
        if not frames:
            self.var_record_status.set("Aucune donnée audio capturée.")
            if show_message:
                messagebox.showwarning(
                    "Audio",
                    "Aucune donnée audio n’a été capturée. Vérifie le micro ou réessaie.",
                )
            return

        sd_mod, np_mod = self._recording_modules or (None, None)
        if np_mod is None:
            try:
                _, np_mod = _load_audio_modules()
            except Exception:
                np_mod = None
        if np_mod is None:
            self.var_record_status.set("Erreur : numpy indisponible")
            if show_message:
                messagebox.showwarning(
                    "Audio",
                    "Impossible de finaliser l'enregistrement car numpy est indisponible.",
                )
            return

        try:
            audio_data = np_mod.concatenate(frames, axis=0)
        except Exception:
            audio_data = frames[0]

        try:
            audio_data = np_mod.clip(audio_data, -1.0, 1.0)
            audio_int16 = (audio_data * 32767.0).astype(np_mod.int16)
        except Exception as exc:
            self.var_record_status.set("Erreur lors de la conversion audio.")
            if show_message:
                messagebox.showwarning("Audio", f"Erreur lors de la conversion audio : {exc}")
            return

        if audio_int16.ndim > 1:
            frame_count = audio_int16.shape[0]
            channels = audio_int16.shape[1]
            audio_bytes = audio_int16.reshape(-1).tobytes()
        else:
            frame_count = audio_int16.shape[0]
            channels = self._recording_channels
            audio_bytes = audio_int16.tobytes()

        duration = frame_count / float(self._recording_samplerate or 1)

        try:
            fd, tmp_path = tempfile.mkstemp(prefix="contact_audio_", suffix=".wav")
            os.close(fd)
            with wave.open(tmp_path, "wb") as wf:
                wf.setnchannels(channels)
                wf.setsampwidth(2)
                wf.setframerate(self._recording_samplerate)
                wf.writeframes(audio_bytes)
        except Exception as exc:
            if show_message:
                messagebox.showwarning("Audio", f"Impossible d'enregistrer le fichier audio : {exc}")
            self.var_record_status.set("Erreur lors de la sauvegarde de l'audio.")
            return

        self._temp_recordings.add(tmp_path)
        self.audio_source_path = tmp_path
        self.audio_cleared = False
        self._audio_source_is_temp = True
        self.var_audio_display.set(tmp_path)

        if status_message:
            self.var_record_status.set(
                f"Enregistrement prêt ({duration:.1f} s). Avertissement : {status_message}"
            )
            if show_message:
                messagebox.showwarning(
                    "Audio", f"Avertissement pendant l'enregistrement : {status_message}"
                )
        else:
            self.var_record_status.set(f"Enregistrement prêt ({duration:.1f} s).")

    def on_save(self):
        d = self.var_date.get().strip()
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", d):
            messagebox.showwarning("Date", "Format attendu: YYYY-MM-DD")
            return
        m = safe_float(self.var_min.get(),0)
        if m <= 0:
            m_auto = minutes_from_times(self.var_start.get(), self.var_end.get())
            if m_auto is None or m_auto <= 0:
                messagebox.showwarning("Durée", "Renseigne minutes ou des heures début/fin valides.")
                return
            m = m_auto
        journal = self.txt_journal.get("1.0","end").strip()

        self.result = {
            "contact_id": self.contact_id,
            "date": d,
            "start": self.var_start.get().strip() or None,
            "end": self.var_end.get().strip() or None,
            "minutes": int(round(m)),
            "journal": journal
        }
        if self.audio_cleared:
            self.result["audio_path"] = None
            self.result["audio_removed"] = True
            if self.previous_audio_path:
                self.result["previous_audio_path"] = self.previous_audio_path
        elif self.audio_source_path:
            self.result["audio_source_path"] = self.audio_source_path
            if self._audio_source_is_temp:
                self.result["audio_source_is_temp"] = True
                self._temp_recordings_in_use.add(self.audio_source_path)
        elif self.initial_audio_path:
            self.result["audio_path"] = self.initial_audio_path
        self.destroy()

    def _update_recording_timer(self):
        if not self._recording_active:
            return
        elapsed = time.time() - (self._recording_start_time or time.time())
        self.var_record_status.set(f"Enregistrement en cours… {int(elapsed)}s")
        self._record_timer_job = self.after(200, self._update_recording_timer)

    def _poll_recording_error(self):
        if not self._recording_active:
            return
        if self._recording_error:
            err = self._recording_error
            self.stop_audio_recording(keep_result=False, show_message=True, error=err)
            return
        self._recording_poll_job = self.after(300, self._poll_recording_error)

    def _discard_temp_recording(self, path):
        if not path:
            return
        if path in self._temp_recordings_in_use:
            return
        if path in self._temp_recordings:
            try:
                if os.path.isfile(path):
                    os.remove(path)
            except OSError:
                pass
            self._temp_recordings.discard(path)

    def destroy(self):
        try:
            if self._recording_active or self._recording_thread:
                self.stop_audio_recording(keep_result=False, show_message=False)
        finally:
            pass
        leftovers = [p for p in self._temp_recordings if p not in self._temp_recordings_in_use]
        for path in leftovers:
            try:
                if os.path.isfile(path):
                    os.remove(path)
            except OSError:
                pass
            self._temp_recordings.discard(path)
        super().destroy()

# --------------------------- UI principale ---------------------------

class ConvoiWizard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Convoi Wizard v3 — LOG-FIRST")
        self.geometry("1280x800")
        self.minsize(1120, 720)

        # Données
        self.people = []
        self.global_title = tk.StringVar(value="Convoi relationnel de Kévin")
        self.owner = tk.StringVar(value="Kevin")
        self.current_workspace = None
        self.metadata_path = None
        self.loading = False
        self._autosave_in_progress = False
        self._autosave_pending = False
        self._autosave_pending_force = False
        self._autosave_pending_show_info = False

        self._transcription_jobs = {}
        self._transcription_queue = Queue()
        self._active_transcription = None

        self.default_workspace_var = None
        self.default_workspace = None
        for candidate_var in ("HMT_DEFAULT_WORKSPACE", "CONVOI_DEFAULT_WORKSPACE"):
            raw_value = os.environ.get(candidate_var)
            if raw_value is not None:
                self.default_workspace_var = candidate_var
                cleaned = raw_value.strip()
                if cleaned:
                    self.default_workspace = os.path.expanduser(cleaned)
                break
        self._default_workspace_checked = False

        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # Layout
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)
        self.columnconfigure(2, weight=0)

        self.build_left()
        self.build_center()
        self.build_right()
        self.build_menubar()

        self.select_person(None)
        self.after(100, self.ensure_initial_file)
        self.after(200, self._poll_transcription_queue)

    # -------------------- Workspace path helpers --------------------
    def _normalise_workspace_path(self, workspace, rel_path):
        """Return a normalised relative path and its absolute counterpart.

        Raises ValueError if *rel_path* is empty, absolute or escapes the
        workspace directory.
        """
        if not workspace:
            raise ValueError("Workspace de référence manquant")
        workspace_abs = os.path.abspath(workspace)
        if not isinstance(rel_path, str):
            raise ValueError("Chemin relatif invalide")

        rel_raw = rel_path.strip()
        if not rel_raw:
            raise ValueError("Chemin relatif vide")
        if os.path.isabs(rel_raw):
            raise ValueError("Chemin absolu interdit dans le workspace")
        drive, _ = os.path.splitdrive(rel_raw)
        if drive:
            raise ValueError("Chemin avec lecteur interdit dans le workspace")

        rel_norm = os.path.normpath(rel_raw)
        rel_norm = rel_norm.replace("\\", "/")
        if rel_norm in ("", "."):
            raise ValueError("Chemin relatif invalide")
        if rel_norm.startswith("../") or rel_norm == "..":
            raise ValueError("Chemin hors du workspace")

        abs_path = os.path.normpath(os.path.join(workspace_abs, rel_norm))
        try:
            common = os.path.commonpath([workspace_abs, abs_path])
        except ValueError:
            raise ValueError("Chemin incompatible avec le workspace")
        if common != workspace_abs:
            raise ValueError("Chemin hors du workspace")

        return rel_norm, abs_path

    # -------------------- Menubar --------------------
    def build_menubar(self):
        menubar = tk.Menu(self)

        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Ouvrir un workspace…", command=self.on_import_json)
        filemenu.add_command(label="Exporter le workspace…", command=self.on_export_json)
        filemenu.add_separator()
        filemenu.add_command(label="Quitter", command=self.destroy)
        menubar.add_cascade(label="Fichier", menu=filemenu)

        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="À propos", command=self.on_about)
        menubar.add_cascade(label="Aide", menu=helpmenu)

        self.config(menu=menubar)

    # -------------------- Colonne gauche --------------------
    def build_left(self):
        frame = ttk.Frame(self, padding=8)
        frame.grid(row=0, column=0, sticky="ns")
        frame.rowconfigure(2, weight=1)

        ttk.Label(frame, text="Titre global:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.global_title, width=28).grid(row=0, column=1, sticky="we", pady=2, padx=4)
        ttk.Label(frame, text="Propriétaire:").grid(row=0, column=2, sticky="w")
        ttk.Entry(frame, textvariable=self.owner, width=18).grid(row=0, column=3, sticky="we", pady=2, padx=4)
        self.global_title.trace_add("write", self.on_global_meta_change)
        self.owner.trace_add("write", self.on_global_meta_change)

        ttk.Label(frame, text="Personnes").grid(row=1, column=0, columnspan=4, sticky="w", pady=(8,4))
        self.people_list = tk.Listbox(frame, width=36, height=28, exportselection=False)
        self.people_list.grid(row=2, column=0, columnspan=4, sticky="nswe")
        self.people_list.bind("<<ListboxSelect>>", self.on_select_list)

        btns = ttk.Frame(frame)
        btns.grid(row=3, column=0, columnspan=4, pady=6, sticky="we")
        ttk.Button(btns, text="Ajouter", command=self.on_add_person).grid(row=0, column=0, padx=2)
        ttk.Button(btns, text="Dupliquer", command=self.on_dup_person).grid(row=0, column=1, padx=2)
        ttk.Button(btns, text="Supprimer", command=self.on_del_person).grid(row=0, column=2, padx=2)

    # -------------------- Colonne droite: aide --------------------
    def build_right(self):
        frame = ttk.Frame(self, padding=8)
        frame.grid(row=0, column=2, sticky="ns")
        frame.rowconfigure(1, weight=1)
        ttk.Label(frame, text="Aide / Conseils", font=("TkDefaultFont", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.help_text = tk.Text(frame, width=44, height=36, wrap="word")
        self.help_text.grid(row=1, column=0, sticky="ns")
        self.help_text.insert("1.0",
            "Journal LOG-FIRST:\n"
            "• Ajoute un contact: Date + Heures + Journal (récit libre).\n"
            "• Copie le prompt d’analyse (inclut TOUS les logs), colle-le dans ChatGPT.\n"
            "• Importe le patch JSON renvoyé: tout se met à jour.\n\n"
            "Astuce: Note le canal/format/ressenti dans le texte si tu veux (ils seront inférés).")
        self.help_text.config(state="disabled")

    def set_help(self, key):
        title, body = HELP_TEXT.get(key, (key, ""))
        self.help_text.config(state="normal")
        self.help_text.delete("1.0", "end")
        self.help_text.insert("1.0", f"{title}\n\n{body}")
        self.help_text.config(state="disabled")

    # -------------------- Colonne centrale --------------------
    def build_center(self):
        frame = ttk.Frame(self, padding=8)
        frame.grid(row=0, column=1, sticky="nsew")
        frame.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(frame, text="Prénom / Nom").grid(row=row, column=0, sticky="w")
        self.var_name = tk.StringVar()
        e_name = ttk.Entry(frame, textvariable=self.var_name)
        e_name.grid(row=row, column=1, sticky="we", padx=6, pady=2)
        e_name.bind("<FocusIn>", lambda e: self.set_help("name"))
        self.var_name.trace_add("write", self.on_form_changed)

        ttk.Label(frame, text="Cercle (HMT)").grid(row=row, column=2, sticky="w")
        self.var_circle = tk.StringVar(value="Médiane")
        cb_circ = ttk.Combobox(frame, textvariable=self.var_circle, values=["Noyau","Médiane","Externe"], state="readonly", width=12)
        cb_circ.grid(row=row, column=3, sticky="w", padx=4)
        cb_circ.bind("<<ComboboxSelected>>", lambda e: self.set_help("circle"))
        cb_circ.bind("<FocusIn>", lambda e: self.set_help("circle"))
        self.var_circle.trace_add("write", self.on_form_changed)

        # Sliders (aperçu local, l'analyse peut écraser)
        row += 1
        sliders = ttk.Frame(frame)
        sliders.grid(row=row, column=0, columnspan=4, sticky="we", pady=(6,2))
        for i in range(4): sliders.columnconfigure(i, weight=1)

        ttk.Label(sliders, text="Intimité (0–10)").grid(row=0, column=0, sticky="w")
        ttk.Label(sliders, text="Fiabilité (0–10)").grid(row=0, column=1, sticky="w")
        ttk.Label(sliders, text="Valence (−2..+2)").grid(row=0, column=2, sticky="w")
        ttk.Label(sliders, text="Énergie_coût (0..2)").grid(row=0, column=3, sticky="w")

        self.var_intim = tk.DoubleVar(value=0)
        s1 = ttk.Scale(sliders, from_=0, to=10, variable=self.var_intim, command=lambda v: self.update_preview())
        s1.grid(row=1, column=0, sticky="we", padx=4); s1.bind("<Enter>", lambda e: self.set_help("intimacy"))
        self.var_intim.trace_add("write", self.on_form_changed)

        self.var_reli = tk.DoubleVar(value=0)
        s2 = ttk.Scale(sliders, from_=0, to=10, variable=self.var_reli, command=lambda v: self.update_preview())
        s2.grid(row=1, column=1, sticky="we", padx=4); s2.bind("<Enter>", lambda e: self.set_help("reliability"))
        self.var_reli.trace_add("write", self.on_form_changed)

        self.var_vale = tk.DoubleVar(value=0)
        s3 = ttk.Scale(sliders, from_=-2, to=2, variable=self.var_vale, command=lambda v: self.update_preview())
        s3.grid(row=1, column=2, sticky="we", padx=4); s3.bind("<Enter>", lambda e: self.set_help("valence"))
        self.var_vale.trace_add("write", self.on_form_changed)

        self.var_enrg = tk.DoubleVar(value=0)
        s4 = ttk.Scale(sliders, from_=0, to=2, variable=self.var_enrg, command=lambda v: self.update_preview())
        s4.grid(row=1, column=3, sticky="we", padx=4); s4.bind("<Enter>", lambda e: self.set_help("energy_cost"))
        self.var_enrg.trace_add("write", self.on_form_changed)

        self.lbl_intim_val = ttk.Label(sliders, text="0.0"); self.lbl_intim_val.grid(row=2, column=0, sticky="n", pady=(2,0))
        self.lbl_reli_val = ttk.Label(sliders, text="0.0"); self.lbl_reli_val.grid(row=2, column=1, sticky="n", pady=(2,0))
        self.lbl_vale_val = ttk.Label(sliders, text="0.0"); self.lbl_vale_val.grid(row=2, column=2, sticky="n", pady=(2,0))
        self.lbl_enrg_val = ttk.Label(sliders, text="0.0"); self.lbl_enrg_val.grid(row=2, column=3, sticky="n", pady=(2,0))

        # Match + notes
        row += 1
        ttk.Label(frame, text="Match_besoins (0..2)").grid(row=row, column=0, sticky="w")
        self.var_match = tk.DoubleVar(value=0)
        e_match = ttk.Scale(frame, from_=0, to=2, variable=self.var_match, command=lambda v: self.update_preview())
        e_match.grid(row=row, column=1, sticky="we", padx=6)
        self.var_match.trace_add("write", self.on_form_changed)
        ttk.Label(frame, text="Notes").grid(row=row, column=2, sticky="nw")
        self.txt_notes = tk.Text(frame, height=3)
        self.txt_notes.grid(row=row, column=3, sticky="we", pady=4)
        self.txt_notes.bind("<FocusIn>", lambda e: self.set_help("notes"))
        self.txt_notes.bind("<<Modified>>", self.on_notes_modified)

        # Needs (optionnel)
        row += 1
        needs_box = ttk.LabelFrame(frame, text="Besoins nourris (tags, optionnel)")
        needs_box.grid(row=row, column=0, columnspan=4, sticky="we", pady=(8,2))
        needs_box.columnconfigure(0, weight=1); needs_box.columnconfigure(1, weight=1); needs_box.columnconfigure(2, weight=1)
        self.needs_box = needs_box
        self.need_vars = {}
        self._needs_grid_r = 0
        self._needs_grid_c = 0
        for need in DEFAULT_NEEDS:
            self._add_need_checkbox(need, checked=False)

        # Contacts (journal)
        row += 1
        ct = ttk.LabelFrame(frame, text="Contacts (journal libre)")
        ct.grid(row=row, column=0, columnspan=4, sticky="we", pady=(8,2))
        for i in range(6): ct.columnconfigure(i, weight=1)

        ttk.Button(ct, text="Ajouter contact…", command=self.add_contact_dialog).grid(row=0, column=0, sticky="w", padx=4, pady=2)
        ttk.Button(ct, text="Éditer…", command=self.edit_contact_dialog).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Button(ct, text="Supprimer", command=self.del_contact_row).grid(row=0, column=2, sticky="w", padx=4)
        ttk.Button(ct, text="Copier prompt d’analyse (TOUS logs)", command=self.on_copy_analysis_prompt).grid(row=0, column=3, sticky="w", padx=4)
        ttk.Button(ct, text="Importer analyse (JSON)…", command=self.on_import_analysis_json).grid(row=0, column=4, sticky="w", padx=4)

        cols = ("date","start","end","minutes","journal_preview","contact_id")
        self.contacts_tree = ttk.Treeview(ct, columns=cols, show="headings", height=10)
        headers = {
            "date":"date", "start":"début", "end":"fin", "minutes":"min",
            "journal_preview":"journal (aperçu)", "contact_id":"id"
        }
        for c in cols:
            self.contacts_tree.heading(c, text=headers[c])
            self.contacts_tree.column(c, width=120 if c not in ("journal_preview","contact_id") else (420 if c=="journal_preview" else 160), anchor="w")
        self.contacts_tree.grid(row=1, column=0, columnspan=6, sticky="we", pady=4)
        self.contacts_tree.bind("<Double-1>", lambda e: self.edit_contact_dialog())

        row += 1
        trans_frame = ttk.LabelFrame(frame, text="Transcription audio (Vibe)")
        trans_frame.grid(row=row, column=0, columnspan=4, sticky="we", pady=(4, 0))
        for col in range(3):
            trans_frame.columnconfigure(col, weight=1 if col < 2 else 0)
        self.transcription_contact_var = tk.StringVar(value="Aucune transcription en cours.")
        self.transcription_status_var = tk.StringVar(value="")
        self.transcription_progress_var = tk.DoubleVar(value=0.0)
        self.transcription_progressbar = ttk.Progressbar(
            trans_frame,
            maximum=100,
            variable=self.transcription_progress_var,
            mode="determinate",
        )
        self.transcription_progressbar.grid(row=0, column=0, columnspan=2, sticky="we", padx=4, pady=(4, 2))
        self.transcription_cancel_btn = ttk.Button(
            trans_frame,
            text="Annuler",
            command=self.on_cancel_transcription,
            state="disabled",
        )
        self.transcription_cancel_btn.grid(row=0, column=2, sticky="e", padx=4, pady=(4, 2))
        self.transcription_label = ttk.Label(trans_frame, textvariable=self.transcription_contact_var)
        self.transcription_label.grid(row=1, column=0, sticky="w", padx=4, pady=(0, 4))
        self.transcription_state_label = ttk.Label(
            trans_frame,
            textvariable=self.transcription_status_var,
            foreground="#555555",
        )
        self.transcription_state_label.grid(row=1, column=1, columnspan=2, sticky="e", padx=4, pady=(0, 4))

        # Aperçu score/action
        row += 1
        prev = ttk.Frame(frame); prev.grid(row=row, column=0, columnspan=4, sticky="we", pady=6)
        ttk.Label(prev, text="Score priorité (aperçu):").grid(row=0, column=0, sticky="w")
        self.lbl_score = ttk.Label(prev, text="0.00"); self.lbl_score.grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(prev, text="Action:").grid(row=0, column=2, sticky="e")
        self.action_box = tk.Label(prev, text="—", bg="#cccccc", fg="white", width=18)
        self.action_box.grid(row=0, column=3, sticky="w", padx=6)

        # Commandes
        row += 1
        cmds = ttk.Frame(frame); cmds.grid(row=row, column=0, columnspan=4, sticky="we", pady=(6,0))
        ttk.Button(cmds, text="Sauver la personne", command=self.save_current_person).grid(row=0, column=0, padx=4)
        ttk.Button(cmds, text="Exporter workspace…", command=self.on_export_json).grid(row=0, column=1, padx=4)
        ttk.Button(cmds, text="Ouvrir workspace…", command=self.on_import_json).grid(row=0, column=2, padx=4)

    # ----- needs dynamiques -----
    def _add_need_checkbox(self, need, checked=True):
        v = self.need_vars.get(need)
        if v is None:
            v = tk.BooleanVar(value=checked)
            self.need_vars[need] = v
            chk = ttk.Checkbutton(
                self.needs_box,
                text=need,
                variable=v,
                command=lambda n=need: self.on_need_toggle(n),
            )
            chk.grid(row=self._needs_grid_r, column=self._needs_grid_c, sticky="w", padx=4, pady=2)
            self._needs_grid_c += 1
            if self._needs_grid_c >= 3:
                self._needs_grid_c = 0
                self._needs_grid_r += 1
        else:
            v.set(checked)

    # -------------------- Liste personnes --------------------
    def on_add_person(self):
        if getattr(self, "current_index", None) is not None:
            self.update_current_person_from_form()
        p = {
            "person_id": str(uuid.uuid4()),
            "name": "Prénom",
            "circle": "Médiane",
            "intimacy": 0, "reliability": 0, "valence": 0,
            "energy_cost": 0, "match_besoins": 0,
            "needs_nourished": [],
            "need_support": {},
            "contacts": [],
            "ambivalent": False,
            "status": "",
            "notes": ""
        }
        self.people.append(p)
        self.refresh_people_list()
        self.people_list.select_set(len(self.people)-1)
        self.select_person(len(self.people)-1)
        self.autosave()

    def on_dup_person(self):
        i = self.get_sel_index()
        if i is None: return
        dup = copy.deepcopy(self.people[i])
        dup["name"] = (dup.get("name","") or "SansNom") + " (copie)"
        dup["person_id"] = str(uuid.uuid4())
        dup.pop("person_rel_path", None)
        contacts = []
        for contact in dup.get("contacts", []):
            new_contact = copy.deepcopy(contact)
            new_contact["contact_id"] = str(uuid.uuid4())
            new_contact.pop("transcription_path", None)
            new_contact.pop("transcript_path", None)
            new_contact.pop("audio_path", None)
            contacts.append(new_contact)
        dup["contacts"] = contacts
        self.people.append(dup)
        self.refresh_people_list()
        self.people_list.select_set(len(self.people)-1)
        self.select_person(len(self.people)-1)
        self.autosave()

    def on_del_person(self):
        i = self.get_sel_index()
        if i is None: return
        if messagebox.askyesno("Confirmer", "Supprimer cette personne ?"):
            self.people.pop(i)
            self.refresh_people_list()
            self.select_person(None)
            self.autosave()

    def on_select_list(self, event=None):
        new_idx = self.get_sel_index()
        if new_idx == getattr(self, "current_index", None):
            return
        if getattr(self, "current_index", None) is not None:
            if self.update_current_person_from_form():
                self.autosave()
        self.select_person(new_idx)

    def get_sel_index(self):
        sel = self.people_list.curselection()
        if not sel: return None
        return int(sel[0])

    def refresh_people_list(self):
        sel = self.current_index if hasattr(self, "current_index") else None
        self.people_list.delete(0, "end")
        for p in self.people:
            self.people_list.insert("end", f"{p.get('name','?')} [{circle_norm(p.get('circle',''))}]")
        if sel is not None and 0 <= sel < len(self.people):
            try:
                self.people_list.select_set(sel)
                self.people_list.see(sel)
            except tk.TclError:
                pass

    def ensure_initial_file(self):
        if self.current_workspace:
            return
        if not self._default_workspace_checked:
            self._default_workspace_checked = True
            if self.default_workspace:
                candidate_abs = os.path.abspath(self.default_workspace)
                if os.path.isdir(candidate_abs):
                    if self.load_from_file(candidate_abs, show_info=True):
                        return
                else:
                    messagebox.showwarning(
                        "Workspace par défaut introuvable",
                        (
                            "Le dossier indiqué via la variable d’environnement "
                            f"{self.default_workspace_var} est introuvable :\n{candidate_abs}.\n\n"
                            "Sélectionne un dossier manuellement."
                        ),
                    )
            # Si la variable est définie mais vide ou que le chargement échoue, on retombe sur le flux classique.
        resp = messagebox.askyesnocancel(
            "Workspace requis",
            "Pour éviter toute perte de données, ouvre un workspace existant (Oui) ou crée un nouveau dossier (Non).\n"
            "Annuler fermera l’application.",
        )
        if resp is None:
            messagebox.showwarning(
                "Workspace requis",
                "Aucun dossier sélectionné : l’application va se fermer pour garantir tes données.",
            )
            self.destroy()
            return
        if resp:
            path = filedialog.askdirectory(
                title="Ouvrir un workspace existant",
            )
            if path and self.load_from_file(path):
                return
        else:
            path = filedialog.askdirectory(
                title="Créer un nouveau workspace (dossier vide recommandé)",
            )
            if path:
                if os.path.isdir(path) and os.listdir(path):
                    if not messagebox.askyesno(
                        "Dossier non vide",
                        "Le dossier sélectionné n’est pas vide. Les fichiers existants pourraient être écrasés. Continuer ?",
                    ):
                        path = None
                if path:
                    try:
                        os.makedirs(path, exist_ok=True)
                    except Exception as e:
                        messagebox.showerror("Erreur", f"Impossible de préparer le dossier : {e}")
                    else:
                        try:
                            self.loading = True
                            self.people = []
                            self.current_index = None
                            self.refresh_people_list()
                            self.select_person(None)
                        finally:
                            self.loading = False
                        if self.save_to_file(path, show_message=False):
                            messagebox.showinfo("Workspace prêt", f"Nouveau workspace créé :\n{path}")
                            return
        self.after(200, self.ensure_initial_file)

    def load_from_file(self, path, show_info=True):
        workspace = os.path.abspath(path)
        metadata_path = os.path.join(workspace, "metadata.json")
        if not os.path.isdir(workspace):
            messagebox.showerror("Erreur d’import", "Le dossier sélectionné n’existe pas ou n’est pas un dossier.")
            return False
        if not os.path.isfile(metadata_path):
            messagebox.showerror("Erreur d’import", "metadata.json est introuvable dans le dossier sélectionné.")
            return False

        try:
            with open(metadata_path, "r", encoding="utf-8") as f:
                metadata = json.load(f)
        except Exception as e:
            messagebox.showerror("Erreur d’import", f"Lecture de metadata.json impossible : {e}")
            return False

        people_entries = metadata.get("people", [])
        if people_entries and not isinstance(people_entries, list):
            messagebox.showerror("Erreur d’import", "Le champ 'people' de metadata.json doit être une liste.")
            return False

        loaded_people = []
        warnings = []
        for entry in people_entries:
            if not isinstance(entry, dict):
                warnings.append("Entrée de personne invalide dans metadata.json.")
                continue
            person_rel_path = entry.get("person_path")
            if not person_rel_path:
                warnings.append("Entrée de personne sans 'person_path'.")
                continue
            try:
                person_rel_path, person_abs_path = self._normalise_workspace_path(workspace, person_rel_path)
            except ValueError:
                warnings.append(f"Chemin de personne invalide : {person_rel_path!r}.")
                continue
            if not os.path.isfile(person_abs_path):
                warnings.append(f"Fichier personne manquant : {person_rel_path}")
                continue
            try:
                with open(person_abs_path, "r", encoding="utf-8") as pf:
                    person_data = json.load(pf)
            except Exception as exc:
                warnings.append(f"Lecture impossible de {person_rel_path} : {exc}")
                continue
            if not isinstance(person_data, dict):
                warnings.append(f"Format inattendu dans {person_rel_path}.")
                continue
            person_id = entry.get("person_id") or person_data.get("person_id") or str(uuid.uuid4())
            person_data["person_id"] = person_id
            person_data["person_rel_path"] = person_rel_path
            contacts = []
            for contact in person_data.get("contacts", []) or []:
                if not isinstance(contact, dict):
                    continue
                cid = contact.get("contact_id") or str(uuid.uuid4())
                contact["contact_id"] = cid

                transcript_candidate = contact.get("transcript_path") or contact.get("transcription_path")
                transcript_rel = None
                transcript_abs = None
                transcript_text = ""
                if transcript_candidate:
                    try:
                        transcript_rel, transcript_abs = self._normalise_workspace_path(workspace, transcript_candidate)
                    except ValueError:
                        warnings.append(
                            f"Transcription ignorée pour {cid} : chemin invalide ({transcript_candidate!r})."
                        )
                        transcript_rel = None
                        transcript_abs = None
                    else:
                        try:
                            with open(transcript_abs, "r", encoding="utf-8") as tf:
                                transcript_text = tf.read()
                        except FileNotFoundError:
                            warnings.append(f"Transcription manquante pour {cid} ({transcript_rel}).")
                        except Exception as exc:
                            warnings.append(f"Lecture impossible de {transcript_rel} : {exc}")

                journal_candidate = contact.get("journal_path")
                journal_rel = None
                journal_abs = None
                journal_text = ""
                if journal_candidate:
                    try:
                        journal_rel, journal_abs = self._normalise_workspace_path(workspace, journal_candidate)
                    except ValueError:
                        warnings.append(
                            f"Journal ignoré pour {cid} : chemin invalide ({journal_candidate!r})."
                        )
                        journal_rel = None
                        journal_abs = None
                    else:
                        try:
                            with open(journal_abs, "r", encoding="utf-8") as jf:
                                journal_text = jf.read()
                        except FileNotFoundError:
                            warnings.append(f"Journal manquant pour {cid} ({journal_rel}).")
                        except Exception as exc:
                            warnings.append(f"Lecture impossible de {journal_rel} : {exc}")

                default_journal_rel = os.path.join("people", person_id, "contacts", cid, "journal.txt")
                if not journal_rel:
                    try:
                        journal_rel, journal_abs = self._normalise_workspace_path(workspace, default_journal_rel)
                    except ValueError:
                        journal_rel = default_journal_rel.replace("\\", "/")
                        journal_abs = os.path.normpath(os.path.join(workspace, journal_rel))
                if not journal_text:
                    if journal_abs and os.path.isfile(journal_abs):
                        try:
                            with open(journal_abs, "r", encoding="utf-8") as jf:
                                journal_text = jf.read()
                        except Exception as exc:
                            warnings.append(f"Lecture impossible de {journal_rel} : {exc}")
                    elif transcript_text:
                        journal_text = transcript_text

                if transcript_rel:
                    contact["transcript_path"] = transcript_rel
                else:
                    contact.pop("transcript_path", None)
                if journal_rel:
                    contact["journal_path"] = journal_rel
                else:
                    contact.pop("journal_path", None)

                contact.pop("transcription_path", None)
                contact["journal"] = journal_text

                audio_rel = contact.get("audio_path")
                if audio_rel:
                    try:
                        audio_rel, _ = self._normalise_workspace_path(workspace, audio_rel)
                    except ValueError:
                        warnings.append(
                            f"Chemin audio ignoré pour {cid} : {audio_rel!r} n’est pas autorisé."
                        )
                        contact.pop("audio_path", None)
                    else:
                        contact["audio_path"] = audio_rel
                contacts.append(contact)
            person_data["contacts"] = contacts
            loaded_people.append(person_data)

        try:
            self.loading = True
            self.global_title.set(metadata.get("title", "Convoi relationnel"))
            self.owner.set(metadata.get("owner", ""))
            self.people = loaded_people
            self.current_workspace = workspace
            self.metadata_path = metadata_path
            self.current_index = None
            self.refresh_people_list()
            self.select_person(None if not self.people else 0)
        finally:
            self.loading = False

        if warnings:
            messagebox.showwarning(
                "Import partiel",
                "\n".join(warnings),
            )

        if show_info:
            messagebox.showinfo("Importé", f"{len(self.people)} personnes chargées depuis :\n{workspace}")
        return True

    def build_export_payload(self):
        workspace_for_norm = self.current_workspace or os.getcwd()
        metadata = {
            "workspace_version": WORKSPACE_VERSION,
            "title": self.global_title.get().strip() or "Convoi relationnel",
            "owner": self.owner.get().strip() or "",
            "last_updated": datetime.datetime.utcnow().replace(microsecond=0).isoformat() + "Z",
            "people": [],
        }

        for person in self.people:
            if not isinstance(person, dict):
                continue
            person_id = person.get("person_id") or str(uuid.uuid4())
            person["person_id"] = person_id
            person_rel_candidate = person.get("person_rel_path") or os.path.join("people", person_id, "person.json")
            try:
                person_rel_path, _ = self._normalise_workspace_path(workspace_for_norm, person_rel_candidate)
            except ValueError:
                default_person_path = os.path.join("people", person_id, "person.json")
                person_rel_path, _ = self._normalise_workspace_path(workspace_for_norm, default_person_path)
            person["person_rel_path"] = person_rel_path

            meta_person = {
                "person_id": person_id,
                "name": person.get("name", ""),
                "circle": person.get("circle", ""),
                "person_path": person_rel_path,
                "contacts": [],
            }

            for contact in person.get("contacts", []) or []:
                if not isinstance(contact, dict):
                    continue
                contact_id = contact.get("contact_id") or str(uuid.uuid4())
                contact["contact_id"] = contact_id
                default_journal = os.path.join(
                    "people", person_id, "contacts", contact_id, "journal.txt"
                )
                journal_candidate = contact.get("journal_path") or default_journal
                try:
                    journal_path, _ = self._normalise_workspace_path(
                        workspace_for_norm, journal_candidate
                    )
                except ValueError:
                    journal_path, _ = self._normalise_workspace_path(
                        workspace_for_norm, default_journal
                    )
                contact["journal_path"] = journal_path

                transcription_candidate = contact.get("transcript_path") or contact.get("transcription_path")
                transcript_path = None
                if transcription_candidate:
                    try:
                        transcript_path, _ = self._normalise_workspace_path(
                            workspace_for_norm, transcription_candidate
                        )
                    except ValueError:
                        transcript_path = None
                if transcript_path:
                    contact["transcript_path"] = transcript_path
                else:
                    contact.pop("transcript_path", None)
                contact.pop("transcription_path", None)

                audio_value = contact.get("audio_path")
                audio_rel = None
                if audio_value:
                    audio_candidate = str(audio_value).strip()
                    if self.current_workspace and os.path.isabs(audio_candidate):
                        try:
                            audio_candidate = os.path.relpath(audio_candidate, self.current_workspace)
                        except ValueError:
                            audio_candidate = ""
                    if audio_candidate:
                        try:
                            audio_rel, _ = self._normalise_workspace_path(workspace_for_norm, audio_candidate)
                        except ValueError:
                            audio_rel = None
                if audio_rel:
                    contact["audio_path"] = audio_rel
                else:
                    contact.pop("audio_path", None)

                contact_meta = {
                    "contact_id": contact_id,
                    "journal_path": journal_path,
                }
                if transcript_path:
                    contact_meta["transcript_path"] = transcript_path
                if audio_rel:
                    contact_meta["audio_path"] = audio_rel
                meta_person["contacts"].append(contact_meta)

            metadata["people"].append(meta_person)

        return metadata

    def _write_workspace_files(self, metadata):
        if not self.current_workspace:
            raise RuntimeError("Aucun workspace courant défini")

        workspace = self.current_workspace
        os.makedirs(workspace, exist_ok=True)
        people_dir = os.path.join(workspace, "people")
        os.makedirs(people_dir, exist_ok=True)

        referenced_people_dirs = set()
        people_by_id = {p.get("person_id"): p for p in self.people if isinstance(p, dict)}

        for entry in metadata.get("people", []):
            if not isinstance(entry, dict):
                continue
            person_id = entry.get("person_id")
            person_rel_path = entry.get("person_path")
            if not person_id or not person_rel_path:
                continue
            try:
                person_rel_path, person_abs_path = self._normalise_workspace_path(workspace, person_rel_path)
            except ValueError:
                raise ValueError(f"Chemin de personne invalide : {person_rel_path!r}")
            entry["person_path"] = person_rel_path
            person_dir = os.path.dirname(person_abs_path)
            os.makedirs(person_dir, exist_ok=True)
            dir_name = os.path.basename(person_dir)
            referenced_people_dirs.add(dir_name)

            person = people_by_id.get(person_id)
            if not person:
                continue

            person_copy = copy.deepcopy(person)
            person_copy.pop("person_rel_path", None)

            contacts_dir = os.path.join(person_dir, "contacts")
            os.makedirs(contacts_dir, exist_ok=True)
            existing_contacts = {
                name
                for name in os.listdir(contacts_dir)
                if os.path.isdir(os.path.join(contacts_dir, name))
            }
            expected_contacts = set()

            contacts_data = []
            for contact in person_copy.get("contacts", []) or []:
                if not isinstance(contact, dict):
                    continue
                contact_id = contact.get("contact_id") or str(uuid.uuid4())
                contact["contact_id"] = contact_id
                expected_contacts.add(contact_id)

                journal_candidate = contact.get("journal_path") or os.path.join(
                    "people", person_id, "contacts", contact_id, "journal.txt"
                )
                try:
                    journal_path, journal_abs_path = self._normalise_workspace_path(
                        workspace, journal_candidate
                    )
                except ValueError:
                    raise ValueError(
                        f"Chemin de journal invalide pour le contact {contact_id} : {journal_candidate!r}"
                    )
                contact["journal_path"] = journal_path

                transcription_candidate = contact.get("transcript_path") or contact.get("transcription_path")
                if transcription_candidate:
                    try:
                        transcript_path, _ = self._normalise_workspace_path(
                            workspace, transcription_candidate
                        )
                    except ValueError:
                        raise ValueError(
                            f"Chemin de transcription invalide pour le contact {contact_id} : {transcription_candidate!r}"
                        )
                    contact["transcript_path"] = transcript_path
                else:
                    contact.pop("transcript_path", None)
                contact.pop("transcription_path", None)

                contact_dir = os.path.join(contacts_dir, contact_id)
                os.makedirs(contact_dir, exist_ok=True)

                os.makedirs(os.path.dirname(journal_abs_path), exist_ok=True)
                with open(journal_abs_path, "w", encoding="utf-8") as jf:
                    jf.write(contact.get("journal", "") or "")

                contact_copy = {k: v for k, v in contact.items() if k != "journal"}
                audio_rel = contact_copy.get("audio_path")
                if audio_rel:
                    try:
                        audio_rel, _ = self._normalise_workspace_path(workspace, audio_rel)
                    except ValueError:
                        raise ValueError(
                            f"Chemin audio invalide pour le contact {contact_id} : {audio_rel!r}"
                        )
                    contact_copy["audio_path"] = audio_rel
                contacts_data.append(contact_copy)

            for name in existing_contacts - expected_contacts:
                shutil.rmtree(os.path.join(contacts_dir, name), ignore_errors=True)

            person_copy["contacts"] = contacts_data

            with open(person_abs_path, "w", encoding="utf-8") as pf:
                json.dump(person_copy, pf, ensure_ascii=False, indent=2)

        existing_people_dirs = {
            name
            for name in os.listdir(people_dir)
            if os.path.isdir(os.path.join(people_dir, name))
        }
        for name in existing_people_dirs - referenced_people_dirs:
            shutil.rmtree(os.path.join(people_dir, name), ignore_errors=True)

    def save_to_file(self, path=None, show_message=False):
        if path:
            self.current_workspace = os.path.abspath(path)
        if not self.current_workspace:
            return False

        self.metadata_path = os.path.join(self.current_workspace, "metadata.json")
        metadata = self.build_export_payload()

        try:
            self._write_workspace_files(metadata)
            with open(self.metadata_path, "w", encoding="utf-8") as f:
                json.dump(metadata, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Erreur", str(e))
            return False

        if show_message:
            messagebox.showinfo("Sauvegardé", f"Workspace sauvegardé :\n{self.current_workspace}")
        return True

    def autosave(self, force=False, show_info=False):
        if self._autosave_in_progress:
            self._autosave_pending = True
            self._autosave_pending_force = self._autosave_pending_force or force
            self._autosave_pending_show_info = self._autosave_pending_show_info or show_info
            return False

        force = force or self._autosave_pending_force
        show_info = show_info or self._autosave_pending_show_info
        self._autosave_pending = False
        self._autosave_pending_force = False
        self._autosave_pending_show_info = False

        if self.loading and not force:
            return False
        if not self.current_workspace:
            if force:
                self.ensure_initial_file()
            return False
        try:
            self._autosave_in_progress = True
            return self.save_to_file(show_message=show_info)
        finally:
            self._autosave_in_progress = False
            if self._autosave_pending:
                pending_force = self._autosave_pending_force
                pending_show = self._autosave_pending_show_info
                self._autosave_pending = False
                self._autosave_pending_force = False
                self._autosave_pending_show_info = False
                self.after_idle(lambda pf=pending_force, ps=pending_show: self.autosave(force=pf, show_info=ps))

    def on_global_meta_change(self, *_):
        if self.loading:
            return
        self.autosave()

    def on_form_changed(self, *_):
        if self.loading:
            return
        if self.update_current_person_from_form():
            self.autosave()

    def on_notes_modified(self, event=None):
        if self.loading:
            self.txt_notes.edit_modified(False)
            return
        if self.txt_notes.edit_modified():
            self.txt_notes.edit_modified(False)
            if self.update_current_person_from_form():
                self.autosave()

    def on_need_toggle(self, need=None):
        if self.loading:
            return
        if self.update_current_person_from_form():
            self.autosave()
        self.update_preview()

    def update_current_person_from_form(self):
        idx = getattr(self, "current_index", None)
        if idx is None or idx < 0 or idx >= len(self.people):
            return False
        original = copy.deepcopy(self.people[idx])
        p = self.people[idx]
        p["name"] = self.var_name.get().strip() or "SansNom"
        p["circle"] = circle_norm(self.var_circle.get())
        p["intimacy"] = round(safe_float(self.var_intim.get()), 2)
        p["reliability"] = round(safe_float(self.var_reli.get()), 2)
        p["valence"] = round(safe_float(self.var_vale.get()), 2)
        p["energy_cost"] = round(safe_float(self.var_enrg.get()), 2)
        p["match_besoins"] = round(safe_float(self.var_match.get()), 2)
        p["needs_nourished"] = [n for n, v in self.need_vars.items() if v.get()]
        base_contacts = original.get("contacts", []) if isinstance(original, dict) else []
        p["contacts"] = self.collect_contacts_list(base_contacts=base_contacts)
        p["notes"] = self.txt_notes.get("1.0", "end").strip()
        self.people[idx] = p
        changed = p != original
        if changed:
            self.refresh_people_list()
        return changed

    # -------------------- Form <-> dict --------------------
    def select_person(self, idx):
        try:
            self.loading = True
            self.current_index = idx
            # reset
            self.var_name.set(""); self.var_circle.set("Médiane")
            self.var_intim.set(0); self.var_reli.set(0); self.var_vale.set(0); self.var_enrg.set(0)
            self.var_match.set(0)
            for v in self.need_vars.values(): v.set(False)
            self.contacts_tree.delete(*self.contacts_tree.get_children())
            self.txt_notes.delete("1.0","end")

            if idx is None:
                return

            p = self.people[idx]
            self.var_name.set(p.get("name",""))
            self.var_circle.set(circle_norm(p.get("circle","Médiane")))
            self.var_intim.set(safe_float(p.get("intimacy",0)))
            self.var_reli.set(safe_float(p.get("reliability",0)))
            self.var_vale.set(safe_float(p.get("valence",0)))
            self.var_enrg.set(safe_float(p.get("energy_cost",0)))
            self.var_match.set(safe_float(p.get("match_besoins",0)))

            for need in p.get("needs_nourished", []):
                self._add_need_checkbox(need, checked=True)

            # contacts
            for c in sorted(p.get("contacts", []), key=lambda x: x.get("date",""), reverse=True):
                preview = self._contact_preview_label(c)
                self.contacts_tree.insert("", "end", values=(
                    c.get("date",""), c.get("start",""), c.get("end",""),
                    c.get("minutes",""), preview, c.get("contact_id","")
                ))

            self.txt_notes.insert("1.0", p.get("notes",""))
        finally:
            self.loading = False
            self.txt_notes.edit_modified(False)
            self.update_preview()

    def save_current_person(self):
        idx = self.current_index
        if idx is None:
            messagebox.showinfo("Info","Sélectionne ou ajoute une personne d'abord.")
            return
        self.update_current_person_from_form()
        if self.autosave(force=True, show_info=True):
            return
        messagebox.showwarning(
            "Attention",
            "La sauvegarde automatique a échoué car aucun fichier JSON n’est défini.",
        )

    def _contact_preview_label(self, contact):
        if not isinstance(contact, dict):
            return ""
        preview = (contact.get("journal", "") or "").strip().replace("\n", " ")
        if len(preview) > 120:
            preview = preview[:117] + "..."
        if not preview:
            preview = "—"
        indicators = []
        if contact.get("audio_path"):
            indicators.append("🎧")
        if contact.get("transcript_path") or contact.get("transcription_path"):
            indicators.append("📝")
        if indicators:
            preview = f"{preview} [{' '.join(indicators)}]"
        return preview

    def collect_contacts_list(self, base_contacts=None):
        out = []
        # On n'a que l'aperçu dans la table; on garde les journaux depuis self.people
        current = {}
        contacts_source = base_contacts
        if contacts_source is None:
            if self.current_index is not None and 0 <= self.current_index < len(self.people):
                contacts_source = self.people[self.current_index].get("contacts", [])
            else:
                contacts_source = []
        for c in contacts_source:
            current[c.get("contact_id")] = c

        for iid in self.contacts_tree.get_children():
            d, st, en, m, preview, cid = self.contacts_tree.item(iid,"values")
            entry = {
                "contact_id": cid or str(uuid.uuid4()),
                "date": str(d).strip(),
                "start": (str(st).strip() or None),
                "end": (str(en).strip() or None),
                "minutes": int(round(safe_float(m,0))),
                "journal": (current.get(cid, {}) or {}).get("journal","")
            }
            # préserver champs enrichis par analyse si déjà présents
            for k in ("channel","needs","valence","mood_after","format","note","tags",
                      "audio_path","transcript_path","transcription_path","journal_path"):
                if k in (current.get(cid, {}) or {}):
                    entry[k] = current[cid][k]
            if entry.get("transcription_path") and not entry.get("transcript_path"):
                entry["transcript_path"] = entry["transcription_path"]
            entry.pop("transcription_path", None)
            out.append(entry)
        return out

    # -------------------- Contacts actions --------------------
    def add_contact_dialog(self):
        dlg = ContactDialog(self, initial=None)
        self.wait_window(dlg)
        if not dlg.result:
            return
        r = dlg.result
        audio_source = r.pop("audio_source_path", None)
        audio_removed = r.pop("audio_removed", False)
        audio_is_temp = r.pop("audio_source_is_temp", False)
        contact_data = self._prepare_contact_assets(
            r,
            audio_source_path=audio_source,
            audio_removed=audio_removed,
            audio_source_is_temp=audio_is_temp,
        )
        preview = self._contact_preview_label(contact_data)
        self.contacts_tree.insert("", "end", values=(
            contact_data.get("date",""), contact_data.get("start",""), contact_data.get("end",""),
            contact_data.get("minutes",""), preview, contact_data.get("contact_id","")
        ))
        # Injecter le journal dans la structure interne (pour ne rien perdre)
        self._upsert_contact_full(contact_data)
        self.update_current_person_from_form()
        self.autosave()

    def edit_contact_dialog(self):
        sel = self.contacts_tree.selection()
        if not sel:
            messagebox.showinfo("Info","Sélectionne une ligne de contact.")
            return
        d, st, en, m, preview, cid = self.contacts_tree.item(sel[0], "values")
        # récupérer le journal complet actuel
        full = self._get_contact_by_id(cid) or {}
        dlg = ContactDialog(self, initial={
            "contact_id": cid,
            "date": d, "start": st, "end": en,
            "minutes": m, "journal": full.get("journal",""),
            "audio_path": full.get("audio_path")
        })
        self.wait_window(dlg)
        if not dlg.result:
            return
        r = dlg.result
        audio_source = r.pop("audio_source_path", None)
        audio_removed = r.pop("audio_removed", False)
        audio_is_temp = r.pop("audio_source_is_temp", False)
        contact_data = self._prepare_contact_assets(
            r,
            audio_source_path=audio_source,
            audio_removed=audio_removed,
            audio_source_is_temp=audio_is_temp,
        )
        preview = self._contact_preview_label(contact_data)
        # maj ligne
        self.contacts_tree.item(sel[0], values=(
            contact_data.get("date",""), contact_data.get("start",""), contact_data.get("end",""),
            contact_data.get("minutes",""), preview, contact_data.get("contact_id","")
        ))
        # maj interne
        self._upsert_contact_full(contact_data)
        self.update_current_person_from_form()
        self.autosave()

    def del_contact_row(self):
        for sel in self.contacts_tree.selection():
            _, _, _, _, _, cid = self.contacts_tree.item(sel,"values")
            # supprimer aussi dans la structure interne
            if self.current_index is not None:
                p = self.people[self.current_index]
                p["contacts"] = [c for c in p.get("contacts",[]) if c.get("contact_id") != cid]
                self.people[self.current_index] = p
            self.contacts_tree.delete(sel)
        # actualiser l’état général après suppression
        self.update_current_person_from_form()
        self.autosave()

    def _get_contact_by_id(self, cid):
        if not cid or self.current_index is None: return None
        for c in (self.people[self.current_index].get("contacts", []) if 0 <= self.current_index < len(self.people) else []):
            if c.get("contact_id") == cid:
                return c
        return None

    def _upsert_contact_full(self, contact_dict):
        if self.current_index is None: return
        p = self.people[self.current_index]
        cs = p.get("contacts", [])
        found = False
        for i,c in enumerate(cs):
            if c.get("contact_id") == contact_dict.get("contact_id"):
                cs[i] = {**c, **contact_dict}  # merge (préserve enrichissements)
                found = True
                break
        if not found:
            cs.append(contact_dict)
        p["contacts"] = cs

    def _prepare_contact_assets(
        self,
        contact_dict,
        audio_source_path=None,
        audio_removed=False,
        audio_source_is_temp=False,
    ):
        contact = copy.deepcopy(contact_dict or {})
        if not isinstance(contact, dict):
            return contact_dict

        idx = getattr(self, "current_index", None)
        if idx is None or idx < 0 or idx >= len(self.people):
            if audio_removed:
                contact["audio_path"] = None
            contact.pop("previous_audio_path", None)
            return contact

        workspace = self.current_workspace
        person = self.people[idx]
        person_id = person.get("person_id") or str(uuid.uuid4())
        if person.get("person_id") != person_id:
            person["person_id"] = person_id
            self.people[idx] = person

        contact_id = contact.get("contact_id") or str(uuid.uuid4())
        contact["contact_id"] = contact_id
        existing_contact = None
        for existing in person.get("contacts", []) or []:
            if existing.get("contact_id") == contact_id:
                existing_contact = existing
                break

        if not workspace:
            if audio_removed:
                contact["audio_path"] = None
            contact.pop("previous_audio_path", None)
            if audio_source_is_temp and audio_source_path:
                try:
                    if os.path.isfile(audio_source_path):
                        os.remove(audio_source_path)
                except OSError:
                    pass
            return contact

        workspace_abs = os.path.abspath(workspace)
        contact_dir_rel = os.path.join("people", person_id, "contacts", contact_id)
        contact_dir_abs = os.path.join(workspace_abs, contact_dir_rel)
        os.makedirs(contact_dir_abs, exist_ok=True)

        transcript_rel = None
        transcript_candidates = [
            contact.get("transcript_path"),
            contact.get("transcription_path"),
            (existing_contact or {}).get("transcript_path"),
            (existing_contact or {}).get("transcription_path"),
        ]
        for candidate in transcript_candidates:
            if not candidate:
                continue
            try:
                transcript_rel, _ = self._normalise_workspace_path(workspace_abs, candidate)
            except ValueError:
                continue
            else:
                break

        journal_default = os.path.join("people", person_id, "contacts", contact_id, "journal.txt")
        try:
            journal_rel, journal_abs = self._normalise_workspace_path(workspace_abs, journal_default)
        except ValueError:
            journal_rel = journal_default.replace("\\", "/")
            journal_abs = os.path.normpath(os.path.join(workspace_abs, journal_rel))
        os.makedirs(os.path.dirname(journal_abs), exist_ok=True)
        try:
            with open(journal_abs, "w", encoding="utf-8") as tf:
                tf.write(contact.get("journal", "") or "")
        except Exception as exc:
            messagebox.showwarning("Transcription", f"Impossible d’écrire le journal : {exc}")

        contact["journal_path"] = journal_rel
        if transcript_rel:
            contact["transcript_path"] = transcript_rel
        else:
            contact.pop("transcript_path", None)
        contact.pop("transcription_path", None)

        previous_audio_rel = contact_dict.get("previous_audio_path")
        existing_audio_rel = (
            contact.get("audio_path")
            or (existing_contact or {}).get("audio_path")
            or previous_audio_rel
        )
        if audio_removed:
            self._cancel_transcription_job(contact_id)
            if existing_audio_rel:
                try:
                    _, existing_abs = self._normalise_workspace_path(workspace_abs, existing_audio_rel)
                except ValueError:
                    existing_abs = None
                if existing_abs and os.path.isfile(existing_abs):
                    try:
                        os.remove(existing_abs)
                    except OSError:
                        pass
            if transcript_rel and transcript_rel != journal_rel:
                try:
                    _, transcript_abs = self._normalise_workspace_path(workspace_abs, transcript_rel)
                except ValueError:
                    transcript_abs = None
                if transcript_abs and os.path.isfile(transcript_abs):
                    try:
                        os.remove(transcript_abs)
                    except OSError:
                        pass
            contact["audio_path"] = None
            contact.pop("transcript_path", None)
        elif audio_source_path:
            src_abs = os.path.abspath(audio_source_path)
            if not os.path.isfile(src_abs):
                messagebox.showwarning("Audio", "Le fichier audio sélectionné est introuvable.")
            else:
                try:
                    common = os.path.commonpath([workspace_abs, src_abs])
                except ValueError:
                    common = None

                source_rel = None
                if common == workspace_abs:
                    source_rel = os.path.relpath(src_abs, workspace_abs).replace("\\", "/")

                contact_prefix = os.path.join("people", person_id, "contacts", contact_id).replace("\\", "/")
                in_contact_dir = source_rel and (
                    source_rel == contact_prefix
                    or source_rel.startswith(contact_prefix + "/")
                )

                copied_abs = None
                if in_contact_dir:
                    contact["audio_path"] = source_rel
                else:
                    dest_name = os.path.basename(src_abs)
                    base, ext = os.path.splitext(dest_name)
                    dest_abs = os.path.join(contact_dir_abs, dest_name)
                    counter = 1
                    while os.path.exists(dest_abs):
                        dest_abs = os.path.join(contact_dir_abs, f"{base}_{counter}{ext}")
                        counter += 1
                    try:
                        shutil.copy2(src_abs, dest_abs)
                    except Exception as exc:
                        messagebox.showwarning("Audio", f"Impossible de copier l’audio : {exc}")
                    else:
                        audio_rel = os.path.relpath(dest_abs, workspace_abs).replace("\\", "/")
                        contact["audio_path"] = audio_rel
                        copied_abs = dest_abs

                if contact.get("audio_path") and existing_audio_rel and existing_audio_rel != contact["audio_path"]:
                    try:
                        _, existing_abs = self._normalise_workspace_path(workspace_abs, existing_audio_rel)
                    except ValueError:
                        existing_abs = None
                    if existing_abs and os.path.isfile(existing_abs):
                        try:
                            os.remove(existing_abs)
                        except OSError:
                            pass
                if audio_source_is_temp and contact.get("audio_path"):
                    try:
                        if os.path.isfile(src_abs):
                            if not copied_abs or os.path.abspath(copied_abs) != os.path.abspath(src_abs):
                                os.remove(src_abs)
                    except OSError:
                        pass
                if contact.get("audio_path"):
                    try:
                        _, audio_abs = self._normalise_workspace_path(workspace_abs, contact["audio_path"])
                    except ValueError:
                        audio_abs = None
                    if audio_abs and os.path.isfile(audio_abs):
                        transcript_output_default = os.path.join(
                            "people", person_id, "contacts", contact_id, "transcript.txt"
                        )
                        try:
                            transcript_output_rel, transcript_output_abs = self._normalise_workspace_path(
                                workspace_abs, transcript_output_default
                            )
                        except ValueError:
                            transcript_output_rel = transcript_output_default.replace("\\", "/")
                            transcript_output_abs = os.path.normpath(
                                os.path.join(workspace_abs, transcript_output_rel)
                            )
                        os.makedirs(os.path.dirname(transcript_output_abs), exist_ok=True)
                        if transcript_rel and transcript_rel != transcript_output_rel:
                            try:
                                _, old_trans_abs = self._normalise_workspace_path(workspace_abs, transcript_rel)
                            except ValueError:
                                old_trans_abs = None
                            if old_trans_abs and os.path.isfile(old_trans_abs):
                                try:
                                    os.remove(old_trans_abs)
                                except OSError:
                                    pass
                        if os.path.isfile(transcript_output_abs):
                            try:
                                os.remove(transcript_output_abs)
                            except OSError:
                                pass
                        self._cancel_transcription_job(contact_id)
                        display_date = contact.get("date") or "contact"
                        started = self._start_transcription_job(
                            contact_id=contact_id,
                            person_name=person.get("name", "?"),
                            person_index=idx,
                            display_label=f"{person.get('name', '?')} — {display_date}",
                            audio_abs=audio_abs,
                            transcript_rel=transcript_output_rel,
                            transcript_abs=transcript_output_abs,
                            journal_rel=journal_rel,
                            journal_abs=journal_abs,
                        )
                        if started:
                            contact["transcript_path"] = transcript_output_rel
                            transcript_rel = transcript_output_rel
                        else:
                            if transcript_rel:
                                contact["transcript_path"] = transcript_rel
                            else:
                                contact.pop("transcript_path", None)
        else:
            if existing_audio_rel:
                try:
                    audio_rel, _ = self._normalise_workspace_path(workspace_abs, existing_audio_rel)
                except ValueError:
                    contact["audio_path"] = None
                else:
                    contact["audio_path"] = audio_rel

        contact.pop("previous_audio_path", None)

        return contact

    def on_cancel_transcription(self):
        job = self._active_transcription or {}
        contact_id = job.get("contact_id")
        if not contact_id:
            return
        self._cancel_transcription_job(contact_id, user_request=True)

    def _resolve_vibe_executable(self):
        candidate = os.environ.get("VIBE_CLI", "vibe")
        candidate = os.path.expanduser(candidate)
        if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
            return candidate
        found = shutil.which(candidate)
        return found

    def _resolve_vibe_model_path(self):
        candidate = os.environ.get("VIBE_MODEL_PATH")
        if not candidate:
            return None
        candidate = os.path.expanduser(candidate)
        if os.path.isfile(candidate):
            return candidate
        if self.current_workspace and not os.path.isabs(candidate):
            workspace_candidate = os.path.join(self.current_workspace, candidate)
            if os.path.isfile(workspace_candidate):
                return workspace_candidate
        return None

    def _start_transcription_job(
        self,
        *,
        contact_id,
        person_name,
        person_index,
        display_label,
        audio_abs,
        transcript_rel,
        transcript_abs,
        journal_rel,
        journal_abs,
    ):
        vibe_exec = self._resolve_vibe_executable()
        if not vibe_exec:
            messagebox.showwarning(
                "Vibe",
                "Impossible de trouver le binaire 'vibe'. Configure VIBE_CLI ou ajoute-le au PATH.",
            )
            return False

        model_path = self._resolve_vibe_model_path()
        if not model_path:
            messagebox.showwarning(
                "Vibe",
                "Modèle Whisper introuvable. Définis VIBE_MODEL_PATH vers un fichier .bin téléchargé depuis Vibe.",
            )
            return False

        language = os.environ.get("VIBE_LANGUAGE", "french")
        threads_env = os.environ.get("VIBE_THREADS")
        threads = None
        if threads_env:
            try:
                threads = int(threads_env)
            except (TypeError, ValueError):
                threads = None

        cancel_event = threading.Event()
        job = {
            "contact_id": contact_id,
            "person_name": person_name,
            "person_index": person_index,
            "display_label": display_label,
            "audio_abs": audio_abs,
            "transcript_rel": transcript_rel,
            "transcript_abs": transcript_abs,
            "journal_rel": journal_rel,
            "journal_abs": journal_abs,
            "executable": vibe_exec,
            "model": model_path,
            "language": language,
            "threads": threads,
            "cancel_event": cancel_event,
            "process": None,
        }

        worker = threading.Thread(target=self._run_transcription_job, args=(job,), daemon=True)
        job["thread"] = worker
        self._transcription_jobs[contact_id] = job
        self._active_transcription = job
        self._update_transcription_panel(display_label, 0.0, "Initialisation de Vibe…", enable_cancel=True)
        self._set_progress_mode("indeterminate")
        worker.start()
        return True

    def _run_transcription_job(self, job):
        contact_id = job["contact_id"]
        cancel_event = job["cancel_event"]
        command = [
            job["executable"],
            "--file",
            job["audio_abs"],
            "--model",
            job["model"],
            "--format",
            "txt",
            "--write",
            job["transcript_abs"],
            "--language",
            job["language"],
        ]
        if job.get("threads"):
            command.extend(["--n-threads", str(job["threads"])])
        temperature = os.environ.get("VIBE_TEMPERATURE")
        if temperature:
            command.extend(["--temperature", temperature])

        env = os.environ.copy()
        env.setdefault("NO_COLOR", "1")

        try:
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                cwd=self.current_workspace or None,
                env=env,
            )
        except Exception as exc:
            self._transcription_queue.put(("error", contact_id, f"Échec du démarrage de Vibe : {exc}"))
            self._transcription_jobs.pop(contact_id, None)
            return

        job["process"] = process
        self._transcription_queue.put(("start", contact_id, job["display_label"]))
        output_tail = deque(maxlen=200)

        try:
            if process.stdout:
                for line in process.stdout:
                    if cancel_event.is_set():
                        break
                    if not line:
                        continue

                    normalized = line.replace("\r", "\n")
                    segments = [seg for seg in normalized.split("\n") if seg]
                    if not segments:
                        segments = [line]

                    for segment in segments:
                        cleaned_segment = segment.strip()
                        if cleaned_segment:
                            output_tail.append(cleaned_segment + "\n")
                        else:
                            output_tail.append(segment)

                        progress_value = self._extract_progress(segment)
                        if progress_value is not None:
                            self._transcription_queue.put((
                                "progress",
                                contact_id,
                                progress_value,
                                cleaned_segment or None,
                            ))
            if cancel_event.is_set() and process.poll() is None:
                try:
                    process.terminate()
                except Exception:
                    pass
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    try:
                        process.kill()
                    except Exception:
                        pass
            else:
                process.wait()
        finally:
            self._transcription_jobs.pop(contact_id, None)

        if cancel_event.is_set():
            self._transcription_queue.put(("cancelled", contact_id, job["display_label"]))
            return

        if process.returncode != 0:
            tail = "".join(output_tail).strip()
            if not tail:
                tail = "Transcription interrompue avec le code de sortie %s." % process.returncode
            self._transcription_queue.put(("error", contact_id, tail))
            return

        transcript_text = ""
        try:
            with open(job["transcript_abs"], "r", encoding="utf-8") as tf:
                transcript_text = tf.read()
        except FileNotFoundError:
            transcript_text = "".join(output_tail)
        except Exception as exc:
            self._transcription_queue.put(("error", contact_id, f"Lecture transcription impossible : {exc}"))
            return

        payload = {
            "person_index": job["person_index"],
            "transcript_rel": job["transcript_rel"],
            "transcript_abs": job["transcript_abs"],
            "journal_rel": job["journal_rel"],
            "journal_abs": job["journal_abs"],
            "text": transcript_text,
            "display_label": job["display_label"],
        }
        self._transcription_queue.put(("done", contact_id, payload))

    def _extract_progress(self, line):
        if not line:
            return None
        lower = line.lower()
        if "%" not in line:
            return None
        match = re.search(r"(?:progress|prog|percent)\s*[:=]?\s*(\d{1,3}(?:[\.,]\d+)?)\s*%", lower)
        if not match:
            match = re.search(r"(\d{1,3}(?:[\.,]\d+)?)\s*%", line)
            if match and "progress" not in lower:
                # Évite de confondre avec le texte final contenant un pourcentage.
                snippet = line.strip()
                if snippet and not snippet.lower().startswith("transcription"):
                    if "->" in snippet or snippet.count(":") >= 2:
                        return None
        if not match:
            return None
        try:
            raw_value = match.group(1).replace(",", ".")
            value = float(raw_value)
        except (TypeError, ValueError):
            return None
        if 0 <= value <= 100:
            return value
        return None

    def _cancel_transcription_job(self, contact_id, user_request=False):
        job = self._transcription_jobs.get(contact_id)
        if not job:
            return
        job["cancel_event"].set()
        process = job.get("process")
        if process and process.poll() is None:
            try:
                process.terminate()
            except Exception:
                pass
        if user_request:
            self._transcription_queue.put(("cancel-requested", contact_id, job.get("display_label")))

    def _poll_transcription_queue(self):
        try:
            while True:
                event = self._transcription_queue.get_nowait()
                self._handle_transcription_event(event)
        except Empty:
            pass
        finally:
            self.after(200, self._poll_transcription_queue)

    def _set_progress_mode(self, mode):
        current = self.transcription_progressbar.cget("mode")
        if mode == "indeterminate":
            if current != "indeterminate":
                self.transcription_progressbar.config(mode="indeterminate")
            self.transcription_progressbar.start(60)
        else:
            if current == "indeterminate":
                self.transcription_progressbar.stop()
                self.transcription_progressbar.config(mode="determinate")

    def _update_transcription_panel(self, label=None, progress=None, status=None, enable_cancel=None):
        if label is not None:
            self.transcription_contact_var.set(label)
        if progress is not None:
            try:
                self.transcription_progress_var.set(float(progress))
            except (TypeError, ValueError):
                pass
        if status is not None:
            self.transcription_status_var.set(status)
        if enable_cancel is not None:
            state = "normal" if enable_cancel else "disabled"
            self.transcription_cancel_btn.config(state=state)

    def _handle_transcription_event(self, event):
        if not event:
            return
        kind = event[0]
        contact_id = event[1] if len(event) > 1 else None

        if kind == "start":
            label = event[2]
            job = self._transcription_jobs.get(contact_id)
            if job:
                self._active_transcription = job
            self._set_progress_mode("indeterminate")
            self._update_transcription_panel(label, 0.0, "Modèle en cours de chargement…", enable_cancel=True)
        elif kind == "progress":
            percent = event[2]
            message = event[3] if len(event) > 3 else None
            if self._active_transcription and self._active_transcription.get("contact_id") == contact_id:
                self._set_progress_mode("determinate")
                numeric_percent = None
                if isinstance(percent, (int, float)):
                    numeric_percent = float(percent)
                else:
                    try:
                        numeric_percent = float(str(percent).replace(",", "."))
                    except (TypeError, ValueError):
                        numeric_percent = None

                if numeric_percent is not None:
                    if abs(numeric_percent - round(numeric_percent)) < 0.05:
                        status = f"{int(round(numeric_percent))}%"
                    else:
                        status = f"{numeric_percent:.1f}%"
                else:
                    status = f"{percent}%"

                if message and "progress" in message.lower():
                    status = message

                progress_value = numeric_percent if numeric_percent is not None else percent
                self._update_transcription_panel(None, progress_value, status, enable_cancel=True)
        elif kind == "cancel-requested":
            label = event[2] if len(event) > 2 else None
            if self._active_transcription and self._active_transcription.get("contact_id") == contact_id:
                self._set_progress_mode("indeterminate")
                self._update_transcription_panel(label, None, "Annulation en cours…", enable_cancel=False)
        elif kind == "cancelled":
            if self._active_transcription and self._active_transcription.get("contact_id") == contact_id:
                self._set_transcription_idle("Transcription annulée.")
        elif kind == "error":
            message = event[2] if len(event) > 2 else "Erreur inconnue"
            if self._active_transcription and self._active_transcription.get("contact_id") == contact_id:
                self._set_transcription_idle("Erreur lors de la transcription.")
            messagebox.showerror("Vibe", message)
        elif kind == "done":
            payload = event[2] if len(event) > 2 else {}
            self._apply_transcription_to_contact(contact_id, payload)
            if self._active_transcription and self._active_transcription.get("contact_id") == contact_id:
                self._set_transcription_idle("Transcription terminée.")

    def _set_transcription_idle(self, status_text=""):
        self._active_transcription = None
        self.transcription_progressbar.stop()
        self.transcription_progressbar.config(mode="determinate")
        self.transcription_progress_var.set(0.0)
        self.transcription_contact_var.set("Aucune transcription en cours.")
        self.transcription_status_var.set(status_text)
        self.transcription_cancel_btn.config(state="disabled")

    def _apply_transcription_to_contact(self, contact_id, payload):
        if not isinstance(payload, dict):
            return
        person_index = payload.get("person_index")
        if person_index is None or person_index < 0 or person_index >= len(self.people):
            return

        transcript_rel = payload.get("transcript_rel")
        transcript_abs = payload.get("transcript_abs")
        journal_rel = payload.get("journal_rel")
        journal_abs = payload.get("journal_abs")
        transcript_text = (payload.get("text") or "").strip()
        if not transcript_text:
            messagebox.showwarning("Vibe", "La transcription générée est vide.")
            return

        person = self.people[person_index]
        contacts = person.get("contacts", [])
        updated = False
        for idx, entry in enumerate(contacts):
            if entry.get("contact_id") != contact_id:
                continue
            contact = copy.deepcopy(entry)
            base_journal = contact.get("journal", "") or ""
            if TRANSCRIPT_HEADER in base_journal:
                base_part = base_journal.split(TRANSCRIPT_HEADER)[0].rstrip()
            else:
                base_part = base_journal.rstrip()
            if base_part:
                new_journal = base_part + TRANSCRIPT_HEADER + transcript_text
            else:
                new_journal = transcript_text
            contact["journal"] = new_journal
            if transcript_rel:
                contact["transcript_path"] = transcript_rel
            else:
                contact.pop("transcript_path", None)
            if journal_rel:
                contact["journal_path"] = journal_rel
            contacts[idx] = contact
            updated = True
            break

        if not updated:
            return

        person["contacts"] = contacts
        self.people[person_index] = person

        if transcript_abs:
            try:
                os.makedirs(os.path.dirname(transcript_abs), exist_ok=True)
                with open(transcript_abs, "w", encoding="utf-8") as tf:
                    tf.write(transcript_text)
            except Exception as exc:
                messagebox.showwarning("Vibe", f"Impossible d’écrire la transcription : {exc}")

        if journal_abs:
            try:
                os.makedirs(os.path.dirname(journal_abs), exist_ok=True)
                with open(journal_abs, "w", encoding="utf-8") as jf:
                    jf.write(self.people[person_index]["contacts"][idx]["journal"])
            except Exception as exc:
                messagebox.showwarning("Vibe", f"Impossible d’actualiser le journal : {exc}")

        if person_index == self.current_index:
            self._refresh_contact_row(contact_id)

        self.autosave()

    def _refresh_contact_row(self, contact_id):
        if self.current_index is None:
            return
        contact = self._get_contact_by_id(contact_id)
        if not contact:
            return
        for iid in self.contacts_tree.get_children():
            values = self.contacts_tree.item(iid, "values")
            if len(values) >= 6 and values[5] == contact_id:
                self.contacts_tree.item(
                    iid,
                    values=(
                        contact.get("date", ""),
                        contact.get("start", ""),
                        contact.get("end", ""),
                        contact.get("minutes", ""),
                        self._contact_preview_label(contact),
                        contact_id,
                    ),
                )
                break

    def on_close(self):
        for job in list(self._transcription_jobs.values()):
            job["cancel_event"].set()
            process = job.get("process")
            if process and process.poll() is None:
                try:
                    process.terminate()
                except Exception:
                    pass
        for job in list(self._transcription_jobs.values()):
            thread = job.get("thread")
            if thread and thread.is_alive():
                thread.join(timeout=1.5)
        self.destroy()

    # -------------------- Analyse: prompt & import patch --------------------
    def build_analysis_prompt(self, person):
        import copy, json as _json
        p = copy.deepcopy(person)
        # ordonner les contacts par date croissante pour lecture
        p["contacts"] = sorted(p.get("contacts", []), key=lambda c: c.get("date",""))
        instructions = (
            "Analyse relationnelle (convoi/attachement) — MODE LOG-FIRST.\n"
            "Tâches:\n"
            "1) Déduis, à partir des journaux de contacts, les éléments manquants: "
            "   canal (presentiel/audio/visio/texte), besoins nourris, valence par contact, tags utiles.\n"
            "2) Calcule need_support (minutes_per_week et occ_per_week) par besoin sur 28 jours (pondère selon canal si pertinent).\n"
            "3) Diagnostique ambivalence (>=2 interactions négatives sur 5 derniers contacts) et ROI relationnel.\n"
            "4) Propose mises à jour globales: intimacy, reliability, valence, match_besoins, needs_nourished; "
            "   propose aussi un plan de rituels réalistes (2–3) et des notes à ajouter.\n"
            "5) Si des infos manquent, ajoute un champ 'questions' avec des questions claires.\n\n"
            "IMPORTANT: Réponds strictement au format JSON PATCH ci-dessous (pas de texte hors JSON).\n"
        )
        schema = {
            "updates": {
                "intimacy": 0.0,
                "reliability": 0.0,
                "valence": 0.0,
                "energy_cost": 0.0,
                "match_besoins": 0.0,
                "needs_nourished": ["..."],
                "need_support": {"toucher": {"minutes_per_week": 0.0}},
                "ambivalent": False,
                "status": ""
            },
            "notes_append": "…",
            "plan": ["…"],
            "questions": ["…"],
            "contacts_updates": [
                {
                    "contact_id": "uuid",
                    "channel": "presentiel|audio|visio|texte",
                    "needs": ["…"],
                    "valence": 0.0,
                    "mood_after": 0.0,
                    "format": "…",
                    "note": "…",
                    "tags": ["…"]
                }
            ]
        }
        return (
            instructions
            + "\n\nDONNÉES PERSONNE (JSON):\n"
            + _json.dumps(p, ensure_ascii=False, indent=2)
            + "\n\nSCHÉMA DE SORTIE (EXEMPLE DE PATCH):\n"
            + _json.dumps(schema, ensure_ascii=False, indent=2)
        )

    def _try_copy_to_clipboard(self, text):
        """Copy *text* to the system clipboard as reliably as possible."""

        def _normalize(value):
            if not isinstance(value, str):
                return None
            return value.replace("\r\n", "\n")

        def _attempt_tk(use_utf8=False):
            try:
                self.clipboard_clear()
                if use_utf8:
                    self.tk.call("clipboard", "append", "-type", "UTF8_STRING", text)
                else:
                    self.clipboard_append(text)
                # Force the clipboard ownership to update immediately.
                self.update()
                stored = _normalize(self.clipboard_get())
                return stored == _normalize(text)
            except tk.TclError:
                return False

        if _attempt_tk(False) or _attempt_tk(True):
            return True

        try:  # pyperclip fallback (may rely on system tools)
            import pyperclip  # type: ignore
        except Exception:
            pyperclip = None  # type: ignore

        if pyperclip is not None:
            try:
                pyperclip.copy(text)
                stored = _normalize(pyperclip.paste())
                if stored == _normalize(text):
                    return True
            except Exception:
                pass

        def _run_command(cmd, input_bytes):
            try:
                completed = subprocess.run(
                    cmd,
                    input=input_bytes,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    check=False,
                )
                return completed.returncode == 0
            except Exception:
                return False

        enc_text = text.encode("utf-8")

        commands = []
        if sys.platform.startswith("darwin"):
            commands.append(["pbcopy"])
        elif sys.platform.startswith("win"):
            commands.append(["powershell", "-NoProfile", "-Command", "Set-Clipboard -Value ([Console]::In.ReadToEnd())"])
            commands.append(["clip.exe"])
        else:
            for tool in ("wl-copy", "xclip", "xsel"):
                exe = shutil.which(tool)
                if not exe:
                    continue
                if tool == "xclip":
                    commands.append([exe, "-selection", "clipboard"])
                elif tool == "xsel":
                    commands.append([exe, "--clipboard", "--input"])
                else:  # wl-copy
                    commands.append([exe])

        for cmd in commands:
            if _run_command(cmd, enc_text):
                # Attempt to verify with whichever API is available.
                if pyperclip is not None:
                    try:
                        stored = _normalize(pyperclip.paste())
                    except Exception:
                        stored = None
                else:
                    try:
                        stored = _normalize(self.clipboard_get())
                    except tk.TclError:
                        stored = None
                if stored == _normalize(text):
                    return True

        return False

    def _show_prompt_manual_window(self, text):
        win = tk.Toplevel(self)
        win.title("Prompt d’analyse — copie manuelle")
        win.geometry("900x620")
        win.transient(self)

        info = ttk.Label(
            win,
            text="Copie manuelle requise : sélectionne le texte ci-dessous (Ctrl/Cmd+A) puis Ctrl/Cmd+C.",
            wraplength=840,
            justify="left",
        )
        info.pack(anchor="w", padx=12, pady=(12, 6))

        text_box = tk.Text(win, wrap="word")
        text_box.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        text_box.insert("1.0", text)
        text_box.focus_set()
        text_box.tag_add("sel", "1.0", "end")

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=12, pady=(0, 12))

        def _retry_copy():
            if self._try_copy_to_clipboard(text):
                messagebox.showinfo(
                    "Copie réussie",
                    "Le prompt a finalement été copié dans le presse-papiers.",
                    parent=win,
                )
                win.destroy()

        ttk.Button(btns, text="Réessayer la copie", command=_retry_copy).pack(side="left")
        ttk.Button(btns, text="Fermer", command=win.destroy).pack(side="right")

    def on_copy_analysis_prompt(self):
        idx = self.current_index
        if idx is None:
            messagebox.showinfo("Info","Sélectionne une personne.")
            return
        self.update_current_person_from_form()
        self.autosave()
        p = self.people[idx]
        text = self.build_analysis_prompt(p)
        if self._try_copy_to_clipboard(text):
            messagebox.showinfo("Copié", "Prompt d’analyse (TOUS les logs) copié dans le presse-papier.")
        else:
            messagebox.showwarning(
                "Copie manuelle requise",
                "Impossible de copier automatiquement le prompt. Une fenêtre s’ouvre pour te permettre de le copier manuellement.",
            )
            self._show_prompt_manual_window(text)

    def on_import_analysis_json(self):
        idx = self.current_index
        if idx is None:
            messagebox.showinfo("Info", "Sélectionne une personne.")
            return

        win = tk.Toplevel(self)
        win.title("Importer analyse – Coller le JSON")
        win.transient(self)
        win.grab_set()

        ttk.Label(
            win,
            text=(
                "Colle ici le patch d’analyse JSON (tu peux modifier avant de valider).\n"
                "Tu peux aussi charger un fichier si besoin."
            ),
            wraplength=640,
            justify="left",
        ).pack(padx=12, pady=(12, 6), anchor="w")

        text = tk.Text(win, width=100, height=28, wrap="word")
        text.pack(fill="both", expand=True, padx=12, pady=(0, 6))

        try:
            clip = self.clipboard_get()
            if isinstance(clip, str) and clip.strip():
                text.insert("1.0", clip)
        except Exception:
            pass

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=12, pady=(0, 12))

        def _load_from_file():
            path = filedialog.askopenfilename(
                parent=win,
                filetypes=[("JSON", "*.json"), ("Tous", "*.*")],
            )
            if not path:
                return
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                text.delete("1.0", "end")
                text.insert("1.0", content)
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de lire le fichier: {e}", parent=win)

        def _apply_from_text():
            raw = text.get("1.0", "end").strip()
            if not raw:
                messagebox.showwarning("Vide", "Colle ou saisis un patch avant de valider.", parent=win)
                return
            try:
                patch = json.loads(raw)
            except Exception as e:
                messagebox.showerror("Erreur", f"JSON invalide: {e}", parent=win)
                return

            win.destroy()
            self._apply_analysis_patch(idx, patch)

        ttk.Button(btns, text="Charger depuis un fichier…", command=_load_from_file).pack(
            side="left"
        )
        ttk.Button(btns, text="Valider", command=_apply_from_text).pack(side="right", padx=(4, 0))
        ttk.Button(btns, text="Annuler", command=win.destroy).pack(side="right")

        text.focus_set()

    def _apply_analysis_patch(self, idx, patch):
        try:
            p = self.people[idx]

            # 1) updates globaux
            upd = patch.get("updates", {}) or {}
            for k in ["intimacy", "reliability", "valence", "match_besoins", "energy_cost"]:
                if k in upd:
                    # bornes raisonnables
                    lo, hi = (-2, 2) if k == "valence" else (0, 10 if k not in ("match_besoins", "energy_cost") else 2)
                    p[k] = clamp(upd[k], lo, hi)
            if "needs_nourished" in upd and isinstance(upd["needs_nourished"], list):
                p["needs_nourished"] = [str(x) for x in upd["needs_nourished"]]
            if "need_support" in upd and isinstance(upd["need_support"], dict):
                p["need_support"] = upd["need_support"]
            if "ambivalent" in upd:
                p["ambivalent"] = bool(upd["ambivalent"])
            if "status" in upd:
                p["status"] = str(upd["status"])

            # 2) notes/plan/questions
            add = patch.get("notes_append")
            if add:
                stamp = datetime.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"
                p["notes"] = (p.get("notes", "") + f"\n\n[Analyse {stamp}]\n" + str(add)).strip()
            plan = patch.get("plan")
            if plan and isinstance(plan, list) and plan:
                p["notes"] = (p.get("notes", "") + "\nPlan proposé:\n- " + "\n- ".join(map(str, plan))).strip()
            questions = patch.get("questions")
            if questions:
                p["notes"] = (p.get("notes", "") + "\nQuestions ouvertes:\n- " + "\n- ".join(map(str, questions))).strip()

            # 3) contacts_updates
            cupds = patch.get("contacts_updates") or []
            if cupds:
                # index contacts par id
                idx_by_id = {c.get("contact_id"): i for i, c in enumerate(p.get("contacts", []))}
                for cu in cupds:
                    cid = cu.get("contact_id")
                    if not cid or cid not in idx_by_id:
                        continue
                    i = idx_by_id[cid]
                    c = p["contacts"][i]
                    for field in ("channel", "needs", "valence", "mood_after", "format", "note", "tags"):
                        if field in cu:
                            c[field] = cu[field]
                    p["contacts"][i] = c

            self.people[idx] = p
            self.refresh_people_list()
            self.select_person(idx)
            self.autosave()
            messagebox.showinfo("OK", "Patch d’analyse importé et appliqué.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Import impossible: {e}")

    def update_preview(self):
        p = {
            "intimacy": self.var_intim.get(),
            "reliability": self.var_reli.get(),
            "valence": self.var_vale.get(),
            "energy_cost": self.var_enrg.get(),
            "match_besoins": self.var_match.get()
        }
        score = compute_priority(p)
        act = label_action(score)
        self.lbl_score.config(text=f"{score:.2f}")
        self.action_box.config(text=act, bg=ACTION_COLOR.get(act, "#666666"))
        try:
            self.lbl_intim_val.config(text=f"{float(self.var_intim.get()):.1f}")
            self.lbl_reli_val.config(text=f"{float(self.var_reli.get()):.1f}")
            self.lbl_vale_val.config(text=f"{float(self.var_vale.get()):.1f}")
            self.lbl_enrg_val.config(text=f"{float(self.var_enrg.get()):.1f}")
        except Exception:
            pass

    # -------------------- Export / Import global --------------------
    def on_export_json(self):
        if getattr(self, "current_index", None) is not None:
            self.update_current_person_from_form()
        if not self.people:
            messagebox.showwarning("Vide", "Ajoute au moins une personne avant d’exporter.")
            return
        path = filedialog.askdirectory(
            title="Exporter le workspace (sélectionne le dossier de destination)",
        )
        if not path:
            return
        if os.path.isdir(path) and os.listdir(path):
            if not messagebox.askyesno(
                "Dossier non vide",
                "Le dossier sélectionné contient déjà des fichiers. Continuer et écraser les données du workspace ?",
            ):
                return
        self.save_to_file(path, show_message=True)

    def on_import_json(self):
        if getattr(self, "current_index", None) is not None:
            self.update_current_person_from_form()
            self.autosave()
        path = filedialog.askdirectory(title="Ouvrir un workspace existant")
        if not path:
            return
        self.load_from_file(path, show_info=True)

    # -------------------- Divers --------------------
    def on_about(self):
        messagebox.showinfo(
            "À propos",
            "Convoi Wizard v3 – LOG-FIRST + Analyse via patch\n"
            "• Contacts: date/heure + journal libre (aucune case superflue)\n"
            "• Prompt d’analyse (TOUS logs) → import patch JSON\n"
            "• Le patch peut tout mettre à jour (globaux + par contact)"
        )

# --------------------------- Lancement ---------------------------

def main():
    """Launch the Convoi Wizard Tkinter application."""

    # Recharge l'éventuel fichier .env après une modification dynamique de CONVOI_ENV_FILE.
    _load_env_from_config()

    app = ConvoiWizard()
    app.mainloop()


if __name__ == "__main__":
    main()

