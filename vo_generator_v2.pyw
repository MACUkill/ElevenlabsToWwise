# -*- coding: utf-8 -*-
"""
============================================================
 VO Generator v2 — Thug Life  (Faza 0b)
============================================================
Spec: raport "Dialogue & VO System.md" (§6, §7.2, §14, Aneks B).

Excel (Casting/Vocabulary/linie) → walidacja → PLAN + podgląd
drzewka → [potwierdzenie] → ElevenLabs TTS/S2S → wav (loudness-
norm + marker-label=LineID) → WAAPI import do VOICEOVER →
eventy Play_* → 4 scalone CSV.

Konfiguracja: pipeline_config.json obok skryptu — struktura
drzewka (typy obiektów), targety LUFS, frazy wycinane z napisów,
ścieżki, ustawienia GUI. Tworzy się sam z defaultami raportu.
Zmiany struktury robimy ŚWIADOMIE w configu, nie klikaniem.

Zasady twarde:
  - LineID nieśmiertelne; plik = vo_{lineid}.wav; event = Play_+id.
  - Bark/Effort: Random Container per {Speaker}x{Intent}, 1 event.
  - Story/Conv: Sound Voice + event per linia.
  - Locked → Force_Regen → hash → cache (§14.0a); nic bez planu.

Zależności: pandas, openpyxl, requests, customtkinter, waapi-client.
Opcjonalnie: pyloudnorm + numpy.
============================================================
"""

import os, re, json, wave, struct, hashlib, random, threading
import requests
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from customtkinter import filedialog

try:
    from waapi import WaapiClient
    HAS_WAAPI = True
except Exception:
    HAS_WAAPI = False

try:
    import numpy as np
    import pyloudnorm as pyln
    HAS_LOUDNORM = True
except Exception:
    HAS_LOUDNORM = False

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "pipeline_config.json")

# ------------------------------------------------------------
# KONFIG — defaulty = żelazny klucz z raportu (Aneks B, §7.2)
# ------------------------------------------------------------
DEFAULT_CONFIG = {
    "api_key": "",
    "tts_model": "eleven_v3",
    "s2s_model": "eleven_multilingual_sts_v2",
    "format": "pcm_44100",
    "audio_root": "\\Actor-Mixer Hierarchy\\Default Work Unit\\VOICEOVER",
    "event_root": "\\Events\\Default Work Unit\\VOICEOVER",
    "ue_root": "/Game/WwiseAudio/Events",
    "loudnorm": True,
    "marker_offset_ms": 8,
    # Type -> gałąź w VOICEOVER (i lustrzana w Events)
    "type_tree": {
        "Bark":   ["Barks"],
        "Effort": ["Efforts"],
        "Story":  ["Story"],
        "ConvB":  ["Conversations", "Background"],
        "ConvH":  ["Conversations", "Highlighted"],
    },
    # typy obiektów Wwise (zmieniaj tylko świadomie — raport §7.2!)
    "tree_object": "ActorMixer",        # poziomy z type_tree
    "speaker_object": "ActorMixer",     # poziom {Speaker} (§7.2a: trim głosu)
    "container_object": "RandomSequenceContainer",  # {Intent} dla Bark/Effort
    "context_object": "Folder",         # {Context} dla Story/Conv
    "type_lufs": {"Bark": -18.0, "Effort": -16.0, "Story": -19.0,
                  "ConvB": -20.0, "ConvH": -19.0},
    # frazy wycinane z napisów (regexy, po jednej; kolejność ma znaczenie)
    "subtitle_strip_regex": [
        "\\[.*?\\]",                # [tagi reżyserskie]
        "\\*.*?\\*",                # *dźwięki*
        "(?:\\s*\\.\\.\\.)+\\s*$",  # ogony pauz TTS "... ..." na końcu
    ],
}

TYPES = ["Bark", "Effort", "Story", "ConvB", "ConvH"]
SEQUENCE_TYPES = ["Story", "ConvB", "ConvH"]
CONTAINER_TYPES = ["Bark", "Effort"]
EFFORT_POOL_SPEAKERS = ["Male", "Female"]
META_SHEETS = ["Casting", "Vocabulary"]
TYPE_CATEGORY = {"Bark": "Bark", "Story": "StoryScene",
                 "ConvB": "AmbientChatter", "ConvH": "StoryWorld"}
DEFAULT_POLICY = {"ConvB": "AbortWithReaction", "ConvH": "PauseResume",
                  "Story": "NotInterruptible"}
ID_RE = re.compile(r"^([A-Za-z0-9]+)_(Bark|Effort|Story|ConvB|ConvH)_([A-Za-z0-9]+)_(\d{2,3})$")


def load_config():
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))  # deep copy
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            user = json.load(f)
        for k, v in user.items():
            cfg[k] = v
    except Exception:
        pass
    return cfg


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


# ============================================================
# LOGIKA (bez GUI — testowalna)
# ============================================================

def norm(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def truthy(v):
    return norm(v).upper() in ("1", "TRUE", "YES", "T", "1.0")


def load_workbook_data(path):
    sheets = pd.read_excel(path, sheet_name=None)
    errors = []

    casting = {}
    if "Casting" not in sheets:
        errors.append("Brak arkusza Casting.")
    else:
        for _, r in sheets["Casting"].fillna("").iterrows():
            sp = norm(r.get("Speaker"))
            if not sp:
                continue
            casting[sp] = {
                "VoiceID": norm(r.get("Voice_ID")),
                "Gender": norm(r.get("Gender")),
                "EffortSet": norm(r.get("EffortSet")) or norm(r.get("Gender")),
                "Pool": norm(r.get("Pool")) or "Background",
                "ActorStatus": norm(r.get("ActorStatus")) or "AI",
                "DisplayName": norm(r.get("DisplayName")) or sp,
            }

    vocab = {"Intent": set(), "Quest": set(), "Topic": set()}
    if "Vocabulary" not in sheets:
        errors.append("Brak arkusza Vocabulary.")
    else:
        for _, r in sheets["Vocabulary"].fillna("").iterrows():
            for k in vocab:
                v = norm(r.get(k))
                if v:
                    vocab[k].add(v)

    lines = []
    for name, df in sheets.items():
        if name in META_SHEETS:
            continue
        df = df.fillna("")
        for i, r in df.iterrows():
            row = {
                "sheet": name, "row": i + 2,
                "LineID": norm(r.get("Line_ID")),
                "Speaker": norm(r.get("Speaker")),
                "Type": norm(r.get("Type")),
                "Context": norm(r.get("Context")),
                "Seq": int(float(r.get("Sequence_Order") or 0)),
                "Text": norm(r.get("Spoken_Text")),
                "S2S": norm(r.get("S2S_Reference")),
                "Policy": norm(r.get("InterruptPolicy")),
                "ResumeLineID": norm(r.get("ResumeLineID")),
                "PostDelay": float(r.get("PostDelay") or 0),
                "CooldownLine": float(r.get("CooldownLine") or 0),
                "CooldownIntent": float(r.get("CooldownIntent") or 0),
                "bCombatCritical": truthy(r.get("bCombatCritical")),
                "bPlayOnce": truthy(r.get("bPlayOnce")),
                "CategoryOverride": norm(r.get("CategoryOverride")),
                "Locked": truthy(r.get("Locked")),
                "ForceRegen": truthy(r.get("Force_Regen")),
            }
            if not row["Speaker"] and not row["Type"] and not row["Text"]:
                continue
            lines.append(row)

    return casting, vocab, lines, errors


def validate_line(row, casting, vocab, seen_ids, seen_seq):
    errs = []
    lid, sp, ty, ctx = row["LineID"], row["Speaker"], row["Type"], row["Context"]

    if not lid:
        errs.append("brak Line_ID (odpal Apps Script)")
    else:
        m = ID_RE.match(lid)
        if not m:
            errs.append(f"Line_ID '{lid}' nie pasuje do gramatyki")
        elif (m.group(1), m.group(2), m.group(3)) != (sp, ty, ctx):
            errs.append(f"Line_ID '{lid}' niespójne z kolumnami Speaker/Type/Context")
        if lid in seen_ids:
            errs.append(f"duplikat Line_ID (też w {seen_ids[lid]})")

    if ty not in TYPES:
        errs.append(f"Type '{ty}' spoza listy")
    if ty == "Effort":
        if sp not in casting and sp not in EFFORT_POOL_SPEAKERS:
            errs.append(f"Effort-Speaker '{sp}' nieznany")
    elif sp not in casting:
        errs.append(f"Speaker '{sp}' nie istnieje w Casting")

    if ty in ("Bark", "Effort") and vocab["Intent"] and ctx not in vocab["Intent"]:
        errs.append(f"Intent '{ctx}' spoza Vocabulary")
    if ty == "Story" and not re.match(r"^Q\d{2}", ctx):
        errs.append("Story-Context musi zaczynać się od Q##")
    if ty == "ConvB" and re.match(r"^Q\d{2}", ctx):
        errs.append("ConvB z kotwicą Q## — to powinno być ConvH")

    if ty in SEQUENCE_TYPES:
        if row["Seq"] < 1:
            errs.append("sekwencja wymaga Sequence_Order >= 1")
        key = (ty + "_" + ctx, row["Seq"])
        if key in seen_seq:
            errs.append(f"duplikat Sequence_Order={row['Seq']} w rozmowie {key[0]}")
        seen_seq[key] = True

    if not row["Text"] and not row["S2S"]:
        errs.append("wiersz bez treści")

    voice = casting.get(sp, {}).get("VoiceID", "")
    if not voice and ty != "Effort" and not row["S2S"]:
        errs.append(f"Speaker '{sp}' bez Voice_ID w Casting")
    return errs


def derive(row, casting, cfg):
    sp, ty, ctx, lid = row["Speaker"], row["Type"], row["Context"], row["LineID"]
    d = {}
    cat = row["CategoryOverride"] or TYPE_CATEGORY.get(ty, "")
    if ty == "Bark" and sp == "Player" and not row["CategoryOverride"]:
        cat = "PlayerVO"
    d["Category"] = cat
    d["ConversationID"] = (ty + "_" + ctx) if ty in SEQUENCE_TYPES else ""
    d["file"] = "vo_" + lid.lower() + ".wav"
    d["event"] = f"Play_{sp}_{ty}_{ctx}" if ty in CONTAINER_TYPES else f"Play_{lid}"
    d["tree"] = cfg["type_tree"][ty]
    d["voice"] = casting.get(sp, {}).get("VoiceID", "")
    return d


def content_hash(row, casting, tts_model, s2s_model, fmt):
    voice = casting.get(row["Speaker"], {}).get("VoiceID", "")
    basis = "|".join([row["Text"], row["S2S"], voice,
                      s2s_model if row["S2S"] else tts_model, fmt])
    return hashlib.sha1(basis.encode("utf-8")).hexdigest()


def build_plan(lines, casting, manifest, out_dir, cfg):
    tts_model, s2s_model, fmt = cfg["tts_model"], cfg["s2s_model"], cfg["format"]
    plan = []
    for row in lines:
        d = derive(row, casting, cfg)
        h = content_hash(row, casting, tts_model, s2s_model, fmt)
        path = os.path.join(out_dir, d["file"])
        exists = os.path.exists(path)
        old = manifest.get(row["LineID"], {})
        if row["Locked"]:
            action, why = "LOCKED", "final aktorski — nie tykam"
        elif row["ForceRegen"]:
            action, why = "REGEN", "Force_Regen"
        elif not exists:
            action, why = "NEW", "brak pliku"
        elif old.get("hash") != h:
            diff = []
            ob = (old.get("basis") or "").split("|")
            nb = [row["Text"], row["S2S"], d["voice"],
                  s2s_model if row["S2S"] else tts_model, fmt]
            names = ["tekst", "S2S ref", "głos", "model", "format"]
            for n, (a, b) in zip(names, zip(ob + [""] * 5, nb)):
                if a != b:
                    diff.append(n)
            action, why = "REGEN", "zmiana: " + (", ".join(diff) or "danych")
        else:
            action, why = "CACHE", ""
        plan.append({"row": row, "d": d, "hash": h, "path": path,
                     "action": action, "why": why})
    return plan


def build_tree_preview(plan, cfg, audio_root, event_root):
    """Tekstowy podgląd drzewka Wwise, które powstanie/zostanie użyte."""
    TAG = {"ActorMixer": "[AM]", "RandomSequenceContainer": "[RC]",
           "Folder": "[F] ", "Sound Voice": "[SV]", "Event": "[EV]"}
    audio, events = {}, {}
    for p in plan:
        if p["action"] == "LOCKED":
            continue
        row, d = p["row"], p["d"]
        node = audio
        chain = [(cfg["tree_object"], s) for s in d["tree"]]
        chain.append((cfg["speaker_object"], row["Speaker"]))
        if row["Type"] in CONTAINER_TYPES:
            chain.append((cfg["container_object"], row["Context"]))
        else:
            chain.append((cfg["context_object"], row["Context"]))
        chain.append(("Sound Voice", row["LineID"]))
        for otype, name in chain:
            node = node.setdefault((otype, name), {})
        enode = events
        for s in d["tree"] + [row["Speaker"]]:
            enode = enode.setdefault(("Folder", s), {})
        enode.setdefault(("Event", d["event"]), {})

    lines_out = []

    def walk(node, indent):
        for (otype, name), child in sorted(node.items(), key=lambda kv: (kv[0][0] != "Folder", kv[0][1])):
            lines_out.append("  " * indent + f"{TAG.get(otype, '[?]')} {name}")
            walk(child, indent + 1)

    lines_out.append(f"AUDIO: {audio_root}")
    walk(audio, 1)
    lines_out.append("")
    lines_out.append(f"EVENTY: {event_root}")
    walk(events, 1)
    return "\n".join(lines_out)


def strip_tags(text, patterns):
    t = text
    for pat in patterns:
        try:
            t = re.sub(pat, "", t)
        except re.error:
            pass
    return re.sub(r"\s+", " ", t).strip()


def add_wav_marker(path, label, offset_ms=8):
    with open(path, "rb") as f:
        data = f.read()
    if data[:4] != b"RIFF" or data[8:12] != b"WAVE":
        raise ValueError("nie-RIFF/WAVE: " + path)
    pos, sr = 12, None
    while pos + 8 <= len(data):
        cid = data[pos:pos + 4]
        size = struct.unpack("<I", data[pos + 4:pos + 8])[0]
        if cid == b"fmt ":
            sr = struct.unpack("<I", data[pos + 12:pos + 16])[0]
        pos += 8 + size + (size & 1)
    if not sr:
        raise ValueError("brak fmt: " + path)
    sample_off = int(sr * offset_ms / 1000)
    cue = struct.pack("<I", 1) + struct.pack("<II4sIII", 1, sample_off, b"data", 0, 0, sample_off)
    cue_chunk = b"cue " + struct.pack("<I", len(cue)) + cue
    lab = label.encode("utf-8") + b"\x00"
    if len(lab) & 1:
        lab += b"\x00"
    labl = struct.pack("<I", 1) + lab
    adtl = b"adtl" + b"labl" + struct.pack("<I", len(labl)) + labl
    list_chunk = b"LIST" + struct.pack("<I", len(adtl)) + adtl
    out = data + cue_chunk + list_chunk
    out = out[:4] + struct.pack("<I", len(out) - 8) + out[8:]
    with open(path, "wb") as f:
        f.write(out)


def loudness_normalize(path, target_lufs):
    if not HAS_LOUDNORM:
        return False
    with wave.open(path, "rb") as w:
        sr, n, sw = w.getframerate(), w.getnframes(), w.getsampwidth()
        raw = w.readframes(n)
        ch = w.getnchannels()
    if sw != 2:
        return False
    audio = np.frombuffer(raw, dtype=np.int16).astype(np.float64) / 32768.0
    if ch > 1:
        audio = audio.reshape(-1, ch).mean(axis=1)
    meter = pyln.Meter(sr)
    try:
        loud = meter.integrated_loudness(audio)
    except Exception:
        return False
    gain = 10 ** ((target_lufs - loud) / 20)
    peak = np.abs(audio).max() * gain
    if peak > 0.98:
        gain *= 0.98 / peak
    out = np.clip(audio * gain, -1, 1)
    with wave.open(path, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(sr)
        w.writeframes((out * 32767).astype(np.int16).tobytes())
    return True


def write_csvs(plan, casting, out_dir, ue_root, strip_patterns):
    dt_lines, st_rows, conv_rows = [], [], {}
    for p in plan:
        row, d = p["row"], p["d"]
        if row["Type"] == "Effort":
            continue  # Efforts poza DataTable (§6.2)
        ue_folder = "/".join(["VOICEOVER"] + d["tree"] + [row["Speaker"]])
        ev = d["event"]
        dt_lines.append({
            "---": row["LineID"], "Speaker": row["Speaker"], "Category": d["Category"],
            "Intent": row["Context"],
            "Event": f"{ue_root}/{ue_folder}/{ev}.{ev}",
            "SubtitleKey": row["LineID"],
            "CooldownLine": row["CooldownLine"], "CooldownIntent": row["CooldownIntent"],
            "bCombatCritical": row["bCombatCritical"], "bPlayOnce": row["bPlayOnce"],
            "ConversationID": d["ConversationID"], "SequenceOrder": row["Seq"],
            "PostDelay": row["PostDelay"],
        })
        st_rows.append({"Key": row["LineID"], "SourceString": strip_tags(row["Text"], strip_patterns)})
        cid = d["ConversationID"]
        if cid:
            c = conv_rows.setdefault(cid, {"---": cid, "InterruptPolicy": "", "ResumeLineID": ""})
            if row["Policy"]:
                c["InterruptPolicy"] = row["Policy"]
            if row["ResumeLineID"]:
                c["ResumeLineID"] = row["ResumeLineID"]
    for cid, c in conv_rows.items():
        if not c["InterruptPolicy"]:
            c["InterruptPolicy"] = DEFAULT_POLICY.get(cid.split("_")[0], "Abort")

    cast_rows = [{"---": sp, **c} for sp, c in casting.items()]
    paths = {}
    for fname, rows in [("DT_VO_Lines.csv", dt_lines), ("ST_VO.csv", st_rows),
                        ("DT_VO_Casting.csv", cast_rows),
                        ("DT_Conversations.csv", list(conv_rows.values()))]:
        fpath = os.path.join(out_dir, fname)
        pd.DataFrame(rows).to_csv(fpath, index=False, encoding="utf-8-sig")
        paths[fname] = fpath
    return paths


# ============================================================
# GUI
# ============================================================
class VOGen2(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("VO Generator v2 — Thug Life (VOICEOVER pipeline)")
        self.geometry("1360x900")
        self.cfg = load_config()
        self.data = None
        self.sheet_vars = {}
        self.xlsx_path = None
        self.out_dir = os.path.join(SCRIPT_DIR, "Output_Audio")
        self.ref_dir = os.path.join(SCRIPT_DIR, "S2S_Reference")
        os.makedirs(self.out_dir, exist_ok=True)
        os.makedirs(self.ref_dir, exist_ok=True)
        self.manifest_path = os.path.join(self.out_dir, "generation_manifest.json")
        self.manifest = self._load_manifest()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- persystencja ----------
    def _load_manifest(self):
        try:
            with open(self.manifest_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_manifest(self):
        with open(self.manifest_path, "w", encoding="utf-8") as f:
            json.dump(self.manifest, f, indent=1, ensure_ascii=False)

    def _collect_cfg(self):
        c = self.cfg
        c["api_key"] = self.api_entry.get().strip()
        c["tts_model"] = self.tts_model.get()
        c["s2s_model"] = self.s2s_model.get()
        c["format"] = self.fmt.get()
        c["audio_root"] = self.audio_root.get().strip().rstrip("\\")
        c["event_root"] = self.event_root.get().strip().rstrip("\\")
        c["ue_root"] = self.ue_root.get().strip().rstrip("/")
        c["loudnorm"] = bool(self.do_loudnorm.get())
        pats = [l.strip() for l in self.strip_box.get("0.0", "end").splitlines() if l.strip()]
        c["subtitle_strip_regex"] = pats
        return c

    def _on_close(self):
        try:
            save_config(self._collect_cfg())
        except Exception:
            pass
        self.destroy()

    # ---------- UI ----------
    def _build_ui(self):
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=8, pady=8)
        main = self.tabs.add("Główne")
        sett = self.tabs.add("Ustawienia")

        # ===== GŁÓWNE =====
        left = ctk.CTkFrame(main, width=380)
        left.pack(side="left", fill="y", padx=(0, 8), pady=4)
        right = ctk.CTkFrame(main)
        right.pack(side="left", fill="both", expand=True, pady=4)

        ctk.CTkButton(left, text="📄 Wczytaj xlsx", command=self.load_xlsx).pack(fill="x", padx=8, pady=(10, 2))
        self.file_lbl = ctk.CTkLabel(left, text="Brak pliku")
        self.file_lbl.pack(anchor="w", padx=10)
        ctk.CTkLabel(left, text="Arkusze:", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
        self.sheet_frame = ctk.CTkScrollableFrame(left, height=260)
        self.sheet_frame.pack(fill="x", padx=8, pady=4)
        row = ctk.CTkFrame(left, fg_color="transparent")
        row.pack(fill="x", padx=8)
        ctk.CTkButton(row, text="Wszystkie", command=lambda: self._set_sheets(True)).pack(side="left", expand=True, fill="x", padx=(0, 4))
        ctk.CTkButton(row, text="Żadne", command=lambda: self._set_sheets(False)).pack(side="left", expand=True, fill="x")

        self.dry_run = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(left, text="Dry-run (tylko walidacja + CSV)", variable=self.dry_run).pack(anchor="w", padx=10, pady=(10, 2))

        ctk.CTkButton(left, text="🌳 PODGLĄD DRZEWKA WWISE", fg_color="#2471A3",
                      command=lambda: self.run_threaded(preview_only=True)).pack(fill="x", padx=8, pady=(10, 2))
        self.go_btn = ctk.CTkButton(left, text="▶ PLAN → GENERUJ", height=52,
                                    font=("Arial", 16, "bold"), fg_color="#27AE60",
                                    hover_color="#1E8449", state="disabled",
                                    command=lambda: self.run_threaded(preview_only=False))
        self.go_btn.pack(fill="x", padx=8, pady=(4, 12))

        ctk.CTkLabel(right, text="Log / Plan / Drzewko:", font=("Arial", 12, "bold")).pack(anchor="w", padx=8, pady=(6, 0))
        self.log_box = ctk.CTkTextbox(right, font=("Consolas", 12))
        self.log_box.pack(fill="both", expand=True, padx=8, pady=8)

        # ===== USTAWIENIA =====
        s = ctk.CTkScrollableFrame(sett)
        s.pack(fill="both", expand=True, padx=4, pady=4)

        def lbl(text):
            ctk.CTkLabel(s, text=text, font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=(12, 2))

        lbl("ElevenLabs API Key")
        self.api_entry = ctk.CTkEntry(s, show="*")
        self.api_entry.insert(0, self.cfg["api_key"])
        self.api_entry.pack(fill="x", padx=10)

        lbl("Modele / format wyjściowy")
        rowm = ctk.CTkFrame(s, fg_color="transparent")
        rowm.pack(fill="x", padx=10)
        self.tts_model = ctk.StringVar(value=self.cfg["tts_model"])
        self.s2s_model = ctk.StringVar(value=self.cfg["s2s_model"])
        self.fmt = ctk.StringVar(value=self.cfg["format"])
        self.tts_menu = ctk.CTkOptionMenu(rowm, variable=self.tts_model, values=[self.cfg["tts_model"]])
        self.tts_menu.pack(side="left", expand=True, fill="x", padx=(0, 4))
        self.s2s_menu = ctk.CTkOptionMenu(rowm, variable=self.s2s_model, values=[self.cfg["s2s_model"]])
        self.s2s_menu.pack(side="left", expand=True, fill="x", padx=(0, 4))
        ctk.CTkOptionMenu(rowm, variable=self.fmt, values=["pcm_44100", "pcm_24000"], width=120).pack(side="left")
        ctk.CTkButton(s, text="Odśwież modele z API", fg_color="#8E44AD", command=self.fetch_models).pack(anchor="w", padx=10, pady=4)

        def path_row(label_text, initial, fetch=False):
            lbl(label_text)
            fr = ctk.CTkFrame(s, fg_color="transparent")
            fr.pack(fill="x", padx=10)
            ent = ctk.CTkEntry(fr)
            ent.insert(0, initial)
            ent.pack(side="left", expand=True, fill="x")
            if fetch:
                ctk.CTkButton(fr, text="⤓ Pobierz z Wwise", width=130, fg_color="#E07A22",
                              hover_color="#B56017",
                              command=lambda e=ent: self.fetch_wwise_path(e)).pack(side="left", padx=(6, 0))
            return ent

        self.audio_root = path_row("Wwise: Audio Root (Actor-Mixer Hierarchy → VOICEOVER)", self.cfg["audio_root"], fetch=True)
        self.event_root = path_row("Wwise: Events Root (Events → VOICEOVER)", self.cfg["event_root"], fetch=True)
        self.ue_root = path_row("Unreal: katalog eventów (WwiseAudio → Events)", self.cfg["ue_root"], fetch=False)

        lbl(f"Loudness-norm per Type (pyloudnorm: {'OK' if HAS_LOUDNORM else 'BRAK — py -m pip install pyloudnorm'})")
        self.do_loudnorm = ctk.BooleanVar(value=self.cfg["loudnorm"] and HAS_LOUDNORM)
        ctk.CTkCheckBox(s, text=f"Normalizuj surowe wav (targety: {self.cfg['type_lufs']})",
                        variable=self.do_loudnorm,
                        state="normal" if HAS_LOUDNORM else "disabled").pack(anchor="w", padx=10)

        lbl("Frazy wycinane z napisów (regex, jedna na linię)")
        self.strip_box = ctk.CTkTextbox(s, height=90, font=("Consolas", 12))
        self.strip_box.pack(fill="x", padx=10)
        self.strip_box.insert("0.0", "\n".join(self.cfg["subtitle_strip_regex"]))

        ctk.CTkLabel(s, text=f"Struktura drzewka (typy obiektów, gałęzie per Type): edytuj świadomie w pliku\n{CONFIG_PATH}\n— defaulty = raport §7.2 / Aneks B. Ustawienia zapisują się przy generacji i zamknięciu.",
                     justify="left", text_color="#999999").pack(anchor="w", padx=10, pady=14)
        ctk.CTkButton(s, text="💾 Zapisz ustawienia teraz", command=lambda: (save_config(self._collect_cfg()), self.log("Ustawienia zapisane."))).pack(anchor="w", padx=10, pady=(0, 14))

        self.log("VO Generator v2 gotowy. Wczytaj xlsx (zakładka Główne).")
        self.log(f"Config: {CONFIG_PATH}" + ("" if os.path.exists(CONFIG_PATH) else " (powstanie przy zapisie)"))
        if not HAS_WAAPI:
            self.log("UWAGA: brak waapi-client — możliwy tylko dry-run.")

    def log(self, msg):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.update()

    def _set_sheets(self, val):
        for v in self.sheet_vars.values():
            v.set(val)

    # ---------- akcje ----------
    def fetch_wwise_path(self, entry):
        if not HAS_WAAPI:
            self.log("⛔ brak waapi-client.")
            return
        try:
            with WaapiClient() as client:
                res = client.call("ak.wwise.ui.getSelectedObjects", {"options": {"return": ["path"]}})
                objs = res.get("objects") if res else None
                if objs:
                    path = objs[0]["path"]
                    entry.delete(0, "end")
                    entry.insert(0, path)
                    self.log(f"⤓ Pobrano z Wwise: {path}")
                else:
                    self.log("⛔ Nic nie zaznaczono w Wwise.")
        except Exception as e:
            self.log(f"⛔ Brak połączenia z Wwise ({e.__class__.__name__}).")

    def fetch_models(self):
        key = self.api_entry.get().strip()
        if not key:
            self.log("BŁĄD: wpisz API key (Ustawienia).")
            return
        try:
            r = requests.get("https://api.elevenlabs.io/v1/models", headers={"xi-api-key": key}, timeout=15)
            r.raise_for_status()
            models = r.json()
            tts = [m["model_id"] for m in models if m.get("can_do_text_to_speech")]
            s2s = [m["model_id"] for m in models if m.get("can_do_voice_conversion")]
            if tts:
                self.tts_menu.configure(values=tts)
                self.tts_model.set(next((m for m in tts if "v3" in m), tts[0]))
            if s2s:
                self.s2s_menu.configure(values=s2s)
                self.s2s_model.set(next((m for m in s2s if "sts" in m), s2s[0]))
            self.log(f"Modele: {len(tts)} TTS, {len(s2s)} S2S.")
        except Exception as e:
            self.log(f"Błąd API: {e}")

    def load_xlsx(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        self.xlsx_path = path
        casting, vocab, lines, errors = load_workbook_data(path)
        self.data = (casting, vocab, lines)
        for w in self.sheet_frame.winfo_children():
            w.destroy()
        self.sheet_vars.clear()
        for name in sorted({l["sheet"] for l in lines}):
            var = ctk.BooleanVar(value=True)
            self.sheet_vars[name] = var
            ctk.CTkCheckBox(self.sheet_frame, text=name, variable=var).pack(anchor="w", pady=1)
        self.file_lbl.configure(text=os.path.basename(path))
        self.log(f"\nWczytano: {os.path.basename(path)} — Casting: {len(casting)}, "
                 f"Intenty: {len(vocab['Intent'])}, linii: {len(lines)}.")
        for e in errors:
            self.log("⛔ " + e)
        self.go_btn.configure(state="normal" if not errors else "disabled")

    def run_threaded(self, preview_only=False):
        if not self.data:
            self.log("Najpierw wczytaj xlsx.")
            return
        self.go_btn.configure(state="disabled", text="⏳ PRACUJĘ…")
        threading.Thread(target=self._run_safe, args=(preview_only,), daemon=True).start()

    def _run_safe(self, preview_only):
        try:
            self.run(preview_only)
        except Exception as e:
            self.log(f"⛔ BŁĄD KRYTYCZNY: {e!r}")
        finally:
            self.go_btn.configure(state="normal", text="▶ PLAN → GENERUJ")

    # ---------- główny przebieg ----------
    def _validate_and_plan(self, cfg):
        casting, vocab, lines = self.data
        active = [l for l in lines if self.sheet_vars.get(l["sheet"]) and self.sheet_vars[l["sheet"]].get()]
        self.log("\n=== WALIDACJA ===")
        seen_ids, seen_seq, valid, n_err = {}, {}, [], 0
        for row in active:
            errs = validate_line(row, casting, vocab, seen_ids, seen_seq)
            if row["LineID"]:
                seen_ids[row["LineID"]] = f"{row['sheet']} w.{row['row']}"
            if errs:
                n_err += len(errs)
                self.log(f"⛔ [{row['sheet']} w.{row['row']}] " + "; ".join(errs))
            else:
                valid.append(row)
        self.log(f"Wierszy OK: {len(valid)} / {len(active)}; błędów: {n_err}.")
        if n_err:
            return None
        return build_plan(valid, casting, self.manifest, self.out_dir, cfg)

    def run(self, preview_only=False):
        cfg = self._collect_cfg()
        save_config(cfg)
        casting, vocab, lines = self.data

        plan = self._validate_and_plan(cfg)
        if plan is None:
            self.log("Popraw arkusz i spróbuj znowu. NIC nie wygenerowano.")
            return

        if preview_only:
            self.log("\n=== DRZEWKO WWISE (podgląd — nic nie tworzę) ===")
            self.log(build_tree_preview(plan, cfg, cfg["audio_root"], cfg["event_root"]))
            return

        counts = {}
        for p in plan:
            counts[p["action"]] = counts.get(p["action"], 0) + 1
        self.log("\n=== PLAN ===")
        self.log(f"🆕 nowe: {counts.get('NEW',0)}   🔁 regen: {counts.get('REGEN',0)}   "
                 f"💾 cache: {counts.get('CACHE',0)}   🔒 locked: {counts.get('LOCKED',0)}")
        for p in plan:
            if p["action"] == "REGEN":
                self.log(f"  🔁 {p['row']['LineID']} — {p['why']}")
            elif p["action"] == "LOCKED":
                self.log(f"  🔒 {p['row']['LineID']} — pomijam (final)")
        still_flagged = [p["row"]["LineID"] for p in plan
                         if p["row"]["ForceRegen"] and p["action"] == "REGEN"]
        to_generate = [p for p in plan if p["action"] in ("NEW", "REGEN")]

        if self.dry_run.get():
            self.log("\nDRY-RUN: pomijam audio i Wwise — piszę tylko CSV.")
            paths = write_csvs(plan, casting, self.out_dir, cfg["ue_root"], cfg["subtitle_strip_regex"])
            for v in paths.values():
                self.log(f"  📄 {v}")
            self.log("✅ Dry-run zakończony.")
            return

        if to_generate:
            if not messagebox.askyesno("Plan generacji",
                    f"Nowych: {counts.get('NEW',0)}\nRegeneracji: {counts.get('REGEN',0)}\n"
                    f"Cache (pomijam): {counts.get('CACHE',0)}\nLocked: {counts.get('LOCKED',0)}\n\n"
                    f"Odpalić ElevenLabs dla {len(to_generate)} linii?"):
                self.log("Anulowano — nic nie wygenerowano.")
                return

        self.log("\n=== AUDIO ===")
        for p in to_generate:
            row, d = p["row"], p["d"]
            if not self._make_audio(row, d, p["path"], cfg):
                p["action"] = "FAILED"
                continue
            if cfg["loudnorm"] and HAS_LOUDNORM:
                loudness_normalize(p["path"], cfg["type_lufs"].get(row["Type"], -18.0))
            add_wav_marker(p["path"], row["LineID"], cfg["marker_offset_ms"])
            self.manifest[row["LineID"]] = {
                "hash": p["hash"], "file": d["file"],
                "basis": "|".join([row["Text"], row["S2S"], d["voice"],
                                   cfg["s2s_model"] if row["S2S"] else cfg["tts_model"],
                                   cfg["format"]]),
            }
            self._save_manifest()

        failed = [p for p in plan if p["action"] == "FAILED"]
        if failed:
            self.log(f"⚠️ Nieudane generacje: {len(failed)} — pomijam w imporcie.")

        importable = [p for p in plan if p["action"] in ("NEW", "REGEN", "CACHE")
                      and os.path.exists(p["path"])]
        self.log("\n=== WWISE (WAAPI) ===")
        if not HAS_WAAPI:
            self.log("⛔ brak waapi-client — pomijam import.")
        else:
            try:
                self._wwise_import(importable, cfg)
            except Exception as e:
                self.log(f"⛔ WAAPI: {e!r}")
                return

        self.log("\n=== CSV ===")
        paths = write_csvs(plan, casting, self.out_dir, cfg["ue_root"], cfg["subtitle_strip_regex"])
        for v in paths.values():
            self.log(f"  📄 {v}")

        if still_flagged:
            self.log("\n⚠️ PRZYPOMNIENIE: odznacz Force_Regen w arkuszu dla:")
            for lid in still_flagged:
                self.log("   • " + lid)
        self.log("\n🚀 GOTOWE.")

    # ---------- audio ----------
    def _make_audio(self, row, d, out_path, cfg):
        key, fmt = cfg["api_key"], cfg["format"]
        try:
            if row["S2S"]:
                ref_dir = os.path.join(self.ref_dir, row["S2S"])
                refs = [f for f in os.listdir(ref_dir)
                        if f.lower().endswith((".wav", ".mp3"))] if os.path.isdir(ref_dir) else []
                if not refs:
                    self.log(f"⛔ {row['LineID']}: brak referencji S2S w {row['S2S']}")
                    return False
                ref = os.path.join(ref_dir, random.choice(refs))
                url = f"https://api.elevenlabs.io/v1/speech-to-speech/{d['voice']}?output_format={fmt}"
                with open(ref, "rb") as fa:
                    r = requests.post(url, headers={"xi-api-key": key},
                                      data={"model_id": cfg["s2s_model"]},
                                      files={"audio": (os.path.basename(ref), fa, "audio/wav")},
                                      timeout=120)
            else:
                url = f"https://api.elevenlabs.io/v1/text-to-speech/{d['voice']}?output_format={fmt}"
                r = requests.post(url, json={"text": row["Text"], "model_id": cfg["tts_model"]},
                                  headers={"xi-api-key": key}, timeout=120)
            if r.status_code != 200:
                self.log(f"⛔ {row['LineID']}: {r.status_code} {r.text[:120]}")
                return False
            sr = int(fmt.split("_")[1])
            with wave.open(out_path, "wb") as w:
                w.setnchannels(1)
                w.setsampwidth(2)
                w.setframerate(sr)
                w.writeframes(r.content)
            self.log(f"🎙️ {row['LineID']} ({'S2S' if row['S2S'] else 'TTS'})")
            return True
        except Exception as e:
            self.log(f"⛔ {row['LineID']}: {e!r}")
            return False

    # ---------- wwise ----------
    def _wwise_import(self, plan_items, cfg):
        audio_root, event_root = cfg["audio_root"], cfg["event_root"]
        created_events = set()
        with WaapiClient() as client:
            def ensure(parent, otype, name, props=None):
                args = {"parent": parent, "type": otype, "name": name,
                        "onNameConflict": "merge"}
                if props:
                    args.update(props)
                client.call("ak.wwise.core.object.create", args)
                return parent + "\\" + name

            for p in plan_items:
                row, d = p["row"], p["d"]
                sp, ty, ctx, lid = row["Speaker"], row["Type"], row["Context"], row["LineID"]

                cur = audio_root
                for seg in d["tree"]:
                    cur = ensure(cur, cfg["tree_object"], seg)
                cur = ensure(cur, cfg["speaker_object"], sp)
                if ty in CONTAINER_TYPES:
                    cur = ensure(cur, cfg["container_object"], ctx, {"@RandomOrSequence": 1})
                else:
                    cur = ensure(cur, cfg["context_object"], ctx)

                client.call("ak.wwise.core.audio.import", {
                    "importOperation": "useExisting",
                    "default": {"importLanguage": "English(US)"},
                    "imports": [{
                        "objectPath": f"{cur}\\<Sound Voice>{lid}",
                        "audioFile": p["path"],
                        "@Notes": strip_tags(row["Text"], cfg["subtitle_strip_regex"]),
                    }],
                })

                ev = d["event"]
                if ev in created_events:
                    continue
                ecur = event_root
                for seg in d["tree"]:
                    ecur = ensure(ecur, "Folder", seg)
                ecur = ensure(ecur, "Folder", sp)
                target = cur if ty in CONTAINER_TYPES else f"{cur}\\{lid}"
                client.call("ak.wwise.core.object.create", {
                    "parent": ecur, "type": "Event", "name": ev,
                    "onNameConflict": "merge",
                    "children": [{"name": "", "type": "Action",
                                  "@ActionType": 1, "@Target": target}],
                })
                created_events.add(ev)
                self.log(f"  🎯 {ev}")
        self.log(f"Import: {len(plan_items)} plików, {len(created_events)} eventów.")


if __name__ == "__main__":
    VOGen2().mainloop()
