# -*- coding: utf-8 -*-
"""
============================================================
 VO Generator v2  (Phase 0b)
============================================================
Spec: report "Dialogue & VO System.md" (§6, §7.2, §14, Appendix B).

Excel (Casting/Vocabulary/lines) → validation → PLAN + tree
preview → [confirmation] → ElevenLabs TTS/S2S → wav (loudness-
norm + marker-label=LineID) → WAAPI import into VOICEOVER →
play_vo_* events → 4 merged CSVs.

Configuration: pipeline_config.json next to the script - tree
structure (object types), LUFS targets, phrases stripped from subtitles,
paths, GUI settings. Created automatically with the report's defaults.
Structure changes are made DELIBERATELY in the config, not by clicking.

Hard rules:
  - LineID is immortal; file = vo_{lineid}.wav; event = play_+stem.
  - Bark/Effort: Random Container per {Speaker}x{Intent}, 1 event.
  - Story/Conv: Sound Voice + event per line.
  - Locked → Force_Regen → hash → cache (§14.0a); nothing without a plan.

Dependencies: pandas, openpyxl, requests, customtkinter, waapi-client.
Optional: pyloudnorm + numpy.
============================================================
"""

import os, re, json, math, wave, struct, hashlib, random, threading
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
# CONFIG - defaults = the ironclad naming key from the report (Appendix B, §7.2)
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
    # Type -> branch in VOICEOVER (mirrored in Events)
    "type_tree": {
        "Bark":   ["Barks"],
        "Effort": ["Efforts"],
        "Story":  ["Story"],
        "ConvB":  ["Conversations", "Background"],
        "ConvH":  ["Conversations", "Highlighted"],
    },
    # Wwise object types (change only deliberately - report §7.2!)
    "tree_object": "ActorMixer",        # levels from type_tree
    "speaker_object": "ActorMixer",     # {Speaker} level (§7.2a: voice trim)
    "container_object": "RandomSequenceContainer",  # {Intent} for Bark/Effort
    "context_object": "Folder",         # {Context} for Story/Conv
    "type_lufs": {"Bark": -18.0, "Effort": -16.0, "Story": -19.0,
                  "ConvB": -20.0, "ConvH": -19.0},
    # phrases stripped from subtitles (regexes, one each; order matters)
    "subtitle_strip_regex": [
        "\\[.*?\\]",                # [director tags]
        "\\*.*?\\*",                # *sounds*
        "(?:\\s*\\.\\.\\.)+\\s*$",  # trailing TTS pause tails "... ..." at the end
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
# LOGIC (no GUI - testable)
# ============================================================

def norm(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def truthy(v):
    return norm(v).upper() in ("1", "TRUE", "YES", "T", "1.0")


NUMBER_COLUMNS = ["Sequence_Order", "PostDelay", "CooldownLine", "CooldownIntent"]


def to_number(v):
    """Excel cell -> float; empty = 0, unparseable = None (reported by validation)."""
    s = norm(v)
    if not s:
        return 0.0
    try:
        f = float(s.replace(",", "."))
    except ValueError:
        return None
    return f if math.isfinite(f) else None


def load_workbook_data(path):
    sheets = pd.read_excel(path, sheet_name=None)
    errors = []

    casting = {}
    if "Casting" not in sheets:
        errors.append("Missing Casting sheet.")
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
        errors.append("Missing Vocabulary sheet.")
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
            nums = {c: to_number(r.get(c)) for c in NUMBER_COLUMNS}
            row = {
                "sheet": name, "row": i + 2,
                "LineID": norm(r.get("Line_ID")),
                "Speaker": norm(r.get("Speaker")),
                "Type": norm(r.get("Type")),
                "Context": norm(r.get("Context")),
                "Seq": int(nums["Sequence_Order"] or 0),
                "Text": norm(r.get("Spoken_Text")),
                "S2S": norm(r.get("S2S_Reference")),
                "Policy": norm(r.get("InterruptPolicy")),
                "ResumeLineID": norm(r.get("ResumeLineID")),
                "PostDelay": nums["PostDelay"] or 0.0,
                "CooldownLine": nums["CooldownLine"] or 0.0,
                "CooldownIntent": nums["CooldownIntent"] or 0.0,
                "BadNumbers": [c for c, v in nums.items() if v is None],
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

    for col in row.get("BadNumbers", []):
        errs.append(f"{col} is not a number")

    if not lid:
        errs.append("missing Line_ID (run the Apps Script)")
    else:
        m = ID_RE.match(lid)
        if not m:
            errs.append(f"Line_ID '{lid}' does not match the grammar")
        elif (m.group(1), m.group(2), m.group(3)) != (sp, ty, ctx):
            errs.append(f"Line_ID '{lid}' inconsistent with the Speaker/Type/Context columns")
        if lid in seen_ids:
            errs.append(f"duplicate Line_ID (also in {seen_ids[lid]})")

    if ty not in TYPES:
        errs.append(f"Type '{ty}' not in the list")
    if ty == "Effort":
        if sp not in casting and sp not in EFFORT_POOL_SPEAKERS:
            errs.append(f"Effort-Speaker '{sp}' unknown")
    elif sp not in casting:
        errs.append(f"Speaker '{sp}' does not exist in Casting")

    if ty in ("Bark", "Effort") and vocab["Intent"] and ctx not in vocab["Intent"]:
        errs.append(f"Intent '{ctx}' not in Vocabulary")
    if ty == "Story" and not re.match(r"^Q\d{2}", ctx):
        errs.append("Story-Context must start with Q##")
    if ty == "ConvB" and re.match(r"^Q\d{2}", ctx):
        errs.append("ConvB with a Q## anchor - this should be ConvH")

    if ty in SEQUENCE_TYPES:
        if row["Seq"] < 1:
            errs.append("sequence requires Sequence_Order >= 1")
        key = (ty + "_" + ctx, row["Seq"])
        if key in seen_seq:
            errs.append(f"duplicate Sequence_Order={row['Seq']} in conversation {key[0]}")
        seen_seq[key] = True

    if not row["Text"] and not row["S2S"]:
        errs.append("row without content")

    voice = casting.get(sp, {}).get("VoiceID", "")
    if not voice and ty != "Effort" and not row["S2S"]:
        errs.append(f"Speaker '{sp}' has no Voice_ID in Casting")
    return errs


def container_name(row):
    """Bark/Effort variant container (= event target), named by the file key."""
    return f"vo_{row['Speaker']}_{row['Type']}_{row['Context']}".lower()


def event_name(row):
    """Event = "play_" + file stem (Appendix B: Wwise/file world = lowercase;
    bark without the NN part - the event points at the variant container).
    Single source for the import, the CSVs and the orphan diff - if they
    disagree, the diff reports every generated event as an orphan."""
    if row["Type"] in CONTAINER_TYPES:
        return "play_" + container_name(row)
    return "play_vo_" + row["LineID"].lower()


def derive(row, casting, cfg):
    sp, ty, ctx, lid = row["Speaker"], row["Type"], row["Context"], row["LineID"]
    d = {}
    cat = row["CategoryOverride"] or TYPE_CATEGORY.get(ty, "")
    if ty == "Bark" and sp == "Player" and not row["CategoryOverride"]:
        cat = "PlayerVO"
    d["Category"] = cat
    d["ConversationID"] = (ty + "_" + ctx) if ty in SEQUENCE_TYPES else ""
    d["file"] = "vo_" + lid.lower() + ".wav"
    d["event"] = event_name(row)
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
            action, why = "LOCKED", "actor final - not touching"
        elif row["ForceRegen"]:
            action, why = "REGEN", "Force_Regen"
        elif not exists:
            action, why = "NEW", "file missing"
        elif old.get("hash") != h:
            diff = []
            ob = (old.get("basis") or "").split("|")
            nb = [row["Text"], row["S2S"], d["voice"],
                  s2s_model if row["S2S"] else tts_model, fmt]
            names = ["text", "S2S ref", "voice", "model", "format"]
            for n, (a, b) in zip(names, zip(ob + [""] * 5, nb)):
                if a != b:
                    diff.append(n)
            action, why = "REGEN", "changed: " + (", ".join(diff) or "data")
        else:
            action, why = "CACHE", ""
        plan.append({"row": row, "d": d, "hash": h, "path": path,
                     "action": action, "why": why})
    return plan


def build_tree_preview(plan, cfg, audio_root, event_root):
    """Text preview of the Wwise tree that will be created/used."""
    TAG = {"ActorMixer": "[AM]", "PropertyContainer": "[PC]", "RandomSequenceContainer": "[RC]",
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
            # container = event target -> key-based name (rule: simple structure,
            # playable objects named by the file key)
            chain.append((cfg["container_object"], container_name(row)))
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
    lines_out.append(f"EVENTS: {event_root}")
    walk(events, 1)
    return "\n".join(lines_out)


# ---------------- Orphans: Wwise vs Excel diff ----------------

MANAGED_TYPES = ["Bark", "Story", "ConvB", "ConvH"]  # Efforts = manual, outside the diff
DIFF_OBJ_TYPES = {"ActorMixer", "PropertyContainer", "RandomSequenceContainer", "SwitchContainer",
                  "BlendContainer", "Folder", "Sound", "Event"}


def expected_wwise_paths(lines, cfg, audio_root, event_root):
    """Full set of paths that SHOULD exist according to Excel (objects + events).
    Computed from ALL rows with a LineID (not just the checked sheets!) -
    otherwise an unchecked sheet would report false orphans."""
    paths = set()
    for row in lines:
        if row["Type"] not in MANAGED_TYPES or not row["LineID"]:
            continue
        tree = cfg["type_tree"].get(row["Type"], [])
        cur = audio_root
        for seg in tree:
            cur += "\\" + seg
            paths.add(cur)
        cur += "\\" + row["Speaker"]
        paths.add(cur)
        if row["Type"] in CONTAINER_TYPES:
            cur += "\\" + container_name(row)
        else:
            cur += "\\" + row["Context"]
        paths.add(cur)
        paths.add(cur + "\\" + row["LineID"])  # Sound Voice

        ev = event_name(row)
        ecur = event_root
        for seg in tree:
            ecur += "\\" + seg
            paths.add(ecur)
        ecur += "\\" + row["Speaker"]
        paths.add(ecur)
        paths.add(ecur + "\\" + ev)
    return paths


def managed_roots(cfg, audio_root, event_root):
    """Roots of the branches managed by the generator + the set of protected paths
    (the roots themselves are never deleted - they are structure, not content)."""
    roots, protected = {}, set()
    for ty in MANAGED_TYPES:
        tree = cfg["type_tree"].get(ty, [])
        if not tree:
            continue
        for base in (audio_root, event_root):
            cur = base
            for seg in tree:
                cur += "\\" + seg
                protected.add(cur)
            # unique roots (Conversations appears once despite ConvB/ConvH)
            full = base + "\\" + tree[0]
            roots[full] = (base, tree[0], full)
    return list(roots.values()), protected


def find_orphans(actual_objects, expected, protected):
    """actual_objects: [{path,type,id,...}] from WAQL. Returns TOP-MOST orphans
    (deleting a parent takes its children - without top-most there would be double deletion)."""
    # Wwise object names are case-insensitive - compare the same way, otherwise
    # a mere case difference vs Excel becomes a false orphan (and gets deleted).
    known = {p.lower() for p in expected} | {p.lower() for p in protected}
    orphan_paths = {}
    for o in actual_objects:
        p, t = o.get("path", ""), o.get("type", "")
        if not p or t not in DIFF_OBJ_TYPES:
            continue  # AudioFileSource/Action etc. go with the parent
        if p.lower() in known:
            continue
        orphan_paths[p] = o
    topmost = []
    for p, o in orphan_paths.items():
        parent = p.rsplit("\\", 1)[0]
        if parent not in orphan_paths:
            topmost.append(o)
    topmost.sort(key=lambda o: o.get("path", ""))
    return topmost


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
        raise ValueError("not RIFF/WAVE: " + path)
    pos, sr = 12, None
    while pos + 8 <= len(data):
        cid = data[pos:pos + 4]
        size = struct.unpack("<I", data[pos + 4:pos + 8])[0]
        if cid == b"fmt ":
            sr = struct.unpack("<I", data[pos + 12:pos + 16])[0]
        pos += 8 + size + (size & 1)
    if not sr:
        raise ValueError("missing fmt: " + path)
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
    if not np.isfinite(loud):
        return False  # silence = -inf LUFS -> infinite gain would write garbage
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
            continue  # Efforts stay out of the DataTable (§6.2)
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
        self.title("VO Generator v2 (VOICEOVER pipeline)")
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

    # ---------- persistence ----------
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
        main = self.tabs.add("Main")
        sett = self.tabs.add("Settings")

        # ===== MAIN =====
        left = ctk.CTkFrame(main, width=380)
        left.pack(side="left", fill="y", padx=(0, 8), pady=4)
        right = ctk.CTkFrame(main)
        right.pack(side="left", fill="both", expand=True, pady=4)

        ctk.CTkButton(left, text="Load xlsx", command=self.load_xlsx).pack(fill="x", padx=8, pady=(10, 2))
        self.file_lbl = ctk.CTkLabel(left, text="No file")
        self.file_lbl.pack(anchor="w", padx=10)
        ctk.CTkLabel(left, text="Sheets:", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
        self.sheet_frame = ctk.CTkScrollableFrame(left, height=260)
        self.sheet_frame.pack(fill="x", padx=8, pady=4)
        row = ctk.CTkFrame(left, fg_color="transparent")
        row.pack(fill="x", padx=8)
        ctk.CTkButton(row, text="All", command=lambda: self._set_sheets(True)).pack(side="left", expand=True, fill="x", padx=(0, 4))
        ctk.CTkButton(row, text="None", command=lambda: self._set_sheets(False)).pack(side="left", expand=True, fill="x")

        self.dry_run = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(left, text="Dry-run (validation + CSV only)", variable=self.dry_run).pack(anchor="w", padx=10, pady=(10, 2))

        ctk.CTkButton(left, text="WWISE TREE PREVIEW", fg_color="#2471A3",
                      command=lambda: self.run_threaded(preview_only=True)).pack(fill="x", padx=8, pady=(10, 2))
        ctk.CTkButton(left, text="ORPHANS (Wwise vs Excel diff)", fg_color="#7D3C98",
                      hover_color="#5B2C6F",
                      command=self.orphans_threaded).pack(fill="x", padx=8, pady=(0, 2))
        self.go_btn = ctk.CTkButton(left, text="PLAN → GENERATE", height=52,
                                    font=("Arial", 16, "bold"), fg_color="#27AE60",
                                    hover_color="#1E8449", state="disabled",
                                    command=lambda: self.run_threaded(preview_only=False))
        self.go_btn.pack(fill="x", padx=8, pady=(4, 12))

        ctk.CTkLabel(right, text="Log / Plan / Tree:", font=("Arial", 12, "bold")).pack(anchor="w", padx=8, pady=(6, 0))
        self.log_box = ctk.CTkTextbox(right, font=("Consolas", 12))
        self.log_box.pack(fill="both", expand=True, padx=8, pady=8)

        # ===== SETTINGS =====
        s = ctk.CTkScrollableFrame(sett)
        s.pack(fill="both", expand=True, padx=4, pady=4)

        def lbl(text):
            ctk.CTkLabel(s, text=text, font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=(12, 2))

        lbl("ElevenLabs API Key")
        self.api_entry = ctk.CTkEntry(s, show="*")
        self.api_entry.insert(0, self.cfg["api_key"])
        self.api_entry.pack(fill="x", padx=10)

        lbl("Models / output format")
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
        ctk.CTkButton(s, text="Refresh models from API", fg_color="#8E44AD", command=self.fetch_models).pack(anchor="w", padx=10, pady=4)

        def path_row(label_text, initial, fetch=False):
            lbl(label_text)
            fr = ctk.CTkFrame(s, fg_color="transparent")
            fr.pack(fill="x", padx=10)
            ent = ctk.CTkEntry(fr)
            ent.insert(0, initial)
            ent.pack(side="left", expand=True, fill="x")
            if fetch:
                ctk.CTkButton(fr, text="Fetch from Wwise", width=130, fg_color="#E07A22",
                              hover_color="#B56017",
                              command=lambda e=ent: self.fetch_wwise_path(e)).pack(side="left", padx=(6, 0))
            return ent

        self.audio_root = path_row("Wwise: Audio Root (Actor-Mixer Hierarchy → VOICEOVER)", self.cfg["audio_root"], fetch=True)
        self.event_root = path_row("Wwise: Events Root (Events → VOICEOVER)", self.cfg["event_root"], fetch=True)
        self.ue_root = path_row("Unreal: events directory (WwiseAudio → Events)", self.cfg["ue_root"], fetch=False)

        lbl(f"Loudness-norm per Type (pyloudnorm: {'OK' if HAS_LOUDNORM else 'MISSING - py -m pip install pyloudnorm'})")
        self.do_loudnorm = ctk.BooleanVar(value=self.cfg["loudnorm"] and HAS_LOUDNORM)
        ctk.CTkCheckBox(s, text=f"Normalize raw wav (targets: {self.cfg['type_lufs']})",
                        variable=self.do_loudnorm,
                        state="normal" if HAS_LOUDNORM else "disabled").pack(anchor="w", padx=10)

        lbl("Phrases stripped from subtitles (regex, one per line)")
        self.strip_box = ctk.CTkTextbox(s, height=90, font=("Consolas", 12))
        self.strip_box.pack(fill="x", padx=10)
        self.strip_box.insert("0.0", "\n".join(self.cfg["subtitle_strip_regex"]))

        ctk.CTkLabel(s, text=f"Tree structure (object types, branches per Type): edit deliberately in the file\n{CONFIG_PATH}\n- defaults = report §7.2 / Appendix B. Settings are saved on generation and on close.",
                     justify="left", text_color="#999999").pack(anchor="w", padx=10, pady=14)
        ctk.CTkButton(s, text="Save settings now", command=lambda: (save_config(self._collect_cfg()), self.log("Settings saved."))).pack(anchor="w", padx=10, pady=(0, 14))

        self.log("VO Generator v2 ready. Load an xlsx (Main tab).")
        self.log(f"Config: {CONFIG_PATH}" + ("" if os.path.exists(CONFIG_PATH) else " (will be created on save)"))
        if not HAS_WAAPI:
            self.log("WARNING: waapi-client missing - only dry-run is possible.")

    def log(self, msg):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.update()

    def _set_sheets(self, val):
        for v in self.sheet_vars.values():
            v.set(val)

    # ---------- actions ----------
    def fetch_wwise_path(self, entry):
        if not HAS_WAAPI:
            self.log("⛔ waapi-client missing.")
            return
        try:
            with WaapiClient(allow_exception=True) as client:
                res = client.call("ak.wwise.ui.getSelectedObjects", {"options": {"return": ["path"]}})
                objs = res.get("objects") if res else None
                if objs:
                    path = objs[0]["path"]
                    entry.delete(0, "end")
                    entry.insert(0, path)
                    self.log(f"⤓ Fetched from Wwise: {path}")
                else:
                    self.log("⛔ Nothing selected in Wwise.")
        except Exception as e:
            self.log(f"⛔ No connection to Wwise ({e.__class__.__name__}).")

    def fetch_models(self):
        key = self.api_entry.get().strip()
        if not key:
            self.log("ERROR: enter the API key (Settings).")
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
            self.log(f"Models: {len(tts)} TTS, {len(s2s)} S2S.")
        except Exception as e:
            self.log(f"API error: {e}")

    def load_xlsx(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            casting, vocab, lines, errors = load_workbook_data(path)
        except Exception as e:
            # .pyw has no console - without this a failed load is completely silent
            self.log(f"⛔ Failed to load {os.path.basename(path)}: {e!r}")
            return
        self.xlsx_path = path
        self.data = (casting, vocab, lines)
        for w in self.sheet_frame.winfo_children():
            w.destroy()
        self.sheet_vars.clear()
        for name in sorted({l["sheet"] for l in lines}):
            var = ctk.BooleanVar(value=True)
            self.sheet_vars[name] = var
            ctk.CTkCheckBox(self.sheet_frame, text=name, variable=var).pack(anchor="w", pady=1)
        self.file_lbl.configure(text=os.path.basename(path))
        self.log(f"\nLoaded: {os.path.basename(path)} - Casting: {len(casting)}, "
                 f"Intents: {len(vocab['Intent'])}, lines: {len(lines)}.")
        for e in errors:
            self.log("⛔ " + e)
        self.go_btn.configure(state="normal" if not errors else "disabled")

    def orphans_threaded(self):
        if not self.data:
            self.log("Load an xlsx first.")
            return
        if not HAS_WAAPI:
            self.log("⛔ waapi-client missing - the diff requires a connection to Wwise.")
            return
        threading.Thread(target=self._orphans_safe, daemon=True).start()

    def _orphans_safe(self):
        try:
            self.find_and_clean_orphans()
        except Exception as e:
            self.log(f"⛔ ORPHANS - ERROR: {e!r}")

    def find_and_clean_orphans(self):
        """Wwise vs Excel diff (report §14): reports TOP-MOST orphans in the managed
        branches (Barks/Story/Conversations) + deletion after confirmation.
        Efforts and everything outside the managed branches - UNTOUCHED."""
        cfg = self._collect_cfg()
        casting, vocab, lines = self.data

        # Expected: ALL rows with a LineID, regardless of the sheet checkboxes.
        expected = expected_wwise_paths(lines, cfg, cfg["audio_root"], cfg["event_root"])
        roots, protected = managed_roots(cfg, cfg["audio_root"], cfg["event_root"])
        if not expected:
            # e.g. a template/stale workbook: the whole Wwise tree would count as orphans
            self.log("⛔ ORPHANS: no rows with a Line_ID in this workbook - refusing to diff.")
            return

        self.log("\n=== ORPHANS: Wwise scan ===")
        actual = []
        with WaapiClient(allow_exception=True) as client:
            for base, top, full in roots:
                try:
                    res = client.call("ak.wwise.core.object.get", {
                        "waql": f'$ "{full}" select descendants',
                        "options": {"return": ["id", "path", "type", "name"]},
                    })
                    objs = (res or {}).get("return", [])
                    actual.extend(objs)
                    self.log(f"  {full}: {len(objs)} objects")
                except Exception:
                    self.log(f"  {full}: does not exist (skipping)")

            orphans = find_orphans(actual, expected, protected)
            if not orphans:
                self.log("✅ Zero orphans - Wwise matches Excel 1:1 (in the managed branches).")
                return

            self.log(f"\n⚠️ Orphans (top-most, deleting a parent takes its children): {len(orphans)}")
            for o in orphans:
                self.log(f"  🧹 [{o.get('type','?')}] {o.get('path','?')}")
            self.log("\nNOTE: an object added MANUALLY in Barks/Story/Conversations will also end up on this list"
                     " - these branches are fully managed by the generator (Efforts are outside the diff).")

            if not messagebox.askyesno("Orphans in Wwise",
                    f"Found {len(orphans)} objects that are not in Excel\n"
                    f"(full list in the log).\n\nDELETE them from the Wwise project?"):
                self.log("Cancelled - nothing deleted (the report stays in the log).")
                return

            deleted = 0
            for o in orphans:
                try:
                    client.call("ak.wwise.core.object.delete", {"object": o["id"]})
                    deleted += 1
                except Exception as e:
                    self.log(f"  ⛔ failed to delete {o.get('path','?')}: {e!r}")
            self.log(f"✅ Deleted {deleted}/{len(orphans)} orphans.")

    def run_threaded(self, preview_only=False):
        if not self.data:
            self.log("Load an xlsx first.")
            return
        self.go_btn.configure(state="disabled", text="WORKING…")
        threading.Thread(target=self._run_safe, args=(preview_only,), daemon=True).start()

    def _run_safe(self, preview_only):
        try:
            self.run(preview_only)
        except Exception as e:
            self.log(f"⛔ CRITICAL ERROR: {e!r}")
        finally:
            self.go_btn.configure(state="normal", text="PLAN → GENERATE")

    # ---------- main run ----------
    def _validate_and_plan(self, cfg):
        casting, vocab, lines = self.data
        active = [l for l in lines if self.sheet_vars.get(l["sheet"]) and self.sheet_vars[l["sheet"]].get()]
        self.log("\n=== VALIDATION ===")
        seen_ids, seen_seq, valid, n_err = {}, {}, [], 0
        for row in active:
            errs = validate_line(row, casting, vocab, seen_ids, seen_seq)
            if row["LineID"]:
                seen_ids[row["LineID"]] = f"{row['sheet']} row {row['row']}"
            if errs:
                n_err += len(errs)
                self.log(f"⛔ [{row['sheet']} row {row['row']}] " + "; ".join(errs))
            else:
                valid.append(row)
        self.log(f"Rows OK: {len(valid)} / {len(active)}; errors: {n_err}.")
        if n_err:
            return None
        return build_plan(valid, casting, self.manifest, self.out_dir, cfg)

    def _csv_items(self, plan, cfg):
        """The CSVs are merged tables that replace the UE DataTables wholesale, so
        they must hold ALL valid lines - including the unchecked sheets. Writing
        only the checked ones would silently drop the rest on the next reimport."""
        casting, vocab, lines = self.data
        items = {id(p["row"]): p for p in plan}
        seen_ids = {p["row"]["LineID"]: f"{p['row']['sheet']} row {p['row']['row']}" for p in plan}
        seen_seq = {(p["d"]["ConversationID"], p["row"]["Seq"]): True
                    for p in plan if p["d"]["ConversationID"]}
        added = skipped = 0
        for row in lines:
            if id(row) in items:
                continue
            if validate_line(row, casting, vocab, seen_ids, seen_seq):
                skipped += 1
                continue
            seen_ids[row["LineID"]] = f"{row['sheet']} row {row['row']}"
            items[id(row)] = {"row": row, "d": derive(row, casting, cfg)}
            added += 1
        if added:
            self.log(f"  + {added} lines from unchecked sheets")
        if skipped:
            self.log(f"  ⚠️ {skipped} invalid lines in unchecked sheets left out of the CSVs")
        return [items[id(l)] for l in lines if id(l) in items]

    def run(self, preview_only=False):
        cfg = self._collect_cfg()
        save_config(cfg)
        casting, vocab, lines = self.data

        plan = self._validate_and_plan(cfg)
        if plan is None:
            self.log("Fix the sheet and try again. NOTHING was generated.")
            return

        if preview_only:
            self.log("\n=== WWISE TREE (preview - nothing is created) ===")
            self.log(build_tree_preview(plan, cfg, cfg["audio_root"], cfg["event_root"]))
            return

        counts = {}
        for p in plan:
            counts[p["action"]] = counts.get(p["action"], 0) + 1
        self.log("\n=== PLAN ===")
        self.log(f"🆕 new: {counts.get('NEW',0)}   🔁 regen: {counts.get('REGEN',0)}   "
                 f"💾 cache: {counts.get('CACHE',0)}   🔒 locked: {counts.get('LOCKED',0)}")
        for p in plan:
            if p["action"] == "REGEN":
                self.log(f"  🔁 {p['row']['LineID']} - {p['why']}")
            elif p["action"] == "LOCKED":
                self.log(f"  🔒 {p['row']['LineID']} - skipping (final)")
        still_flagged = [p["row"]["LineID"] for p in plan
                         if p["row"]["ForceRegen"] and p["action"] == "REGEN"]
        to_generate = [p for p in plan if p["action"] in ("NEW", "REGEN")]

        if self.dry_run.get():
            self.log("\nDRY-RUN: skipping audio and Wwise - writing CSV only.")
            paths = write_csvs(self._csv_items(plan, cfg), casting, self.out_dir, cfg["ue_root"], cfg["subtitle_strip_regex"])
            for v in paths.values():
                self.log(f"  📄 {v}")
            self.log("✅ Dry-run finished.")
            return

        if to_generate:
            if not cfg["api_key"]:
                self.log("⛔ No ElevenLabs API key (Settings) - nothing generated.")
                return
            if not messagebox.askyesno("Generation plan",
                    f"New: {counts.get('NEW',0)}\nRegenerations: {counts.get('REGEN',0)}\n"
                    f"Cache (skipping): {counts.get('CACHE',0)}\nLocked: {counts.get('LOCKED',0)}\n\n"
                    f"Run ElevenLabs for {len(to_generate)} lines?"):
                self.log("Cancelled - nothing generated.")
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
            self.log(f"⚠️ Failed generations: {len(failed)} - skipping them in the import.")

        importable = [p for p in plan if p["action"] in ("NEW", "REGEN", "CACHE")
                      and os.path.exists(p["path"])]
        self.log("\n=== WWISE (WAAPI) ===")
        if not HAS_WAAPI:
            self.log("⛔ waapi-client missing - skipping import.")
        else:
            try:
                self._wwise_import(importable, cfg)
            except Exception as e:
                self.log(f"⛔ WAAPI: {e!r}")
                return

        self.log("\n=== CSV ===")
        paths = write_csvs(self._csv_items(plan, cfg), casting, self.out_dir, cfg["ue_root"], cfg["subtitle_strip_regex"])
        for v in paths.values():
            self.log(f"  📄 {v}")

        if still_flagged:
            self.log("\n⚠️ REMINDER: uncheck Force_Regen in the sheet for:")
            for lid in still_flagged:
                self.log("   • " + lid)
        self.log("\n🚀 DONE.")

    # ---------- audio ----------
    def _make_audio(self, row, d, out_path, cfg):
        key, fmt = cfg["api_key"], cfg["format"]
        try:
            if not d["voice"]:
                self.log(f"⛔ {row['LineID']}: Speaker '{row['Speaker']}' has no Voice_ID in Casting")
                return False
            if row["S2S"]:
                ref_dir = os.path.join(self.ref_dir, row["S2S"])
                refs = [f for f in os.listdir(ref_dir)
                        if f.lower().endswith((".wav", ".mp3"))] if os.path.isdir(ref_dir) else []
                if not refs:
                    self.log(f"⛔ {row['LineID']}: no S2S reference in {row['S2S']}")
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
        with WaapiClient(allow_exception=True) as client:
            def ensure(parent, otype, name, props=None):
                args = {"parent": parent, "type": otype, "name": name,
                        "onNameConflict": "merge"}
                if props:
                    args.update(props)
                client.call("ak.wwise.core.object.create", args)
                return parent + "\\" + name

            for label, root in (("Audio Root", audio_root), ("Events Root", event_root)):
                try:
                    client.call("ak.wwise.core.object.get", {"waql": f'$ "{root}"'})
                except Exception:
                    raise RuntimeError(f"{label} not found in Wwise: {root} - create it or fix Settings")

            for p in plan_items:
                row, d = p["row"], p["d"]
                sp, ty, ctx, lid = row["Speaker"], row["Type"], row["Context"], row["LineID"]

                cur = audio_root
                for seg in d["tree"]:
                    cur = ensure(cur, cfg["tree_object"], seg)
                cur = ensure(cur, cfg["speaker_object"], sp)
                if ty in CONTAINER_TYPES:
                    # container = event target -> key-based name vo_{speaker}_{type}_{intent}
                    # (structure = simple names; playable objects = file key)
                    cur = ensure(cur, cfg["container_object"], container_name(row), {"@RandomOrSequence": 1})
                else:
                    cur = ensure(cur, cfg["context_object"], ctx)

                client.call("ak.wwise.core.audio.import", {
                    "importOperation": "useExisting",
                    "default": {"importLanguage": "English(US)"},
                    # The source must be named explicitly: without it useExisting adds
                    # a duplicate AudioFileSource to the Sound on every re-import.
                    "imports": [{
                        "objectPath": f"{cur}\\<Sound Voice>{lid}\\<AudioFileSource>{os.path.splitext(d['file'])[0]}",
                        "audioFile": p["path"],
                    }],
                })
                # import-level notes would land on the source (last path element), not the Sound
                client.call("ak.wwise.core.object.setNotes", {
                    "object": f"{cur}\\{lid}",
                    "value": strip_tags(row["Text"], cfg["subtitle_strip_regex"]),
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
        self.log(f"Import: {len(plan_items)} files, {len(created_events)} events.")


if __name__ == "__main__":
    VOGen2().mainloop()
