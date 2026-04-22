import os
import re
import random
import requests
import pandas as pd
import threading
import tkinter as tk
import wave
import customtkinter as ctk
from customtkinter import filedialog
from waapi import WaapiClient, CannotConnectToWaapiException

# Konfiguracja wyglądu GUI
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class VOGeneratorApp(ctk.CTk):
    def start_generation_thread(self):
        self.generate_btn.configure(state="disabled", text="⏳ GENEROWANIE W TOKU...")
        thread = threading.Thread(target=self.run_generation, daemon=True)
        thread.start()

    def __init__(self):
        super().__init__()
        self.title("ElevenLabs to Wwise & UE5 Pipeline (Master Edition S2S)")
        self.geometry("1300x880")
        self.data_dict = {}
        self.sheet_vars = {}

        # --- STRUKTURA FOLDERÓW ---
        self.output_dir = os.path.abspath("Output_Audio")
        self.ref_dir = os.path.abspath("S2S_Reference")

        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
        if not os.path.exists(self.ref_dir):
            os.makedirs(self.ref_dir)

        # --- GŁÓWNY PODZIAŁ NA KOLUMNY ---
        self.paned_window = tk.PanedWindow(
            self, orient=tk.HORIZONTAL, sashwidth=6, bg="#2b2b2b", bd=0
        )
        self.paned_window.pack(fill="both", expand=True, padx=10, pady=10)

        self.left_wrapper = ctk.CTkFrame(self.paned_window, fg_color="transparent")
        self.paned_window.add(self.left_wrapper, minsize=450)

        self.left_panel = ctk.CTkScrollableFrame(self.left_wrapper, width=600)
        self.left_panel.pack(fill="both", expand=True, padx=(0, 5))

        self.right_wrapper = ctk.CTkFrame(self.paned_window, fg_color="transparent")
        self.paned_window.add(self.right_wrapper, minsize=400)

        self.right_panel = ctk.CTkFrame(self.right_wrapper)
        self.right_panel.pack(fill="both", expand=True, padx=(5, 0))

        # --- PRAWY PANEL ---
        self.generate_btn = ctk.CTkButton(
            self.right_panel,
            text="6. GENERUJ AUDIO I WYŚLIJ",
            height=60,
            font=("Arial", 18, "bold"),
            fg_color="#27AE60",
            hover_color="#1E8449",
            command=self.start_generation_thread,
        )
        self.generate_btn.pack(pady=(10, 15), fill="x", padx=10)
        self.generate_btn.configure(state="disabled")

        self.log_label = ctk.CTkLabel(
            self.right_panel, text="Logi Systemowe:", font=("Arial", 14, "bold")
        )
        self.log_label.pack(anchor="w", padx=10, pady=(0, 0))
        self.log_box = ctk.CTkTextbox(self.right_panel, font=("Consolas", 12))
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        self.log_box.insert(
            "0.0",
            "System gotowy.\nUżywaj kolumny 'S2S_Reference' w Excelu dla Voice Changera.\n",
        )

        # --- LEWY PANEL ---
        self.frame_api = ctk.CTkFrame(self.left_panel)
        self.frame_api.pack(pady=5, padx=10, fill="x")
        self.api_label = ctk.CTkLabel(
            self.frame_api, text="1. ElevenLabs API Key:", font=("Arial", 12, "bold")
        )
        self.api_label.pack(anchor="w", padx=10, pady=(5, 0))
        self.api_entry = ctk.CTkEntry(self.frame_api, show="*")
        self.api_entry.pack(anchor="w", padx=10, pady=(0, 10), fill="x")

        self.frame_el_settings = ctk.CTkFrame(self.left_panel)
        self.frame_el_settings.pack(pady=5, padx=10, fill="x")
        self.el_label = ctk.CTkLabel(
            self.frame_el_settings,
            text="2. ElevenLabs Settings:",
            font=("Arial", 12, "bold"),
        )
        self.el_label.grid(row=0, column=0, padx=10, pady=(5, 5), sticky="w")

        self.fetch_models_btn = ctk.CTkButton(
            self.frame_el_settings,
            text="Odśwież z API",
            width=120,
            fg_color="#8E44AD",
            hover_color="#732D91",
            command=self.fetch_api_models,
        )
        self.fetch_models_btn.grid(row=0, column=1, padx=10, pady=(5, 5), sticky="w")

        self.tts_model_label = ctk.CTkLabel(self.frame_el_settings, text="TTS Model:")
        self.tts_model_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.tts_model_var = ctk.StringVar(value="eleven_v3")
        self.tts_model_menu = ctk.CTkOptionMenu(
            self.frame_el_settings, variable=self.tts_model_var, values=["eleven_v3"]
        )
        self.tts_model_menu.grid(row=1, column=1, padx=10, pady=5, sticky="we")

        self.s2s_model_label = ctk.CTkLabel(self.frame_el_settings, text="S2S Model:")
        self.s2s_model_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.s2s_model_var = ctk.StringVar(value="eleven_multilingual_sts_v2")
        self.s2s_model_menu = ctk.CTkOptionMenu(
            self.frame_el_settings,
            variable=self.s2s_model_var,
            values=["eleven_multilingual_sts_v2"],
        )
        self.s2s_model_menu.grid(row=2, column=1, padx=10, pady=5, sticky="we")

        self.quality_label = ctk.CTkLabel(self.frame_el_settings, text="Output Format:")
        self.quality_label.grid(row=3, column=0, padx=10, pady=(5, 10), sticky="w")
        self.quality_var = ctk.StringVar(value="pcm_24000")
        self.quality_menu = ctk.CTkOptionMenu(
            self.frame_el_settings,
            variable=self.quality_var,
            values=["pcm_24000", "pcm_44100", "mp3_44100_128"],
        )
        self.quality_menu.grid(row=3, column=1, padx=10, pady=(5, 10), sticky="we")

        self.frame_targets = ctk.CTkFrame(self.left_panel)
        self.frame_targets.pack(pady=5, padx=10, fill="x")
        self.targets_label = ctk.CTkLabel(
            self.frame_targets,
            text="3. Target Paths (Wwise & UE5):",
            font=("Arial", 12, "bold"),
        )
        self.targets_label.grid(
            row=0, column=0, columnspan=3, padx=10, pady=(5, 5), sticky="w"
        )

        self.audio_root_entry = self.add_path_row(
            self.frame_targets,
            "Audio Root:",
            "\\Containers\\Default Work Unit\\NPC\\NPC_VO",
            1,
        )
        self.event_root_entry = self.add_path_row(
            self.frame_targets,
            "Event Root:",
            "\\Events\\Default Work Unit\\NPC\\NPC_VO",
            2,
        )

        self.ue_root_label = ctk.CTkLabel(self.frame_targets, text="UE5 Soft-Ref:")
        self.ue_root_label.grid(row=3, column=0, padx=10, pady=(0, 10), sticky="w")
        self.ue_root_entry = ctk.CTkEntry(self.frame_targets, width=330)
        self.ue_root_entry.insert(0, "/Game/WwiseAudio/Events")
        self.ue_root_entry.grid(row=3, column=1, padx=5, pady=(0, 10), sticky="we")

        self.frame_recipes = ctk.CTkFrame(self.left_panel)
        self.frame_recipes.pack(pady=5, padx=10, fill="x")
        self.recipes_label = ctk.CTkLabel(
            self.frame_recipes,
            text="4. Hierarchia w Wwise:",
            font=("Arial", 12, "bold"),
        )
        self.recipes_label.pack(anchor="w", padx=10, pady=(5, 5))

        self.naming_entry = self.add_recipe_row(
            self.frame_recipes,
            "Audio File Name:",
            "vo_npc_{category}_{character}_{intent}",
        )
        self.audio_hier_entry = self.add_recipe_row(
            self.frame_recipes,
            "Audio Hierarchy:",
            "<Folder>{Category}\\<ActorMixer>{Character}\\<Folder>{intent}",
        )
        self.event_name_entry = self.add_recipe_row(
            self.frame_recipes, "Event Name:", "Play_{line_id}"
        )
        self.event_hier_entry = self.add_recipe_row(
            self.frame_recipes,
            "Event Hierarchy:",
            "<Folder>{Category}\\<Folder>{Character}\\<Folder>{intent}",
        )

        self.frame_mid = ctk.CTkFrame(self.left_panel)
        self.frame_mid.pack(pady=5, padx=10, fill="x")
        self.file_btn = ctk.CTkButton(
            self.frame_mid, text="5. Wczytaj Excel/CSV", command=self.load_file
        )
        self.file_btn.pack(side="top", fill="x", padx=10, pady=(10, 5))
        self.file_label = ctk.CTkLabel(self.frame_mid, text="Brak pliku")
        self.file_label.pack(side="top", padx=10, pady=(0, 5))

        self.sheet_scroll = ctk.CTkScrollableFrame(self.frame_mid, height=120)
        self.sheet_scroll.pack(fill="x", padx=10, pady=5)

        self.frame_sheet_btns = ctk.CTkFrame(self.frame_mid, fg_color="transparent")
        self.frame_sheet_btns.pack(fill="x", padx=10, pady=(0, 10))
        self.btn_select_all = ctk.CTkButton(
            self.frame_sheet_btns,
            text="Zaznacz Wszystkie",
            command=self.select_all_sheets,
        )
        self.btn_select_all.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.btn_deselect_all = ctk.CTkButton(
            self.frame_sheet_btns,
            text="Odznacz Wszystkie",
            command=self.deselect_all_sheets,
        )
        self.btn_deselect_all.pack(side="left", fill="x", expand=True, padx=(5, 0))

    def add_path_row(self, parent, label_text, default_val, row_idx):
        lbl = ctk.CTkLabel(parent, text=label_text)
        lbl.grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        ent = ctk.CTkEntry(parent, width=330)
        ent.insert(0, default_val)
        ent.grid(row=row_idx, column=1, padx=5, pady=5, sticky="we")
        btn = ctk.CTkButton(
            parent,
            text="Pobierz",
            width=70,
            fg_color="#E07A22",
            hover_color="#B56017",
            command=lambda: self.fetch_path_from_wwise(ent),
        )
        btn.grid(row=row_idx, column=2, padx=5, pady=5)
        return ent

    def add_recipe_row(self, parent, label_text, default_val):
        lbl = ctk.CTkLabel(parent, text=label_text)
        lbl.pack(anchor="w", padx=10)
        ent = ctk.CTkEntry(parent)
        ent.insert(0, default_val)
        ent.pack(anchor="w", padx=10, pady=(0, 5), fill="x")
        return ent

    def log(self, text):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.update()

    def fetch_api_models(self):
        api_key = self.api_entry.get().strip()
        if not api_key:
            self.log("BŁĄD: Wprowadź klucz API!")
            return
        url = "https://api.elevenlabs.io/v1/models"
        headers = {"xi-api-key": api_key}
        try:
            r = requests.get(url, headers=headers)
            if r.status_code == 200:
                models = r.json()
                tts = [m["model_id"] for m in models if m.get("can_do_text_to_speech")]
                s2s = [
                    m["model_id"] for m in models if m.get("can_do_voice_conversion")
                ]
                self.tts_model_menu.configure(values=tts)
                self.s2s_model_menu.configure(values=s2s)
                if tts:
                    self.tts_model_var.set(next((m for m in tts if "v3" in m), tts[0]))
                if s2s:
                    self.s2s_model_var.set(next((m for m in s2s if "sts" in m), s2s[0]))
                self.log(f"✅ Pobrano modele: {len(tts)} TTS, {len(s2s)} S2S.")
        except Exception as e:
            self.log(f"Błąd API: {e}")

    def fetch_path_from_wwise(self, target_entry):
        try:
            with WaapiClient() as client:
                res = client.call(
                    "ak.wwise.ui.getSelectedObjects", {"options": {"return": ["path"]}}
                )
                if res.get("objects"):
                    path = res["objects"][0]["path"]
                    target_entry.delete(0, "end")
                    target_entry.insert(0, path)
                    self.log(f"Pobrano: {path}")
        except Exception:
            self.log("BŁĄD: Brak połączenia z Wwise!")

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.csv")])
        if not path:
            return
        self.data_dict = (
            pd.read_excel(path, sheet_name=None)
            if path.endswith(".xlsx")
            else {os.path.basename(path).split(".")[0]: pd.read_csv(path)}
        )
        for w in self.sheet_scroll.winfo_children():
            w.destroy()
        self.sheet_vars.clear()
        for name in self.data_dict:
            var = ctk.BooleanVar(value=False)
            self.sheet_vars[name] = var
            ctk.CTkCheckBox(
                self.sheet_scroll,
                text=name,
                variable=var,
                command=self.check_generate_state,
            ).pack(anchor="w", pady=2)
        self.file_label.configure(text=os.path.basename(path))
        self.check_generate_state()

    def check_generate_state(self):
        state = (
            "normal" if any(v.get() for v in self.sheet_vars.values()) else "disabled"
        )
        self.generate_btn.configure(state=state)

    def select_all_sheets(self):
        for v in self.sheet_vars.values():
            v.set(True)
        self.check_generate_state()

    def deselect_all_sheets(self):
        for v in self.sheet_vars.values():
            v.set(False)
        self.check_generate_state()

    def save_audio_file(self, content, format_str, output_path):
        if format_str.startswith("pcm_"):
            sample_rate = int(format_str.split("_")[1])
            with wave.open(output_path, "wb") as wav_file:
                wav_file.setnchannels(1)
                wav_file.setsampwidth(2)
                wav_file.setframerate(sample_rate)
                wav_file.writeframes(content)
        else:
            with open(output_path, "wb") as f:
                f.write(content)

    def generate_tts(self, text, voice_id, output):
        api_key = self.api_entry.get().strip()
        fmt = self.quality_var.get()
        url = f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}?output_format={fmt}"
        r = requests.post(
            url,
            json={"text": text, "model_id": self.tts_model_var.get()},
            headers={"xi-api-key": api_key},
        )
        if r.status_code == 200:
            self.save_audio_file(r.content, fmt, output)
            return True
        self.log(f"Błąd TTS: {r.text}")
        return False

    def generate_s2s(self, ref_path, voice_id, output):
        api_key = self.api_entry.get().strip()
        fmt = self.quality_var.get()
        url = f"https://api.elevenlabs.io/v1/speech-to-speech/{voice_id}?output_format={fmt}"
        ext = ref_path.lower().split(".")[-1]
        mime = "audio/mpeg" if ext == "mp3" else "audio/wav"
        with open(ref_path, "rb") as f_audio:
            r = requests.post(
                url,
                headers={"xi-api-key": api_key},
                data={"model_id": self.s2s_model_var.get()},
                files={"audio": (f"ref.{ext}", f_audio, mime)},
            )
            if r.status_code == 200:
                self.save_audio_file(r.content, fmt, output)
                return True
        self.log(f"Błąd S2S: {r.text}")
        return False

    def parse_recipe(self, recipe, fmt_dict):
        try:
            s = recipe.format(**fmt_dict)
        except:
            s = recipe
        nodes = []
        for p in s.split("\\"):
            if not p.strip():
                continue
            m = re.match(r"<([^>]+)>(.*)", p)
            nodes.append(
                {"type": m.group(1) if m else "Folder", "name": m.group(2) if m else p}
            )
        return nodes

    def run_generation(self):
        wwise_audio_root = self.audio_root_entry.get().strip().rstrip("\\")
        wwise_event_root = self.event_root_entry.get().strip().rstrip("\\")
        ue_root = self.ue_root_entry.get().strip().rstrip("/")

        for name, df in self.data_dict.items():
            if not self.sheet_vars.get(name) or not self.sheet_vars[name].get():
                continue
            self.log(f"\n--- Przetwarzanie: {name} ---")
            df = df.fillna("")
            ue_data = []

            try:
                with WaapiClient() as client:
                    for _, row in df.iterrows():
                        cat = str(row.get("Category", "VO")).strip()
                        char = str(row.get("Character_Name", name)).strip()
                        intent = str(row.get("Intent_Tag", "default")).strip()
                        line_id_suffix = str(row.get("Line_ID", "")).strip()
                        text = str(row.get("Spoken_Text", "")).strip()
                        v_id = str(row.get("Voice_ID", "")).strip()
                        regen = str(row.get("Force_Regen", "0")).strip().upper() in [
                            "1",
                            "TRUE",
                            "YES",
                            "T",
                        ]

                        s2s_ref = str(row.get("S2S_Reference", "")).strip()
                        seq = (
                            int(row.get("Sequence_Order", 0))
                            if str(row.get("Sequence_Order")).isdigit()
                            else 0
                        )

                        if not text and not s2s_ref:
                            continue
                        if not line_id_suffix:
                            continue

                        clean_text = re.sub(r"\[.*?\]", "", text).strip()
                        fmt = {
                            "category": cat,
                            "character": char,
                            "intent": intent.replace(".", "_"),
                            "Category": cat,
                            "Character": char,
                            "Intent": intent,
                        }

                        file_name = (
                            self.naming_entry.get()
                            .format(**fmt)
                            .replace("__", "_")
                            .rstrip("_")
                            + f"_{line_id_suffix}"
                        )
                        file_name = file_name.lower()

                        fmt_str = self.quality_var.get()
                        ext = ".mp3" if "mp3" in fmt_str else ".wav"
                        out_path = os.path.join(self.output_dir, f"{file_name}{ext}")

                        # 1. Generowanie ElevenLabs (S2S / TTS)
                        if not os.path.exists(out_path) or regen:
                            if s2s_ref:
                                ref_path = os.path.join(self.ref_dir, s2s_ref)
                                if not os.path.exists(ref_path):
                                    self.log(f"  !!! Błąd S2S: Brak folderu {s2s_ref}")
                                    continue
                                valid = [
                                    f
                                    for f in os.listdir(ref_path)
                                    if f.lower().endswith((".wav", ".mp3"))
                                ]
                                if not valid:
                                    self.log(f"  !!! Błąd S2S: Pusty folder {s2s_ref}")
                                    continue
                                self.log(f"[{file_name}] S2S z referencji: {s2s_ref}")
                                if not self.generate_s2s(
                                    os.path.join(ref_path, random.choice(valid)),
                                    v_id,
                                    out_path,
                                ):
                                    continue
                            else:
                                self.log(f"[{file_name}] TTS: {text[:20]}...")
                                if not self.generate_tts(text, v_id, out_path):
                                    continue
                        else:
                            self.log(f"[{file_name}] Z dysku (Cache).")

                        # 2. Wwise Audio Hierarchy
                        curr_audio = wwise_audio_root
                        audio_nodes = self.parse_recipe(
                            self.audio_hier_entry.get(), fmt
                        )
                        for node in audio_nodes:
                            client.call(
                                "ak.wwise.core.object.create",
                                {
                                    "parent": curr_audio,
                                    "type": node["type"],
                                    "name": node["name"],
                                    "onNameConflict": "merge",
                                },
                            )
                            curr_audio += f"\\{node['name']}"

                        client.call(
                            "ak.wwise.core.audio.import",
                            {
                                "importOperation": "replaceExisting",
                                "default": {"importLanguage": "English(US)"},
                                "imports": [
                                    {
                                        "objectPath": f"{curr_audio}\\<Sound Voice>{file_name}",
                                        "audioFile": out_path,
                                        "@Notes": clean_text,
                                    }
                                ],
                            },
                        )

                        # 3. Wwise Event Hierarchy
                        curr_event = wwise_event_root
                        ue_folder = ""
                        ev_nodes = self.parse_recipe(self.event_hier_entry.get(), fmt)

                        for node in ev_nodes:
                            client.call(
                                "ak.wwise.core.object.create",
                                {
                                    "parent": curr_event,
                                    "type": node["type"],
                                    "name": node["name"],
                                    "onNameConflict": "merge",
                                },
                            )
                            curr_event += f"\\{node['name']}"
                            ue_folder += f"/{node['name']}"

                        ev_name = self.event_name_entry.get().format(
                            **fmt, line_id=file_name
                        )

                        # --- POPRAWKA: Dodano "name": "" wewnątrz children ---
                        e_res = client.call(
                            "ak.wwise.core.object.create",
                            {
                                "parent": curr_event,
                                "type": "Event",
                                "name": ev_name,
                                "onNameConflict": "merge",
                                "children": [
                                    {
                                        "name": "",  # <--- WWISE WYMAGA TEGO PARAMETRU DLA AKCJI
                                        "type": "Action",
                                        "@ActionType": 1,
                                        "@Target": f"{curr_audio}\\{file_name}",
                                    }
                                ],
                            },
                            {"return": ["shortId"]},
                        )

                        event_short_id = 0
                        if e_res and e_res.get("objects"):
                            event_short_id = e_res["objects"][0]["shortId"]
                            self.log(f"  -> Wwise: {ev_name} (ID: {event_short_id})")

                        # 4. UE5 Data
                        ue_ref = (
                            f"AkAudioEvent'{ue_root}{ue_folder}/{ev_name}.{ev_name}'"
                        )
                        ue_data.append(
                            {
                                "RowName": line_id_suffix,
                                "Intent_Tag": intent,
                                "Sequence_Order": seq,
                                "Wwise_Event_Name": ev_name,
                                "Wwise_Event_ID": event_short_id,
                                "Wwise_Event_Ref": ue_ref,
                                "Spoken_Text": clean_text,
                            }
                        )

            except Exception as e:
                self.log(f"Błąd WAAPI: {e}")

            if ue_data:
                csv_path = f"DT_NPC_Dialogues_{name}.csv"
                pd.DataFrame(ue_data).to_csv(
                    csv_path, index=False, encoding="utf-8-sig"
                )
                self.log(f"✅ Zapisano: {csv_path}")

        self.log("\n🚀 GOTOWE!")
        self.generate_btn.configure(state="normal", text="6. GENERUJ AUDIO I WYŚLIJ")


if __name__ == "__main__":
    VOGeneratorApp().mainloop()
