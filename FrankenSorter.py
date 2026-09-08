import os
import shutil
import time
import json
import re
import ollama
import pdfplumber
import openpyxl
from docx import Document
from pptx import Presentation
import customtkinter as ctk
from threading import Thread

import theme

theme.init_fonts()
ctk.set_appearance_mode("Dark")

# the local model doing the routing. any instruct model works — swap the name
# and `ollama pull` it. qwen2.5:7b is the sweet spot on 16gb: smart enough to
# read a doc and actually get the category right, fast enough to not crawl.
OLLAMA_MODEL = "qwen2.5:7b"

# the model is forced to answer in this exact shape. category is an enum, so it
# literally can't invent a folder that doesn't exist — no more cleaning up
# garbage categories after the fact.
ROUTING_SCHEMA = {
    "type": "object",
    "properties": {
        "category": {"type": "string", "enum": ["SCHOOL", "WORK", "PERSONAL", "FINANCE", "UNSORTED"]},
        "subject": {"type": "string"},
        "new_name": {"type": "string"},
    },
    "required": ["category", "subject", "new_name"],
}

VALID_CATEGORIES = {"SCHOOL", "WORK", "PERSONAL", "FINANCE"}


class FrankenSorterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("FrankenSorter")
        self.geometry("760x820")
        self.minsize(700, 760)

        # keep state on the instance so the buttons can lock themselves while a
        # sort is mid-flight — otherwise people spam-click and kick off two runs.
        self.source_folder = ""

        # default to the desktop. expanduser sorts out ~ on both mac and windows
        # so i don't have to hardcode a path per os.
        self.dest_folder = os.path.expanduser("~/Desktop/Franken_Sorted")

        self.stop_requested = False
        self.is_running = False

        self.mode = "dark"
        self.status_kind = "neg"

        self.build_fonts()
        self.build_ui()
        self.apply_theme()

        # ping the model on a daemon thread so the window still opens instantly
        # even if ollama is slow to answer. daemon means it dies with the app.
        Thread(target=self.check_ai_connection, daemon=True).start()

    @property
    def palette(self):
        return theme.PALETTES[self.mode]

    def build_fonts(self):
        """bricolage reads, martian counts — see theme.py for the full mapping."""
        self.font_wordmark = ctk.CTkFont(family=theme.FONT_MONO, size=20, weight="bold")
        self.font_tagline = ctk.CTkFont(family=theme.FONT_BODY, size=14)
        self.font_eyebrow = ctk.CTkFont(family=theme.FONT_MONO, size=11, weight="bold")
        self.font_body = ctk.CTkFont(family=theme.FONT_BODY, size=13)
        self.font_button = ctk.CTkFont(family=theme.FONT_SUBHEADING, size=13)
        self.font_mono_num = ctk.CTkFont(family=theme.FONT_MONO, size=13)
        self.font_log = ctk.CTkFont(family=theme.FONT_BODY, size=12)

    # --- ui construction ---------------------------------------------------

    def build_ui(self):
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=30, pady=24)

        self.build_header(self.main_frame)
        self.build_panel_row(self.main_frame)
        self.build_progress_row(self.main_frame)
        self.build_activity_panel(self.main_frame)
        self.build_footer(self.main_frame)

    def build_header(self, parent):
        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", pady=(0, 24))

        left = ctk.CTkFrame(header, fg_color="transparent")
        left.pack(side="left", anchor="n")

        self.wordmark_label = ctk.CTkLabel(left, text="FrankenSorter", font=self.font_wordmark, anchor="w")
        self.wordmark_label.pack(anchor="w")

        self.tagline_canvas = ctk.CTkCanvas(left, height=26, width=220, highlightthickness=0, bd=0)
        self.tagline_canvas.pack(anchor="w", pady=(4, 0))
        self.draw_tagline()

        right = ctk.CTkFrame(header, fg_color="transparent")
        right.pack(side="right", anchor="n")

        self.appearance_label = ctk.CTkLabel(right, text=theme.spaced("APPEARANCE"), font=self.font_eyebrow)
        self.appearance_label.pack(anchor="e")

        toggle_row = ctk.CTkFrame(right, fg_color="transparent")
        toggle_row.pack(anchor="e", pady=(6, 0))

        self.mode_label = ctk.CTkLabel(toggle_row, text="DARK", font=self.font_mono_num, width=40, anchor="e")
        self.mode_label.pack(side="left", padx=(0, 8))

        self.appearance_switch = ctk.CTkSwitch(
            toggle_row, text="", width=44, switch_width=44, switch_height=22,
            command=self.on_toggle_theme,
        )
        self.appearance_switch.pack(side="left")
        self.appearance_switch.select()  # start in dark mode

    def draw_tagline(self):
        """the one brand flourish: a chartreuse band swiped under a single word.
        css does this with a background gradient; tk can't, so i just draw a
        rectangle behind the text on a canvas — same look, done once, on the hero."""
        c = self.tagline_canvas
        c.delete("all")
        palette = self.palette
        c.configure(bg=palette["bg"])
        f = self.font_tagline
        lead = "Local AI, "
        word = "sorted."
        c.create_text(0, 13, text=lead, anchor="w", fill=palette["ink_2"], font=f, tags="lead")
        lead_width = f.measure(lead)
        word_width = f.measure(word)
        c.configure(width=lead_width + word_width + 4)
        band_y = 20
        c.create_rectangle(
            lead_width, band_y, lead_width + word_width, band_y + 4,
            fill=palette["brand"], outline="",
        )
        c.create_text(lead_width, 13, text=word, anchor="w", fill=palette["ink"], font=f)

    def build_panel_row(self, parent):
        self.src_frame = ctk.CTkFrame(parent, corner_radius=0, border_width=1)
        self.src_frame.pack(fill="x", pady=(0, 10))
        self.src_eyebrow, self.src_value, self.btn_src = self.build_folder_panel(
            self.src_frame, theme.spaced("SOURCE"), "No folder selected", "Choose…", self.select_source
        )

        self.dest_frame = ctk.CTkFrame(parent, corner_radius=0, border_width=1)
        self.dest_frame.pack(fill="x", pady=(0, 20))
        self.dest_eyebrow, self.dest_value, self.btn_dest = self.build_folder_panel(
            self.dest_frame, theme.spaced("DESTINATION"), self.dest_folder, "Change…", self.select_dest
        )

    def build_folder_panel(self, frame, eyebrow_text, value_text, button_text, command):
        text_col = ctk.CTkFrame(frame, fg_color="transparent")
        text_col.pack(side="left", fill="x", expand=True, padx=20, pady=14)

        eyebrow = ctk.CTkLabel(text_col, text=eyebrow_text, font=self.font_eyebrow, anchor="w")
        eyebrow.pack(anchor="w")

        value = ctk.CTkLabel(text_col, text=value_text, font=self.font_body, anchor="w")
        value.pack(anchor="w", pady=(4, 0))

        button = ctk.CTkButton(
            frame, text=button_text, width=110, corner_radius=0, border_width=1,
            font=self.font_button, command=command,
        )
        button.pack(side="right", padx=20, pady=14)
        return eyebrow, value, button

    def build_progress_row(self, parent):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=(0, 5))

        self.progress = ctk.CTkProgressBar(row, height=4, corner_radius=0)
        self.progress.pack(side="left", fill="x", expand=True)
        self.progress.set(0)

        self.progress_pct = ctk.CTkLabel(row, text="0%", font=self.font_mono_num, width=44, anchor="e")
        self.progress_pct.pack(side="right", padx=(10, 0))

    def build_activity_panel(self, parent):
        self.activity_eyebrow = ctk.CTkLabel(parent, text=theme.spaced("ACTIVITY"), font=self.font_eyebrow, anchor="w")
        self.activity_eyebrow.pack(fill="x", pady=(15, 6))

        self.log_frame = ctk.CTkFrame(parent, corner_radius=0, border_width=1)
        self.log_frame.pack(fill="both", expand=True)

        self.log_display = ctk.CTkTextbox(
            self.log_frame, corner_radius=0, border_width=0, font=self.font_log, wrap="word",
        )
        self.log_display.pack(fill="both", expand=True, padx=1, pady=1)
        self.log_display.configure(state="disabled")

    def build_footer(self, parent):
        footer = ctk.CTkFrame(parent, fg_color="transparent")
        footer.pack(fill="x", pady=(18, 0))

        self.status_label = ctk.CTkLabel(footer, text="Checking local AI…", font=self.font_body, anchor="w")
        self.status_label.pack(side="left")

        self.btn_run = ctk.CTkButton(
            footer, text="Sort Files", width=140, corner_radius=0,
            command=self.start_process, state="disabled", font=self.font_button,
        )
        self.btn_run.pack(side="right")

        self.btn_stop = ctk.CTkButton(
            footer, text="Stop", width=110, corner_radius=0, border_width=1,
            command=self.request_stop, state="disabled", font=self.font_button,
        )
        self.btn_stop.pack(side="right", padx=(0, 10))

    # --- theming -------------------------------------------------------

    def on_toggle_theme(self):
        self.mode = "dark" if self.appearance_switch.get() else "light"
        self.mode_label.configure(text="DARK" if self.mode == "dark" else "LIGHT")
        self.apply_theme()

    def apply_theme(self):
        p = self.palette
        ctk.set_appearance_mode("Dark" if self.mode == "dark" else "Light")

        self.configure(fg_color=p["bg"])
        self.wordmark_label.configure(text_color=p["ink"])
        self.draw_tagline()

        self.appearance_label.configure(text_color=p["ink_3"])
        self.mode_label.configure(text_color=p["ink_3"])
        self.appearance_switch.configure(
            fg_color=p["surface_2"], progress_color=p["brand"],
            button_color=p["bg"], button_hover_color=p["surface"],
        )

        for frame in (self.src_frame, self.dest_frame, self.log_frame):
            frame.configure(fg_color=p["surface"], border_color=p["rule"])

        for eyebrow in (self.src_eyebrow, self.dest_eyebrow, self.activity_eyebrow):
            eyebrow.configure(text_color=p["ink_3"])

        for value in (self.src_value, self.dest_value):
            value.configure(text_color=p["ink"])

        brand_hover = "#C8D62F"
        self.btn_run.configure(fg_color=p["brand"], text_color=p["brand_fg"], hover_color=brand_hover)

        for outline_btn in (self.btn_src, self.btn_dest, self.btn_stop):
            outline_btn.configure(
                fg_color="transparent", border_color=p["rule_2"], text_color=p["ink_2"],
                hover_color=p["surface_2"],
            )

        self.progress.configure(fg_color=p["surface_2"], progress_color=p["brand"])
        self.progress_pct.configure(text_color=p["ink_2"])

        self.log_display.configure(fg_color=p["surface"], text_color=p["ink_2"])

        self.status_label.configure(text_color=p[self.status_kind])

    # --- thread-safe ui updates ---
    # anything touching a widget has to run on the main thread. the worker
    # thread hands updates back with self.after(0, ...) — poke the ui straight
    # from a sub-thread and it crashes on macos.

    def update_status(self, text, kind):
        def apply():
            self.status_kind = kind
            self.status_label.configure(text=text, text_color=self.palette[kind])
        self.after(0, apply)

    def update_progress(self, value):
        def apply():
            self.progress.set(value)
            self.progress_pct.configure(text=f"{round(value * 100)}%")
        self.after(0, apply)

    def log(self, message):
        self.after(0, self._append_log, message)

    def _append_log(self, message):
        self.log_display.configure(state="normal")
        self.log_display.insert("end", f"· {message}\n")
        self.log_display.see("end")
        self.log_display.configure(state="disabled")

    def toggle_ui(self, is_running):
        state_run = "disabled" if is_running else "normal"
        state_stop = "normal" if is_running else "disabled"
        self.after(0, lambda: self.btn_run.configure(state=state_run))
        self.after(0, lambda: self.btn_stop.configure(state=state_stop))
        self.after(0, lambda: self.btn_src.configure(state=state_run))
        self.after(0, lambda: self.btn_dest.configure(state=state_run))

    # --- core logic ---

    def check_ai_connection(self):
        """on startup, make sure ollama is up and the model is actually pulled —
        the two most common 'why isn't it working' reasons, caught upfront."""
        try:
            names = [m.model for m in ollama.list().models]
            if any(n.startswith(OLLAMA_MODEL) for n in names):
                self.update_status(f"{OLLAMA_MODEL} ready", "pos")
            else:
                self.update_status(f"Run: ollama pull {OLLAMA_MODEL}", "neg")
        except Exception:
            self.update_status("Local AI offline — start Ollama", "neg")

    def select_source(self):
        folder = ctk.filedialog.askdirectory(title="Choose the folder to sort")
        if folder:
            self.source_folder = folder
            self.src_value.configure(text=folder)
            self.btn_run.configure(state="normal")
            self.log(f"Source set to {folder}")

    def select_dest(self):
        folder = ctk.filedialog.askdirectory(title="Choose where sorted files go")
        if folder:
            self.dest_folder = folder
            self.dest_value.configure(text=folder)
            self.log(f"Destination set to {folder}")

    def request_stop(self):
        """flag the run to bail out — but only after the current file finishes,
        so we never interrupt a move half-done and lose a file."""
        if self.is_running:
            self.stop_requested = True
            self.btn_stop.configure(text="Stopping…", state="disabled")
            self.log("Stopping after the current file…")

    def start_process(self):
        if not self.source_folder: return
        self.is_running = True
        self.stop_requested = False
        self.progress.set(0)
        self.toggle_ui(is_running=True)
        self.btn_stop.configure(text="Stop")

        # do the actual work off-thread so the window stays responsive.
        Thread(target=self.organize_logic, daemon=True).start()

    def get_real_extension(self, filename):
        """if a file lost its extension somewhere, take a guess from the tail of
        the name so it doesn't end up extension-less on the other side."""
        ext = os.path.splitext(filename)[1].lower()
        if ext == '':
            fname_lower = filename.lower()
            if fname_lower.endswith('pdf'): return '.pdf'
            if fname_lower.endswith('docx'): return '.docx'
            if fname_lower.endswith('pptx'): return '.pptx'
            if fname_lower.endswith('xlsx'): return '.xlsx'
        return ext

    def extract_text(self, path, filename):
        """grab a sample of the document's text for the model to read. capped so
        we're not shoving a whole textbook into the prompt."""
        ext = self.get_real_extension(filename)

        # don't even try to open something huge — it'll just hang the app.
        if os.path.getsize(path) > 50 * 1024 * 1024:  # 50 mb
            self.log("Skipped — file is larger than 50MB")
            return None

        try:
            if ext == ".pdf":
                with pdfplumber.open(path) as pdf:
                    return "".join([p.extract_text() or "" for p in pdf.pages[:2]])[:1200]

            elif ext == ".docx":
                doc = Document(path)
                return "\n".join([p.text for p in doc.paragraphs[:15]])[:1200]

            elif ext == ".pptx":
                prs = Presentation(path)
                text_content = []
                for slide in list(prs.slides)[:3]:
                    for shape in slide.shapes:
                        # python-pptx shapes don't all have text, and pylance
                        # doesn't know that, so getattr our way through safely.
                        if getattr(shape, "has_text_frame", False):
                            text_frame = getattr(shape, "text_frame", None)
                            if text_frame:
                                for paragraph in getattr(text_frame, "paragraphs", []):
                                    text_content.append(getattr(paragraph, "text", ""))
                return "\n".join(text_content)[:1200]

            elif ext == ".xlsx":
                # data_only so we read the actual values, not the formulas.
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                return f"Workbook sheets: {', '.join(wb.sheetnames)}"

            return None
        except Exception:
            self.log("Couldn't read file — skipping")
            return None

    def get_ai_decision(self, filename, content):
        """hybrid router: regex grabs the dead-obvious stuff instantly, then the
        model handles the judgment calls. it's pinned to a json schema, so it
        can't wander off and hand back a paragraph instead of an answer."""
        # fast path — a real course code in the name is a school file, full stop,
        # and i trust that over whatever the model thinks. the catch: "scan0042"
        # and "IMG_1234" look exactly like course codes, so blocklist the usual
        # camera/scanner/generic junk prefixes before trusting the match.
        course_match = re.search(r'\b([A-Za-z]{3,4})\s?(\d{3,5})\b', filename)
        forced_course = None
        if course_match:
            prefix = course_match.group(1).lower()
            junk_prefixes = {"scan", "img", "dsc", "dcim", "mvi", "vid", "mov",
                             "pxl", "pic", "doc", "file", "page", "copy", "scr"}
            if prefix not in junk_prefixes:
                forced_course = (course_match.group(1) + course_match.group(2)).upper()

        school_keywords = ["assignment", "rubric", "lecture", "homework", "math", "calculus", "course", "exam", "syllabus"]
        is_obvious_school = any(word in filename.lower() for word in school_keywords)

        # slow path — hand the model the name plus a peek at the contents and let
        # it reason about where the file actually belongs.
        system_prompt = (
            "You sort files into folders. Read the filename and the text sample, "
            "then decide where the file belongs based on what it's actually about "
            "— not just its name. Answer in the given JSON schema, nothing else."
        )
        user_prompt = f"""Filename: {filename}
Text sample: {content[:1000] if content else '(no readable text)'}

Categories:
- SCHOOL: coursework, lectures, assignments, syllabi, exams
- WORK: job docs, resumes, meeting notes, client material
- FINANCE: invoices, receipts, statements, taxes, budgets
- PERSONAL: anything personal that isn't the above
- UNSORTED: only if you genuinely can't tell

subject: for SCHOOL, the course code (e.g. INFO2603) or the subject name; otherwise "General".
new_name: a short, readable PascalCase name for the file. no extension."""

        try:
            response = ollama.chat(
                model=OLLAMA_MODEL,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                format=ROUTING_SCHEMA,        # forces valid json in the exact shape we want
                options={"temperature": 0},   # deterministic — the same file sorts the same way every time
            )
            data = json.loads(response["message"]["content"])

            category = str(data.get("category", "UNSORTED")).upper().strip()
            subject = str(data.get("subject", "General")).strip().replace(" ", "")
            new_name = str(data.get("new_name", "")).strip()

            # regex beats the model when it actually found a course code.
            if forced_course:
                category = "SCHOOL"
                subject = forced_course
            elif is_obvious_school and category != "SCHOOL":
                category = "SCHOOL"

            # tidy up the edge cases.
            if subject.upper() in ["NONE", "", "N/A", "NULL"]:
                subject = "General"
            if category not in VALID_CATEGORIES:
                category = "UNSORTED"

            # never trust the model with the extension — strip anything illegal
            # from its name and staple the real extension back on.
            base = os.path.splitext(new_name)[0] if new_name else os.path.splitext(filename)[0]
            base = re.sub(r'[\\/:*?"<>|]', "", base).strip()
            new_name = f"{base}{self.get_real_extension(filename)}" if base else filename

            return {"category": category, "subject": subject, "new_name": new_name}
        except Exception:
            self.log("Couldn't read the model's answer — leaving this one unsorted")
            return None

    def organize_logic(self):
        """the main loop: walk the folder, read each file, ask where it goes, move it."""
        try:
            files = [f for f in os.listdir(self.source_folder) if not f.startswith('.')]
            # skip our own output folders so a re-run doesn't try to sort what it
            # already sorted and loop on itself.
            files = [f for f in files if f != "None" and f != "UNSORTED" and f not in VALID_CATEGORIES]

            total_files = len(files)
            if total_files == 0:
                self.log("Nothing to sort — folder is empty.")
                return

            self.log(f"Sorting {total_files} files…")

            for index, filename in enumerate(files):
                # bail here if stop was pressed — safe spot, between files.
                if self.stop_requested: break

                path = os.path.join(self.source_folder, filename)
                if os.path.isdir(path): continue

                self.log(f"Reading {filename}")
                content = self.extract_text(path, filename)
                decision = self.get_ai_decision(filename, content)

                # if the model bailed entirely, drop the file in UNSORTED rather
                # than lose track of it.
                cat = decision["category"] if decision else "UNSORTED"
                sub = decision["subject"] if decision else "General"
                name = decision["new_name"] if decision else filename

                # build the folder path under wherever the user pointed us.
                target = os.path.join(self.dest_folder, cat)
                if cat == "SCHOOL":
                    target = os.path.join(target, sub)

                os.makedirs(target, exist_ok=True)
                final_path = os.path.join(target, name)

                # name clash? tack on a timestamp instead of clobbering the file.
                if os.path.exists(final_path):
                    base, ext = os.path.splitext(name)
                    final_path = os.path.join(target, f"{base}_{int(time.time())}{ext}")

                try:
                    shutil.move(path, final_path)
                    self.log(f"Sorted into {cat}/{sub}")
                except Exception as e:
                    self.log(f"Couldn't move file: {e}")

                self.update_progress((index + 1) / total_files)

            self.log("Done.")

        except Exception as e:
            self.log(f"Something went wrong: {e}")
        finally:
            # always hand the ui back, even if the loop blew up mid-run.
            self.is_running = False
            self.toggle_ui(is_running=False)


if __name__ == "__main__":
    app = FrankenSorterApp()
    app.mainloop()
