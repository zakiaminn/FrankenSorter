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

# --- UI Configuration (Retro-Modern / Cyberpunk Vibe) ---
ctk.set_appearance_mode("Dark")
BG_COLOR = "#15151e"
PANEL_COLOR = "#1e1e2b"
ACCENT_COLOR = "#00f0ff"  # Neon Cyan
TEXT_COLOR = "#e0e0e0"

# ASCII Pixel Art Logo for FrankenSorter
PIXEL_LOGO = """
███████╗██████╗  █████╗ ███╗   ██╗██╗  ██╗███████╗███╗   ██╗
██╔════╝██╔══██╗██╔══██╗████╗  ██║██║ ██╔╝██╔════╝████╗  ██║
█████╗  ██████╔╝███████║██╔██╗ ██║█████╔╝ █████╗  ██╔██╗ ██║
██╔══╝  ██╔══██╗██╔══██║██║╚██╗██║██╔═██╗ ██╔══╝  ██║╚██╗██║
██║     ██║  ██║██║  ██║██║ ╚████║██║  ██╗███████╗██║ ╚████║
╚═╝     ╚═╝  ╚═╝╚═╝  ╚═╝╚═╝  ╚═══╝╚═╝  ╚═╝╚══════╝╚═╝  ╚═══╝
                        - S O R T E R -
"""

class FrankenSorterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("FrankenSorter: Neural Directory Engine")
        self.geometry("750x800") 
        self.minsize(700, 750)
        self.configure(fg_color=BG_COLOR)
        
        # --- Application State Variables ---
        # We track state here to prevent users from spam-clicking buttons
        # while the background threads are busy doing heavy I/O tasks.
        self.source_folder = ""
        
        # OS-Agnostic Default: Automatically points to the user's Desktop
        # os.path.expanduser gracefully handles Mac (~/) and Windows (C:\Users\...) paths.
        self.dest_folder = os.path.expanduser("~/Desktop/Franken_Sorted")
        
        self.stop_requested = False
        self.is_running = False

        self.build_ui()
        
        # Start a background daemon thread to ping the local AI.
        # Daemon=True ensures this thread dies instantly when the user closes the app.
        Thread(target=self.check_ai_connection, daemon=True).start()

    def build_ui(self):
        """Constructs the graphical user interface using CustomTkinter."""
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=30, pady=20)

        # Centered Logo
        # justify="left" keeps the ASCII characters aligned with each other
        # anchor="center" centers the entire block within the window
        self.logo_label = ctk.CTkLabel(
            self.main_frame, text=PIXEL_LOGO, justify="left",
            font=ctk.CTkFont(family="Courier", size=10, weight="bold"),
            text_color=ACCENT_COLOR
        )
        self.logo_label.pack(anchor="center", pady=(0, 10))

        # Centered Title
        self.title_label = ctk.CTkLabel(
            self.main_frame, text="Neural Directory Routing System v1.0", 
            font=ctk.CTkFont(family="Courier", size=16, weight="bold"), text_color="#ffffff"
        )
        self.title_label.pack(anchor="center", pady=(0, 20))

        # --- SOURCE Selection Panel ---
        self.src_frame = ctk.CTkFrame(self.main_frame, fg_color=PANEL_COLOR, corner_radius=0, border_width=1, border_color="#333344")
        self.src_frame.pack(fill="x", pady=(0, 10))
        
        self.src_label = ctk.CTkLabel(
            self.src_frame, text="[ SOURCE: UNMAPPED ]", 
            text_color="#8888aa", font=ctk.CTkFont(family="Courier", size=12)
        )
        self.src_label.pack(side="left", padx=20, pady=10)
        
        self.btn_src = ctk.CTkButton(
            self.src_frame, text="[ SET SOURCE ]", width=120, corner_radius=0,
            command=self.select_source, fg_color="transparent", border_width=1, 
            border_color=ACCENT_COLOR, text_color=ACCENT_COLOR, hover_color="#1a3a4a"
        )
        self.btn_src.pack(side="right", padx=20, pady=10)

        # --- DESTINATION Selection Panel ---
        self.dest_frame = ctk.CTkFrame(self.main_frame, fg_color=PANEL_COLOR, corner_radius=0, border_width=1, border_color="#333344")
        self.dest_frame.pack(fill="x", pady=(0, 10))
        
        self.dest_label = ctk.CTkLabel(
            self.dest_frame, text=f"DEST: {self.dest_folder}", 
            text_color="#ffffff", font=ctk.CTkFont(family="Courier", size=12)
        )
        self.dest_label.pack(side="left", padx=20, pady=10)
        
        self.btn_dest = ctk.CTkButton(
            self.dest_frame, text="[ SET TARGET ]", width=120, corner_radius=0,
            command=self.select_dest, fg_color="transparent", border_width=1, 
            border_color="#00ff41", text_color="#00ff41", hover_color="#003300"
        )
        self.btn_dest.pack(side="right", padx=20, pady=10)

        # --- Progress & Logging ---
        self.progress = ctk.CTkProgressBar(self.main_frame, height=2, corner_radius=0, fg_color="#222", progress_color=ACCENT_COLOR)
        self.progress.pack(fill="x", pady=(10, 5))
        self.progress.set(0)

        self.log_display = ctk.CTkTextbox(
            self.main_frame, height=240, fg_color="#0d0d14", corner_radius=0, border_width=1, border_color="#333344",
            text_color="#00ff41", font=ctk.CTkFont(family="Courier", size=12) # Matrix-style terminal text
        )
        self.log_display.pack(fill="both", expand=True, pady=5)
        self.log_display.configure(state="disabled")

        # --- Bottom Control Panel ---
        self.bottom_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.bottom_frame.pack(fill="x", pady=(15, 0))

        self.status_label = ctk.CTkLabel(
            self.bottom_frame, text="SYS_STATUS: OFFLINE", 
            text_color="#ff3366", font=ctk.CTkFont(family="Courier", size=12, weight="bold")
        )
        self.status_label.pack(side="left")

        self.btn_stop = ctk.CTkButton(
            self.bottom_frame, text="[ ABORT ]", width=120, corner_radius=0,
            command=self.request_stop, state="disabled", 
            fg_color="transparent", border_width=1, text_color="#ff3366", border_color="#ff3366", hover_color="#3a1c22"
        )
        self.btn_stop.pack(side="right", padx=(15, 0))

        self.btn_run = ctk.CTkButton(
            self.bottom_frame, text="[ EXECUTE_SORT ]", width=140, corner_radius=0,
            command=self.start_process, state="disabled",
            fg_color=ACCENT_COLOR, text_color="#000000", hover_color="#00b3cc", font=ctk.CTkFont(weight="bold")
        )
        self.btn_run.pack(side="right")

    # --- UI Updaters (Thread-Safe) ---
    # We use self.after(0, ...) to pass UI updates from the background thread 
    # back to the main thread. Directly modifying UI from a sub-thread causes crashes on macOS.
    
    def update_status(self, text, color):
        self.after(0, lambda: self.status_label.configure(text=text, text_color=color))

    def update_progress(self, value):
        self.after(0, lambda: self.progress.set(value))

    def log(self, message):
        self.after(0, self._append_log, message)

    def _append_log(self, message):
        self.log_display.configure(state="normal")
        self.log_display.insert("end", f"> {message}\n")
        self.log_display.see("end")
        self.log_display.configure(state="disabled")

    def toggle_ui(self, is_running):
        state_run = "disabled" if is_running else "normal"
        state_stop = "normal" if is_running else "disabled"
        self.after(0, lambda: self.btn_run.configure(state=state_run))
        self.after(0, lambda: self.btn_stop.configure(state=state_stop))
        self.after(0, lambda: self.btn_src.configure(state=state_run))
        self.after(0, lambda: self.btn_dest.configure(state=state_run))

    # --- Core Application Logic ---

    def check_ai_connection(self):
        """Pings the local Ollama service to ensure the LLM is running."""
        try:
            ollama.list()
            self.update_status("SYS_STATUS: Llama-3.2 ONLINE", ACCENT_COLOR)
        except Exception:
            self.update_status("SYS_STATUS: OLLAMA OFFLINE", "#ff3366")

    def select_source(self):
        folder = ctk.filedialog.askdirectory(title="Map Source Directory")
        if folder:
            self.source_folder = folder
            self.src_label.configure(text=f"SRC: {folder}", text_color="#ffffff")
            self.btn_run.configure(state="normal")
            self.log(f"SOURCE_MAPPED: {folder}")

    def select_dest(self):
        folder = ctk.filedialog.askdirectory(title="Map Destination Directory")
        if folder:
            self.dest_folder = folder
            self.dest_label.configure(text=f"DEST: {folder}")
            self.log(f"TARGET_MAPPED: {folder}")

    def request_stop(self):
        """Flags the process to stop safely after the current iteration finishes."""
        if self.is_running:
            self.stop_requested = True
            self.btn_stop.configure(text="[ HALTING... ]", state="disabled")
            self.log("WARN: Interrupt received. Halting after current file...")

    def start_process(self):
        if not self.source_folder: return
        self.is_running = True
        self.stop_requested = False
        self.progress.set(0)
        self.toggle_ui(is_running=True)
        self.btn_stop.configure(text="[ ABORT ]")
        
        # Push the heavy lifting to a background thread to keep the UI responsive.
        Thread(target=self.organize_logic, daemon=True).start()

    def get_real_extension(self, filename):
        """
        Fallback logic: If a file loses its extension during a bad rename, 
        we attempt to deduce it heuristically from the filename string itself.
        """
        ext = os.path.splitext(filename)[1].lower()
        if ext == '':
            fname_lower = filename.lower()
            if fname_lower.endswith('pdf'): return '.pdf'
            if fname_lower.endswith('docx'): return '.docx'
            if fname_lower.endswith('pptx'): return '.pptx'
            if fname_lower.endswith('xlsx'): return '.xlsx'
        return ext

    def extract_text(self, path, filename):
        """
        ETL Pipeline: Extracts the first 1200 characters of a document.
        We cap the extraction length to avoid overwhelming the LLM's context window.
        """
        ext = self.get_real_extension(filename)
        
        # Prevent attempting to read massive files which could freeze the system
        if os.path.getsize(path) > 50 * 1024 * 1024:  # 50 MB limit
            self.log(f"READ_SKIP: File exceeds 50MB limit.")
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
                        # Defensive programming: Bypass Pylance type-checker limitations 
                        # using getattr() to safely check for text frames in PPTX objects.
                        if getattr(shape, "has_text_frame", False):
                            text_frame = getattr(shape, "text_frame", None)
                            if text_frame:
                                for paragraph in getattr(text_frame, "paragraphs", []):
                                    text_content.append(getattr(paragraph, "text", ""))
                return "\n".join(text_content)[:1200]
            
            elif ext == ".xlsx":
                # Data_only=True ensures we get cell values, not Excel formulas
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                return f"Workbook sheets: {', '.join(wb.sheetnames)}"
            
            return None
        except Exception as e:
            self.log("READ_FAIL: Cannot parse binary/corrupt data.")
            return None

    def get_ai_decision(self, filename, content):
        """
        Hybrid Routing Engine: Uses Regex heuristics to catch obvious patterns,
        then falls back to Local Llama 3.2 for complex semantic routing.
        """
        # --- 1. HYBRID HEURISTICS (Fast Path) ---
        course_match = re.search(r'([A-Za-z]{3,4}\s?\d{3,5})', filename)
        forced_course = course_match.group(1).upper().replace(" ", "") if course_match else None

        school_keywords = ["assignment", "rubric", "lecture", "homework", "math", "calculus", "course", "exam", "syllabus"]
        is_obvious_school = any(word in filename.lower() for word in school_keywords)

        # --- 2. LLM INFERENCE (Slow Path) ---
        system_prompt = "You are a strict data routing node. Output ONLY valid JSON."
        user_prompt = f"""
        ANALYZE FILE TARGET:
        Filename: "{filename}"
        Extracted Text: "{content[:500] if content else 'Data unavailable.'}"

        ROUTING RULES:
        1. Category MUST be one of: [SCHOOL, WORK, PERSONAL, FINANCE]. 
        2. Subject: If SCHOOL, output the Course Code (e.g. INFO23875). If unknown, "General".
        3. new_name: PascalCase version of the filename. Keep the extension.

        OUTPUT FORMAT:
        {{"category": "SCHOOL", "subject": "MATH101", "new_name": "MathAssignment.jpg"}}
        """
        
        try:
            # We pass the prompt to the local Ollama instance
            response = ollama.chat(model='llama3.2', messages=[
                {'role': 'system', 'content': system_prompt},
                {'role': 'user', 'content': user_prompt}
            ])
            res = response['message']['content']
            
            # Use Regex to extract the JSON block, ignoring any conversational hallucinations
            match = re.search(r'\{[\s\S]*\}', res)
            if not match: return None
                
            data = json.loads(match.group(0))
            
            category = str(data.get("category", "UNSORTED")).upper().strip()
            subject = str(data.get("subject", "General")).title().replace(" ", "").strip()
            new_name = str(data.get("new_name", filename)).strip()

            # --- 3. OVERRIDE MERGE ---
            # If our deterministic regex found a course, we trust it over the AI.
            if forced_course:
                category = "SCHOOL"
                subject = forced_course
            elif is_obvious_school and category != "SCHOOL":
                category = "SCHOOL"
            
            # Final data sanitization
            if subject.upper() in ["NONE", "", "N/A", "NULL"]: subject = "General"
            if category not in ["SCHOOL", "WORK", "PERSONAL", "FINANCE"]: category = "UNSORTED"
            
            return {
                "category": category,
                "subject": subject,
                "new_name": new_name
            }
        except Exception as e:
            self.log(f"WARN: AI Parsing Failed")
            return None

    def organize_logic(self):
        """Main loop that iterates through the directory and applies the ETL & Routing logic."""
        try:
            files = [f for f in os.listdir(self.source_folder) if not f.startswith('.')]
            # Prevent the system from entering an infinite loop by trying to sort its own output folders
            files = [f for f in files if f != "None" and f != "UNSORTED" and f not in ["SCHOOL", "WORK", "PERSONAL", "FINANCE"]]
            
            total_files = len(files)
            if total_files == 0:
                self.log("DIR_EMPTY: No valid targets found.")
                return

            self.log(f"BATCH_START: Processing {total_files} objects...")

            for index, filename in enumerate(files):
                # Check for safe interrupt flag
                if self.stop_requested: break
                
                path = os.path.join(self.source_folder, filename)
                if os.path.isdir(path): continue

                self.log(f"ANALYZING: {filename}")
                content = self.extract_text(path, filename)
                decision = self.get_ai_decision(filename, content)

                # Fallback to UNSORTED if the AI completely crashed
                cat = decision["category"] if decision else "UNSORTED"
                sub = decision["subject"] if decision else "General"
                name = decision["new_name"] if decision else filename

                # Build the dynamic directory path using the user's selected destination
                target = os.path.join(self.dest_folder, cat)
                if cat == "SCHOOL": 
                    target = os.path.join(target, sub)
                
                os.makedirs(target, exist_ok=True)
                final_path = os.path.join(target, name)
                
                # Collision handling: append epoch timestamp to prevent overwriting
                if os.path.exists(final_path):
                    base, ext = os.path.splitext(name)
                    final_path = os.path.join(target, f"{base}_{int(time.time())}{ext}")
                
                try:
                    shutil.move(path, final_path)
                    self.log(f"OK -> Routed to {cat}/{sub}")
                except Exception as e:
                    self.log(f"ERR -> I/O Exception: {e}")

                self.update_progress((index + 1) / total_files)

            self.log("BATCH_COMPLETE: Process terminated.")
        
        except Exception as e:
            self.log(f"SYS_FATAL: {e}")
        finally:
            # Always reset UI state, even if the loop crashed
            self.is_running = False
            self.toggle_ui(is_running=False)

if __name__ == "__main__":
    app = FrankenSorterApp()
    app.mainloop()