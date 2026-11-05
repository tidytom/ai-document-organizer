import os
import sys
import importlib
import importlib.util
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import threading
import json
import csv
import shutil
import hashlib
import re
from datetime import datetime

# --- 1. DEPENDENCY CHECK ---
# (pip package name, import name)
REQUIRED_PACKAGES = [
    ("requests", "requests"),
    ("PyMuPDF", "fitz"),
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("python-docx", "docx"),
]


def check_dependencies():
    """Show a clear message and stop if a required package is missing. Nothing is installed automatically."""
    missing = [pip_name for pip_name, import_name in REQUIRED_PACKAGES
               if importlib.util.find_spec(import_name) is None]
    if not missing:
        return
    message = (
        "Some required Python packages are missing:\n  " + ", ".join(missing) + "\n\n"
        "Open a command prompt in this program's folder and run:\n"
        "  pip install -r requirements.txt\n\n"
        "Then start the program again."
    )
    print(message, file=sys.stderr)
    try:
        _root = tk.Tk()
        _root.withdraw()
        messagebox.showerror("Missing packages", message)
        _root.destroy()
    except Exception:
        pass
    sys.exit(1)


check_dependencies()
import requests  # noqa: E402
try:
    import pymupdf as fitz  # noqa: E402  (PyMuPDF 1.24 and newer)
except ImportError:
    import fitz  # noqa: E402  (older PyMuPDF)
import pandas as pd  # noqa: E402
import docx  # noqa: E402  (python-docx)

# --- 1b. SETTINGS (config.json next to this script) ---
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
DEFAULT_CONFIG = {
    "ollama_url": "http://localhost:11434",
    "model": "llama3.1",
    "categories": [
        "Financial", "Taxes", "Legal", "Business", "Receipts", "Insurance",
        "Medical", "Government", "Education", "Personal", "Software", "Misc",
    ],
}


def load_config():
    """Read config.json. Missing or broken values fall back to the defaults above."""
    config = dict(DEFAULT_CONFIG)
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            for key in DEFAULT_CONFIG:
                if key in data and data[key]:
                    config[key] = data[key]
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"Could not read {CONFIG_PATH}: {e}. Using default settings.", file=sys.stderr)
    if not isinstance(config["categories"], list):
        config["categories"] = list(DEFAULT_CONFIG["categories"])
    config["categories"] = [str(c).strip() for c in config["categories"] if str(c).strip()]
    config["ollama_url"] = str(config["ollama_url"]).rstrip("/")
    config["model"] = str(config["model"])
    return config

# --- 2. CORE UTILITIES ---
def get_md5(filepath):
    hasher = hashlib.md5()
    try:
        with open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(65536), b""):
                hasher.update(chunk)
        return hasher.hexdigest()
    except Exception:
        return None

def extract_text(filepath):
    ext = filepath.lower().split('.')[-1]
    text_sample = ""
    try:
        if ext == 'pdf':
            doc = fitz.open(filepath)
            for page in doc[:3]: 
                text_sample += page.get_text()
            doc.close()
        elif ext in ['xls', 'xlsx']:
            df = pd.read_excel(filepath, nrows=50) 
            text_sample = df.to_string()
        elif ext == 'docx':
            doc = docx.Document(filepath)
            for para in doc.paragraphs[:20]:
                text_sample += para.text + "\n"
        elif ext == 'txt':
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                text_sample = f.read(2000) 
        else:
            # Any other file type is organized by its file name only.
            return "UNSUPPORTED_FORMAT"
    except Exception:
        return "ERROR_READING_FILE"
        
    return text_sample[:2000].strip() 

def clean_string(name):
    cleaned = re.sub(r'[\*\[\]"\'`<>:/\\|?]', '', name)
    cleaned = cleaned.replace(" ", "_")
    return cleaned.strip('._- ')

def safe_folder_name(name):
    return re.sub(r'[\*\[\]"\'`<>:/\\|?]', '', name).strip(' .')


def original_stem(original_name):
    """The original file name without its extension, cleaned for use as a new name."""
    return clean_string(os.path.splitext(original_name)[0]) or "Unnamed_File"


FIELD_DEFAULTS = ["", "Unknown_Entity", "Uncategorized", "Unknown_Year", "Unknown_Institution"]


def fill_empty_fields(fields, original_name):
    """Replace empty AI answers so files never end up named just '.pdf' or in a blank folder."""
    fields = list(fields)[:5] + [""] * (5 - len(fields))
    result = []
    for i, value in enumerate(fields):
        value = (value or "").strip()
        if not value:
            value = original_stem(original_name) if i == 0 else FIELD_DEFAULTS[i]
        result.append(value)
    return tuple(result)

# --- 3. UNIFIED APPLICATION ---
class DocumentOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.config = load_config()
        self.log_file = None
        self.root.title("AI Document Organizer")
        self.root.geometry("650x600")
        self.root.configure(bg="#f5f5f5")
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=8)
        
        # --- UI LAYOUT ---
        top_frame = tk.Frame(self.root, padx=20, pady=20, bg="#f5f5f5")
        top_frame.pack(fill=tk.X)

        lbl_title = tk.Label(top_frame, text="AI Document Organizer", font=('Segoe UI', 16, 'bold'), bg="#f5f5f5")
        lbl_title.pack(pady=(0, 15))

        self.btn_p1 = ttk.Button(top_frame, text="Phase 1: Scan Structure", command=self.start_phase1)
        self.btn_p1.pack(fill=tk.X, pady=5)

        self.btn_p2 = ttk.Button(top_frame, text="Phase 2: AI Content Parsing", command=self.start_phase2)
        self.btn_p2.pack(fill=tk.X, pady=5)

        self.btn_p3 = ttk.Button(top_frame, text="Phase 3: Reorganize & Deduplicate", command=self.start_phase3)
        self.btn_p3.pack(fill=tk.X, pady=5)

        # Built-in Live Console
        console_frame = tk.LabelFrame(self.root, text="Live Log & AI Stream", padx=10, pady=10, bg="#f5f5f5")
        console_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        self.console = scrolledtext.ScrolledText(console_frame, wrap=tk.WORD, font=('Consolas', 9), bg="#1e1e1e", fg="#d4d4d4")
        self.console.pack(fill=tk.BOTH, expand=True)
        self.console.config(state=tk.DISABLED)

        self.log_msg("System Ready. All dependencies loaded.")
        self.log_msg(f"Model: {self.config['model']} at {self.config['ollama_url']} (change these in config.json)")

    # --- LOGGING SYSTEM ---
    def setup_file_logger(self, output_dir):
        """Log files go in a Logs folder inside the program's output folder, not next to the script."""
        log_dir = os.path.join(output_dir, "Logs")
        os.makedirs(log_dir, exist_ok=True)
        date_str = datetime.now().strftime('%Y-%m-%d')
        self.log_file = os.path.join(log_dir, f"ai_organizer_{date_str}.log")

    def ollama_generate_url(self):
        return f"{self.config['ollama_url']}/api/generate"

    def log_msg(self, message, is_chunk=False):
        """Thread-safe logging that updates the GUI console and the text file."""
        timestamp = datetime.now().strftime('%H:%M:%S')
        formatted_msg = message if is_chunk else f"[{timestamp}] {message}\n"
        
        # Write to file safely (only after a phase has chosen its output folder)
        if self.log_file:
            try:
                with open(self.log_file, 'a', encoding='utf-8') as f:
                    f.write(formatted_msg)
            except Exception: pass

        # Update GUI safely using root.after
        self.root.after(0, self._update_console, formatted_msg)

    def _update_console(self, msg):
        self.console.config(state=tk.NORMAL)
        self.console.insert(tk.END, msg)
        self.console.see(tk.END)
        self.console.config(state=tk.DISABLED)

    # --- THREAD MANAGEMENT ---
    def disable_buttons(self):
        self.btn_p1.config(state=tk.DISABLED)
        self.btn_p2.config(state=tk.DISABLED)
        self.btn_p3.config(state=tk.DISABLED)

    def enable_buttons(self):
        self.btn_p1.config(state=tk.NORMAL)
        self.btn_p2.config(state=tk.NORMAL)
        self.btn_p3.config(state=tk.NORMAL)

    # --- PHASE 1: SCANNING ---
    def start_phase1(self):
        target = filedialog.askdirectory(title="Select Folder to Analyze")
        if not target: return
        self.disable_buttons()
        threading.Thread(target=self.run_phase1, args=(target,), daemon=True).start()

    def run_phase1(self, target_folder):
        try:
            processing_dir = os.path.join(os.path.dirname(target_folder), f"Processing_{os.path.basename(target_folder)}")
            os.makedirs(processing_dir, exist_ok=True)
            self.setup_file_logger(processing_dir)
            self.log_msg("--- PHASE 1 INITIATED ---")
            self.log_msg(f"Output folder: {processing_dir}")
            
            structure_file = os.path.join(processing_dir, "1_original_structure.txt")
            self.log_msg("Scanning directory structure...")
            
            with open(structure_file, 'w', encoding='utf-8') as f:
                for root_dir, dirs, files in os.walk(target_folder):
                    level = root_dir.replace(target_folder, '').count(os.sep)
                    indent = ' ' * 4 * (level)
                    f.write(f"{indent}{os.path.basename(root_dir)}/\n")
                    subindent = ' ' * 4 * (level + 1)
                    for file in files:
                        f.write(f"{subindent}{file}\n")
            
            self.log_msg("Sending structure to local Ollama AI...")
            with open(structure_file, 'r', encoding='utf-8') as f:
                structure_data = f.read()

            if len(structure_data) > 10000:
                structure_data = structure_data[:10000] + "\n... [Structure truncated for limits]"

            prompt = (
                "Analyze the following file and folder structure. Describe the likely purpose "
                "of these files and provide a concise summary of the organization.\n\n"
                f"{structure_data}"
            )

            response = requests.post(self.ollama_generate_url(), json={
                "model": self.config["model"], "prompt": prompt, "stream": True
            }, stream=True, timeout=120)
            
            if response.status_code == 200:
                self.log_msg("\n[AI Response Stream]:\n")
                analysis_text = ""
                for line in response.iter_lines():
                    if line:
                        chunk = json.loads(line).get("response", "")
                        analysis_text += chunk
                        self.log_msg(chunk, is_chunk=True) # Stream to GUI
                
                self.log_msg("\n\n[End of AI Response]\n")
                
                output_file = os.path.join(processing_dir, "2_ai_structure_analysis.txt")
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(analysis_text)
                self.log_msg("Phase 1 Complete.")
            else:
                self.log_msg(f"ERROR: Ollama returned status {response.status_code}")
                
        except Exception as e:
            self.log_msg(f"ERROR in Phase 1: {str(e)}")
        finally:
            self.root.after(0, self.enable_buttons)

    # --- PHASE 2: PROCESSING ---
    def start_phase2(self):
        target = filedialog.askdirectory(title="Select Target Directory to Process")
        if not target: return
        self.disable_buttons()
        threading.Thread(target=self.run_phase2, args=(target,), daemon=True).start()

    def get_ai_suggestions(self, text_sample, original_name):
        is_unreadable = text_sample in ["UNSUPPORTED_FORMAT", "ERROR_READING_FILE"] or not text_sample
        category_list = ", ".join(self.config["categories"])

        if is_unreadable:
            prompt = (
                f"I cannot read the text of this file. Based ONLY on the filename '{original_name}', "
                "extract the metadata to organize it.\n"
                "1. FileName: Clean up the filename to be descriptive (under 5 words). No extension.\n"
                "2. Entity: The name of the specific business or person. If unclear, use 'Unknown_Entity'.\n"
                f"3. Category: MUST be one of: [{category_list}].\n"
                "4. Year: The 4-digit year (e.g., 2023). If none found, use 'Unknown_Year'.\n"
                "5. Institution: The bank, agency, software, or vendor. If unclear, use 'Unknown_Institution'.\n"
                "Format response EXACTLY as a single line separated by pipes (|):\n"
                "FileName | Entity | Category | Year | Institution\n"
            )
        else:
            prompt = (
                f"Analyze the text extracted from the file '{original_name}'.\n"
                "Extract the following metadata to organize this document.\n"
                "1. FileName: A short, descriptive name (under 5 words). No extension.\n"
                "2. Entity: The name of the specific business or person. If unclear, use 'Unknown_Entity'.\n"
                f"3. Category: MUST be one of: [{category_list}].\n"
                "4. Year: The 4-digit year (e.g., 2023). If none found, use 'Unknown_Year'.\n"
                "5. Institution: The name of the bank, agency, or vendor. If unclear, use 'Unknown_Institution'.\n"
                "Format response EXACTLY as a single line separated by pipes (|):\n"
                "FileName | Entity | Category | Year | Institution\n\n"
                f"Text:\n{text_sample}"
            )

        try:
            response = requests.post(self.ollama_generate_url(), json={
                "model": self.config["model"], "prompt": prompt, "stream": False
            }, timeout=60)
            
            if response.status_code == 200:
                result = response.json()['response'].strip()
                if "|" in result:
                    lines = [line for line in result.split('\n') if '|' in line]
                    if lines:
                        parts = lines[-1].split("|")
                        if len(parts) >= 5:
                            return fill_empty_fields([clean_string(p) for p in parts[:5]], original_name)
                        elif len(parts) >= 2: 
                            return fill_empty_fields([clean_string(parts[0]), "", clean_string(parts[1]), "", ""], original_name)
            else:
                self.log_msg(f"ERROR: Ollama returned status {response.status_code} for {original_name}")
        except requests.exceptions.Timeout:
            self.log_msg(f"WARNING: Ollama timed out on {original_name}. Skipping to prevent hang.")
        except Exception as e:
            self.log_msg(f"ERROR: API error: {str(e)}")
            
        return original_stem(original_name), "Unknown_Entity", "AI_Error", "Unknown_Year", "Unknown_Institution"

    def run_phase2(self, target_dir):
        try:
            processing_dir = os.path.join(os.path.dirname(target_dir), f"Processing_{os.path.basename(target_dir)}")
            os.makedirs(processing_dir, exist_ok=True)
            self.setup_file_logger(processing_dir)
            self.log_msg("--- PHASE 2 INITIATED ---")
            self.log_msg(f"Output folder: {processing_dir}")
            csv_path = os.path.join(processing_dir, "3_ai_renaming_map.csv")
            
            self.log_msg("Starting Document Analysis...")
            
            with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Original Path", "Original Name", "Proposed Name", "Entity", "Category", "Year", "Institution"])
                
                for root_dir, dirs, files in os.walk(target_dir):
                    for file in files:
                        filepath = os.path.join(root_dir, file)
                        self.log_msg(f"Parsing: {file}")
                        
                        text = extract_text(filepath)
                        p_name, p_ent, p_cat, p_year, p_inst = self.get_ai_suggestions(text, file)
                        
                        writer.writerow([filepath, file, p_name, p_ent, p_cat, p_year, p_inst])
                        if p_cat not in ["Uncategorized", "AI_Error"]:
                            self.log_msg(f"  Result: {p_ent} | {p_cat} | {p_year} | {p_inst} | {p_name}")
            
            self.log_msg(f"Phase 2 Complete. Please review the CSV file in:\n{processing_dir}")
        except Exception as e:
            self.log_msg(f"ERROR in Phase 2: {str(e)}")
        finally:
            self.root.after(0, self.enable_buttons)

    # --- PHASE 3: ORGANIZATION ---
    def start_phase3(self):
        p_dir = filedialog.askdirectory(title="Select the Processing Directory (from Phase 1 & 2)")
        if not p_dir: return
        self.disable_buttons()
        threading.Thread(target=self.run_phase3, args=(p_dir,), daemon=True).start()

    def run_phase3(self, processing_dir):
        try:
            self.log_msg("--- PHASE 3 INITIATED ---")
            csv_path = os.path.join(processing_dir, "3_ai_renaming_map.csv")
            if not os.path.exists(csv_path):
                self.log_msg("ERROR: CSV map not found. Ensure Phase 2 completed successfully.")
                return
            self.setup_file_logger(processing_dir)

            base_name = os.path.basename(processing_dir).replace("Processing_", "")
            parent_dir = os.path.dirname(processing_dir)
            organized_dir = os.path.join(parent_dir, f"{base_name}_Organized")
            os.makedirs(organized_dir, exist_ok=True)
            self.log_msg(f"Creating Master Structure: {organized_dir}")

            seen_hashes = {} 
            
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    orig_path = row.get("Original Path")
                    orig_name = row.get("Original Name") or os.path.basename(orig_path or "")
                    # Empty cells (from the AI or from editing the CSV) fall back to safe defaults.
                    prop_name = clean_string(row.get("Proposed Name") or "") or original_stem(orig_name)
                    
                    entity = safe_folder_name(row.get("Entity") or "") or "Unknown_Entity"
                    category = safe_folder_name(row.get("Category") or "") or "Uncategorized"
                    year = safe_folder_name(row.get("Year") or "") or "Unknown_Year"
                    institution = safe_folder_name(row.get("Institution") or "") or "Unknown_Institution"
                    
                    if not orig_path or not os.path.exists(orig_path):
                        continue
                        
                    _, ext = os.path.splitext(orig_name)
                    # None means the file could not be read; such files are always treated as unique.
                    file_hash = get_md5(orig_path)
                    
                    target_folder = os.path.join(organized_dir, entity, category, year, institution)
                    os.makedirs(target_folder, exist_ok=True)
                    
                    if file_hash is not None and file_hash in seen_hashes:
                        dup_count = 1
                        while True:
                            new_filename = f"{prop_name}_DUPLICATE_{dup_count:02d}{ext}"
                            dest_path = os.path.join(target_folder, new_filename)
                            if not os.path.exists(dest_path): break
                            dup_count += 1
                        self.log_msg(f"Duplicate Found: Naming -> {new_filename}")
                    else:
                        new_filename = f"{prop_name}{ext}"
                        dest_path = os.path.join(target_folder, new_filename)
                        
                        collision_count = 1
                        while os.path.exists(dest_path):
                            new_filename = f"{prop_name}_{collision_count}{ext}"
                            dest_path = os.path.join(target_folder, new_filename)
                            collision_count += 1
                            
                        if file_hash is not None:
                            seen_hashes[file_hash] = dest_path
                        self.log_msg(f"Copying: {new_filename} -> {entity}/{category}/...")
                    
                    try:
                        shutil.copy2(orig_path, dest_path)
                    except Exception as e:
                        self.log_msg(f"ERROR copying {orig_name}: {str(e)}")

            self.log_msg("Phase 3 Complete! Your files are successfully organized.")
        except Exception as e:
            self.log_msg(f"ERROR in Phase 3: {str(e)}")
        finally:
            self.root.after(0, self.enable_buttons)

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentOrganizerApp(root)
    root.mainloop()