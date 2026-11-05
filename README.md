# AI Document Organizer
A Windows desktop program that uses a local AI model (through Ollama) to suggest new names and folders for your documents, then copies them into a neat folder structure.

## What it does
- Reads the text of PDF, Word (.docx), Excel (.xlsx and .xls) and text (.txt) files.
- Asks a local AI model for a short file name, the person or business, a category, the year, and the bank, agency or vendor.
- Other file types (for example photos, ZIP files or programs) are sorted by their file name only.
- Saves all suggestions in a CSV file that you can check and edit before anything is copied.
- **COPIES** your files into a new folder, sorted as Entity / Category / Year / Institution. It never moves or deletes your original files.
- "Deduplicate" means: when two files have exactly the same content, the second copy gets `_DUPLICATE_01` (02, 03 ...) added to its name. Nothing is removed.

## What you need
- Windows 10 or 11 (it also runs on macOS and Linux).
- Python 3.10 or newer.
- [Ollama](https://ollama.com) running on your computer, with a model downloaded. The default model is `llama3.1` (`ollama pull llama3.1`).
- Everything runs on your own computer. Your documents are sent only to the Ollama address in `config.json` (your own computer by default).

## Install
1. Install Python from python.org. During setup, tick "Add Python to PATH".
2. Install Ollama and run `ollama pull llama3.1` in a command prompt.
3. Download this project and unzip it.
4. Open a command prompt in the project folder.
5. Run `pip install -r requirements.txt`

## How to use
1. Make sure Ollama is running.
2. Start the program: `py ai_document_organizer.py`
3. Optional: click **Phase 1: Scan Structure** and choose your folder. The AI writes a short summary of how the folder is organized. This step is for information only.
4. Click **Phase 2: AI Content Parsing** and choose the same folder. The program reads each file and asks the AI for suggestions. This can take a while for big folders. Watch the **Live Log & AI Stream** box.
5. Open the new `Processing_<your folder name>` folder (it is next to your folder). Open `3_ai_renaming_map.csv` in Excel and fix any names or folders you do not like. Save it as CSV.
6. Click **Phase 3: Reorganize & Deduplicate** and choose the `Processing_<your folder name>` folder.
7. Your sorted copies are in the new `<your folder name>_Organized` folder, next to your original folder.

## Settings and options
Edit `config.json` in the program folder with any text editor:
- `ollama_url`: where Ollama runs. Default `http://localhost:11434`.
- `model`: the Ollama model name. Default `llama3.1`. Any model you have pulled will work.
- `categories`: the list of categories the AI must choose from. Add, remove or rename them to fit your files.

If `config.json` is missing or has a mistake, the program uses the defaults above.

## Safety notes
- The program only **copies** files. Your original folder is not changed.
- It never overwrites a file. If a name is already taken, it adds `_1`, `_2` and so on.
- Files with the same content are still copied, with `_DUPLICATE_` in the name, so you can decide what to delete yourself.
- Files that cannot be read are always treated as unique, never as duplicates.
- The `Processing_...` folder holds the CSV map, the AI summaries and a `Logs` folder. These contain your file names, so do not share them if the names are private.
- To undo, delete the `_Organized` and `Processing_...` folders. Your originals are still in place.

## Troubleshooting
- **"Some required Python packages are missing"**: run `pip install -r requirements.txt`, then start the program again.
- **"ERROR: API error" or "Connection refused"**: Ollama is not running. Start Ollama and try again.
- **"Ollama returned status 404"**: the model in `config.json` is not downloaded. Run `ollama pull <model name>`.
- **Many files end up in `AI_Error`**: the AI did not answer in time. Try a smaller or faster model.
- **"CSV map not found"** in Phase 3: choose the `Processing_...` folder, not your original folder, and make sure Phase 2 finished.

## How it works
The program is a Tkinter desktop app. Text is pulled from the first pages of each file with PyMuPDF (PDF), python-docx (Word) and pandas/openpyxl (Excel), and sent to a local Ollama model through its HTTP API with a prompt that asks for five pipe-separated fields. Results go into a CSV map that a person can review. Phase 3 reads that map, builds the folder tree, and copies files with MD5 checks to label exact duplicates. All long work runs on a background thread, and the log updates the window through `root.after`.

## License

MIT. See [LICENSE](LICENSE).
