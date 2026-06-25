# Running the Newspaper Analyzer Locally

This guide documents how to set up and run the **Newspaper Analyzer** —
a tool that reads newspapers.com PDF clippings with Claude Vision,
classifies headlines and subheadlines, detects photographs, and exports
the results to CSV/Excel — on your own computer.

It covers a Windows + Anaconda setup end-to-end, with notes for Mac users
as well.

## What the tool does

1. You upload one or more newspaper clipping PDFs through a local web page.
2. Each page is sent to Anthropic's Claude (Vision) model for analysis.
3. The tool classifies each text element (headline, subheadline, body
   text), detects whether a photograph is present, and extracts metadata
   such as whether a number is mentioned, names mentioned, and locations
   mentioned.
4. You review the extracted rows in a table, select which ones you want,
   and export them to CSV and/or Excel.
5. The tool reports the API cost of each run so usage can be tracked.

## Prerequisites

| Requirement | Why it's needed | How to get it |
|---|---|---|
| Python 3.11+ | Runs the app | Comes bundled with Anaconda/Miniconda, or https://www.python.org/downloads/ |
| poppler | Converts PDF pages to images for analysis | `conda install -c conda-forge poppler` (or `brew install poppler` on Mac) |
| Anthropic API key | Required to call Claude Vision | https://console.anthropic.com → **API Keys** |
| Anthropic account credits | Each page analyzed costs ~$0.003 (Claude Haiku 4.5 Vision) | Add a payment method / credits under **Billing** in the console |

## One-time setup (Windows + Anaconda)

1. **Get the project files.** Unzip the project folder (e.g. to your
   Desktop) so you have a folder containing `app.py`,
   `newspaper_analyzer.py`, `templates/`, `requirements.txt`, and the
   launcher scripts.

2. **Open Anaconda Prompt** (Start menu → search "Anaconda Prompt").

3. **Navigate to the project folder:**
   ```
   cd C:\Users\YourName\Desktop\newspaper-analyzer
   ```

4. **Create a dedicated environment** (keeps this app's packages separate
   from your base Anaconda setup):
   ```
   conda create -n newspaper python=3.11 -y
   conda activate newspaper
   ```

5. **Install poppler:**
   ```
   conda install -c conda-forge poppler -y
   ```

6. **Install the Python dependencies:**
   ```
   pip install -r requirements.txt
   ```

7. **Add your API key.** Copy `.env.example` to a new file named `.env`
   in the project folder, then edit it so it contains:
   ```
   ANTHROPIC_API_KEY=sk-ant-...
   ```
   The `.env` file is excluded from version control (`.gitignore`), so
   your key never gets committed or shared by accident.

Setup is now complete and does not need to be repeated.

## Running the app

### Option A — double-click launcher (Windows)

Double-click `start.bat` in the project folder. It activates the
`newspaper` environment and starts the server automatically. If Windows
blocks the file as downloaded from the internet, right-click it →
**Properties** → check **Unblock** → **OK**, then try again.

### Option B — manual (any OS)

```
cd path\to\newspaper-analyzer
conda activate newspaper
python app.py
```

On Mac/Linux without Anaconda, the equivalent is the included
`run_mac.command` script, or manually:
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

### Using it

Once the server starts, you'll see:
```
Starting Newspaper Analyzer at http://localhost:5000
```
Open **http://localhost:5000** in your browser, upload a PDF, choose pages
and output format, and run the analysis. After it completes, review the
extracted rows, select the ones you want, and download as CSV/Excel.

To stop the server, close the terminal window or press `Ctrl+C` in it.

## Cost notes

- Every analysis run calls the Anthropic API and incurs a small charge
  (displayed in the UI after each run, roughly $0.003/page).
- Simply having the server running, or browsing to the page, costs
  nothing — charges only occur when a PDF is actually analyzed.

## Troubleshooting

| Symptom | Cause / Fix |
|---|---|
| `Cannot uninstall blinker...` during `pip install` | A pre-existing system Python package conflict. Solved by installing inside a virtual environment (`venv` or `conda`) instead of system-wide. |
| `pdftoppm not found` warning | poppler isn't installed or not on PATH. Run `conda install -c conda-forge poppler -y`. |
| Browser can't reach `localhost:5000` | The server isn't running, or it's still installing dependencies — check the terminal window for errors. |
| `.bat` file won't launch | Windows blocked a file downloaded from the internet. Right-click → Properties → Unblock. |
| `ANTHROPIC_API_KEY is not set` error | `.env` file is missing or still has the placeholder `your-key-here` — edit it with your real key. |
| "This is a development server" warning in the terminal | Expected and harmless — it's how Flask always runs locally. |

## Project structure

```
newspaper-analyzer/
├── app.py                  # Flask web app (local entrypoint)
├── newspaper_analyzer.py   # Core analysis logic, CSV/Excel export
├── templates/index.html    # Web UI
├── requirements.txt        # Python dependencies
├── .env.example            # API key template (copy to .env)
├── run_mac.command         # Mac/Linux double-click launcher
├── run_windows.bat         # Windows double-click launcher
└── start.bat               # Windows launcher pre-wired to a conda env
```
