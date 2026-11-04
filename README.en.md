# md2gost

[Русский](README.md) | **English**

A desktop app that turns Markdown into a Word document (`.docx`) formatted
according to **GOST 7.32-2017**, the Russian/CIS standard for research
reports. Write plain Markdown and get a finished report with the correct
margins, fonts, numbered sections, figures, tables and equations, a table of
contents, a bibliography and a title page.

It is built on [pandoc](https://pandoc.org) and
[pandoc-crossref](https://github.com/lierdakil/pandoc-crossref); the app
downloads and sets them up for you on first launch.

### Requirements

| Component | Version | Notes |
|---|---|---|
| Python | **3.10+** | Older versions will not start the app |
| Tkinter | bundled with Python | On Linux install it separately: `sudo apt install python3-tk` |
| pandoc | 3.9+ | Installed automatically into `bin/` |
| pandoc-crossref | matching your pandoc | Installed automatically into `bin/` |
| Internet | first launch only | Needed to download pandoc from GitHub |


### Installation

#### 1. Get the code

```bash
git clone https://github.com/<user>/md2gost.git
cd md2gost
```

Or download the repository as a ZIP and extract it.

#### 2. Create a virtual environment and install dependencies

**Windows (PowerShell / CMD):**

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

**macOS / Linux:**

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

#### 3. Run

```bash
python main.py
```

On first launch the app checks for pandoc and pandoc-crossref. If they are
missing or incompatible, it downloads a matching pair from GitHub into `bin/`
(about 40 MB). Progress is shown in the status bar at the bottom of the
window. Your system-wide pandoc is never modified.

#### Installing pandoc manually (optional)

You can run the installer on its own from a terminal:

```bash
python tools/install_pandoc.py            # interactive mode
python tools/install_pandoc.py --auto     # no prompts
python tools/install_pandoc.py --check    # verify the install (exit code 0 = OK)
python tools/install_pandoc.py --list     # available versions and compatibility
python tools/install_pandoc.py --upgrade  # install the latest compatible pair
python tools/install_pandoc.py --pandoc 3.9  # a specific pandoc version
```

If GitHub reports a rate limit, set the `GITHUB_TOKEN` environment variable to
a personal GitHub token.

### Usage

1. **Markdown (.md)** — choose the file with your report text.
2. **.bib файл** — optional BibTeX bibliography.
3. **Титульный лист** (title page) — pick a template from the list, or click
   the add button and choose your own `.docx`; it is copied into
   `resources/titles/` and stays available next time. "Не использовать"
   means no title page.
4. **Стили тела** (body styles) — the style profile (default: GOST 7.32-2017).
5. Checkboxes:
   - **Авто-оглавление (TOC)** — insert a table of contents;
   - **Нумерация разделов** — number sections `1`, `1.1`, `1.1.1`.
6. Click **"КОНВЕРТИРОВАТЬ"** (Convert) and choose where to save the `.docx`.
7. Watch the log at the bottom of the window for warnings (a figure without a
   caption, a `.bib` with no citations, and so on). When it finishes, the app
   offers to open the result.

**"Отмена"** (Cancel) stops the conversion immediately, including pandoc and
diagram rendering.

> When you first open the document, Word may offer to update fields — accept.
> If the table of contents still has no page numbers, press `Ctrl+A`, then `F9`.

### Using it with an AI assistant

1. Give the assistant
   `resources/styles/GOST_7.32-2017/GOST_7.32-2017.ai.md` as instructions
   (plus `resources/prompts/diagram.md` for diagrams).
2. Ask it to write the report on your topic.
3. Save the answer as `.md` and convert it with md2gost.

### Custom style profiles

Each profile is a folder in `resources/styles/` with a `config.cfg` and a
reference `.docx`. You can also add a Lua filter, pandoc-crossref settings, a
CSL style and extra pandoc arguments. See
[`resources/styles/README.md`](resources/styles/README.md) (in Russian) for
details.

### License

This project is licensed under the [MIT License](LICENSE).

pandoc and pandoc-crossref are licensed under the GPL; the installer saves
their licenses and source links in `bin/licenses/`.
