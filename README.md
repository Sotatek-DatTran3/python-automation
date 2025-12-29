> [!WARNING]
> This tool just use for specific website. In briefly, it track html and interact with it. If you want to use in another website, identify which html tag you want to interact with and change the code logic

# Python Automation Tool

A  CLI tool automates the process of testing chatbot queries using data stored in Excel files.
It reads test prompts (queries) from a specified range in an Excel sheet, sends them to a chatbot for processing, retrieves the chatbot’s responses and reference documents, and then writes the results back into an output sheet or specific cells in the same file.

Example usage:
```sh
.\python-automation-cli.exe read -d "C:\Users\ddat8\Documents\Work\ProcessingServiceEvaluation.xlsx" -s "テストデータ"  -f C3 -t C30 -o E3 J3
```

| Short | Long                  | Required | Description                                                                              | Example                |
| ----- | --------------------- | -------- | ---------------------------------------------------------------------------------------- | ---------------------- |
| `-d`  | `--file-path`         | ✅ Yes    | Path to the Excel file you want to read.                                                 | `-d ./data/input.xlsx` |
| `-s`  | `--sheet-name`        | ❌ No     | Name of the sheet to read from. If not provided, defaults to the active sheet.           | `-s Sheet1`            |
| `-f`  | `--from`              | ✅ Yes    | Starting cell address (e.g., A1).                                                        | `-f A1`                |
| `-t`  | `--to`                | ✅ Yes    | Ending cell address (e.g., D10).                                                         | `-t D10`               |
| `-o`  | `--output`            | ✅ Yes     | Two output start cells (space-separated). First cell for write chatbot response, second cell for write references.    | `-o A1 B1`             |
| `-os` | `--output-sheet-name` | ❌ No     | Sheet name where to write the output. If not provided, defaults to the active sheet. | `-os OutputSheet`      |
| `-`    | `--headless`         | ❌ No     | Run the browser in headless mode (no visible window). Useful for running on servers or in the background. | `--headless`           |


## Prerequisites

- Python 3.11+
- [uv](https://github.com/astral-sh/uv) package manager
- Microsoft Edge browser (for Selenium WebDriver)

## Getting Started

### 1. Clone the Repository

```sh
git clone <repository-url>
cd python-automation
```

### 2. Install Dependencies

```sh
uv sync
```

### 3. Configure Environment Variables

Copy the example environment file and edit with your credentials:

```sh
copy .env.example .env
```

Edit `.env` file:
```ini
log_mail=your.email@example.com
log_pass=YourPassword
host=http://your-host-url/
```

### 4. Run the CLI Tool

```sh
uv run python cli/script_cli.py --help
```

Example usage:
```sh
uv run python cli/script_cli.py read --file-path data.xlsx --from A1 --to A10 --output B2 C2 --sheet-name Sheet1
uv run python cli/script_cli.py read --file-path data.xlsx --from A1 --to A10 --output B2 C2 --sheet-name Sheet1 --headless
```

## Building Executable

### 1. Install PyInstaller

```sh
uv add pyinstaller
```

### 2. Build the .exe File

```sh
uv run pyinstaller --onefile --add-data ".env.example;." --name python-automation-cli cli/script_cli.py
```

### 3. Locate the Executable

The built executable will be in:
```
dist/python-automation-cli.exe
```

### 4. Prepare Distribution Package

Create a distribution folder with:
- `dist/python-automation-cli.exe`
- `.env.example`
- `README_DISTRIBUTION.md` (instructions for end users)

**⚠️ Important:** Never include your actual `.env` file in the distribution!

## Project Structure

```
python-automation/
├── cli/
│   ├── lib/
│   │   ├── handle_excel.py
│   │   ├── handle_query.py
│   │   └── script.py
│   └── script_cli.py
├── .env.example              # Environment variables template
├── .env                      # Your credentials (git-ignored)
├── .gitignore
├── pyproject.toml
└── README.md
```

## Distribution

When sharing the executable with others:

1. Build the `.exe` using the steps above
2. Create a distribution package with:
   - `python-automation-cli.exe`
   - `.env.example`
   - `README_DISTRIBUTION.md`
3. Users will need to:
   - Copy `.env.example` → `.env`
   - Edit `.env` with their credentials
   - Run the executable

## Troubleshooting

### .env file not found
- Ensure `.env` is in the same directory as the executable
- Check that the file is named exactly `.env` (no `.txt` extension)

### Selenium errors
- Ensure Microsoft Edge browser is installed
- Update Edge WebDriver if needed
