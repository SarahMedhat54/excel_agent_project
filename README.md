# Excel Assistant

An AI agent that creates, edits, analyzes and charts Excel files from plain-language requests.

Type what you want ("Create a sales sheet with Month, Product, Quantity, Price", "Rename the column Product to Item", "Create a bar chart showing Total for each Month") and download the resulting `.xlsx`.

## What it can do

- **Create** a new workbook from a description, including sample data
- **Edit** an existing file: add, rename or delete columns, add or delete rows, set cells, add formula columns
- **Analyze** the data (EDA): shape, column types, missing values, duplicate rows, summary stats, correlations, top categories
- **Chart** the data with native Excel bar, line and pie charts, saved inside the workbook
- **Upload** your own `.xlsx` and keep working on it across several requests

## How it works

```
Streamlit UI  ->  Agent (Groq LLM)  ->  MCP server (Excel tools)  ->  .xlsx file
```

The model never touches the file directly. It only decides which tool to call; the tools do the editing with `openpyxl`.

| File | Role |
| --- | --- |
| `server.py` | MCP server exposing the Excel tools |
| `agent.py` | Agent loop: Groq chat completions with tool calling over an MCP stdio session |
| `streamlit_app.py` | Main web UI: upload, request box, table preview, download button |
| `app.py` | Minimal Flask UI, an alternative to Streamlit |
| `chat.py` | Command-line chat that drives the same tools with no LLM |
| `test_server.py` | Quick script that creates a file and adds three charts |

### MCP tools

| Tool | Purpose |
| --- | --- |
| `create_excel` | Create a new workbook from columns and rows |
| `describe_excel` | Return sheets, headers, size and the first rows |
| `edit_excel` | Apply a list of edit operations |
| `run_eda` | Run exploratory data analysis |
| `create_chart` | Add a bar, line or pie chart |

## Setup

Requires Python 3.10+ and a [Groq API key](https://console.groq.com/keys).

```bash
git clone https://github.com/SarahMedhat54/excel_agent_project.git
cd excel_agent_project
python -m venv venv
venv\Scripts\activate        # macOS/Linux: source venv/bin/activate
pip install -r requirements.txt
```

Copy `.env.example` to `.env` and add your key:

```
GROQ_API_KEY=your_key_here
MODEL=openai/gpt-oss-120b
```

`MODEL` is optional and defaults to `openai/gpt-oss-120b`.

## Run

Streamlit UI:

```bash
streamlit run streamlit_app.py
```

Flask UI, served on http://localhost:5000:

```bash
python app.py
```

Command-line chat, no API key needed:

```bash
python chat.py
```

### Docker

```bash
docker build -f dockerfile -t excel-assistant .
docker run -p 8501:8501 --env-file .env excel-assistant
```

Then open http://localhost:8501.

## Example requests

```
Create a sales sheet with these columns: Month, Product, Quantity, Price. Add 8 sample data rows covering 4 months (Jan to Apr) and 2 products (Laptop and Phone).
Add a new row: May, Laptop, 5, 20000.
Rename the column Product to Item.
Add a Total column that multiplies Quantity by Price.
Create a bar chart showing Total for each Month.
Analyze the data: are there any missing values or duplicate rows? Give me the mean, min and max for Quantity, Price and Total, and tell me if Quantity and Total are correlated.
```

Charts are native Excel charts, so open the downloaded file in Excel to see them.

## Notes

- Generated and uploaded workbooks are stored in `files/`, each under a short `file_id`.
- Only `.xlsx` files are supported.
- Formula cells are written as real Excel formulas. For analysis, simple arithmetic formulas are evaluated in Python; other formulas are skipped.

## License

[MIT](LICENSE)
