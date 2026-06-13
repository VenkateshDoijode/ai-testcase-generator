<div align="center">

# 🧪 AITestCaseGen

**AI-powered test case management toolkit for Jira Zephyr Scale**

Generate, import, update, link, and report on test cases — fully automated, with or without cloud AI.

[![Python](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Zephyr Scale](https://img.shields.io/badge/Zephyr-Scale-orange.svg)](https://www.atlassian.com/software/jira/zephyr)
[![HuggingFace](https://img.shields.io/badge/AI-HuggingFace%20%7C%20Ollama-yellow.svg)](https://huggingface.co/)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](#contributing)

</div>

---

## 🚀 Why AITestCaseGen?

Writing test cases manually from requirements is one of the biggest time sinks in QA. **AITestCaseGen** reads your requirements — from Jira, Confluence, documents, or plain text — and generates structured positive, negative, and edge-case test cases, then pushes them directly into Zephyr Scale.

It's not just a generator. It's a **full test case lifecycle CLI**:

| | |
|---|---|
| 🧠 **Generate** | Create test cases from Jira issues, Confluence pages, docs, or text |
| 📥 **Import** | Push generated cases directly into Zephyr Scale |
| ✏️ **Update** | Bulk-edit existing test cases |
| 🔗 **Link** | Connect test cases to Jira issues |
| 🗑️ **Delete** | Remove test cases (with dry-run safety) |
| 📤 **Export** | Pull existing cases out of Zephyr |
| 🔄 **Test Cycles** | Update execution results in bulk |

---

## 🤖 AI Engine Priority

The script automatically picks the best available engine — **zero config required**:

| Priority | Engine | Condition |
|---|---|---|
| 1 | **HuggingFace** | `transformers` installed + model downloadable |
| 2 | **Ollama** | Ollama installed and running locally |
| 3 | **Rule-based** | Always works — no AI needed |

This means the tool **never hard-fails** — even with no internet, no API key, and no GPU.

---

## ✨ Why This Is Different

Most "AI test case generator" tools are tightly coupled to a single cloud provider — if your API key expires or your org blocks external LLM calls, the tool stops working entirely. AITestCaseGen is built around **graceful degradation**:

- **No hard dependency on a paid API** — HuggingFace models run via your Artifactory mirror or locally, Ollama runs fully offline, and the rule-based engine guarantees the script keeps working even with zero AI infrastructure.
- **Enterprise-friendly by default** — designed for locked-down environments with no access to OpenAI/Anthropic APIs; works fully air-gapped.
- **End-to-end lifecycle, not just generation** — generates *and* imports, updates, deletes, links, exports, and updates test cycle executions directly via the Zephyr Scale REST API.
- **Multi-source input** — Jira issues, Confluence pages, document folders (`.docx` / `.pdf` / `.txt`), or raw text all feed the same pipeline.

---

## 📊 Comparison with Similar Approaches

| Capability | AITestCaseGen | ChatGPT / Copy-Paste | Commercial Zephyr AI Add-ons |
|---|---|---|---|
| Works offline / air-gapped | ✅ Ollama or local model | ❌ Cloud API required | ❌ Usually cloud-only |
| No paid API key required | ✅ Rule-based fallback | ❌ Subscription needed | ❌ License cost |
| Direct Zephyr Scale integration | ✅ Full API pipeline | ❌ Manual copy-paste | ✅ Vendor UI only |
| Multi-source input | ✅ Jira, Confluence, docs, text | ⚠️ One source at a time | ⚠️ Often limited |
| Bulk test cycle updates | ✅ CSV/Excel bulk update | ❌ Not supported | ⚠️ Vendor-dependent |
| Open source & customizable | ✅ Edit anything | N/A | ❌ Closed source |
| CI/CD friendly | ✅ Pure CLI | ❌ Manual workflow | ⚠️ Varies |
| Cost | **Free** | Subscription | License/subscription |

> **The core differentiator:** AI generation is just one stage of a connected pipeline — generation, Jira/Zephyr sync, and execution reporting all flow together instead of ending at "here's some text, paste it yourself."

---

## 📁 Project Structure

```
AITestCaseGen/
├── resources/
│   ├── config.ini                    # Jira credentials and settings
│   ├── import_test_case.xlsx         # Input for import / output of AI generator
│   ├── update_test_case.xlsx         # Input for update script
│   ├── delete_Test_case.xlsx         # Input for delete script
│   └── update_test_cycle.xlsx        # Input for test cycle update
├── src/
│   ├── main/
│   │   ├── config.py                 # Reads config.ini
│   │   └── jira_client.py            # Jira / Zephyr Scale REST API client
│   └── test/
│       ├── generate_tc_via_AI.py     # AI test case generator
│       ├── import_testcases.py       # Import test cases to Zephyr Scale
│       ├── update_testcases.py       # Update existing test cases
│       ├── delete_testcases.py       # Delete test cases
│       ├── link_testcases_to_jira.py # Link test cases to Jira issues
│       ├── export_test_cases.py      # Export test cases from Zephyr
│       ├── project_info.py           # Display Jira project information
│       ├── update_test_cycle.py      # Update test cycle execution results
│       └── create_excel_templates.py # Generate blank Excel input templates
└── requirements.txt
```

---

## ⚙️ Prerequisites

- **Python 3.10+**
- Install dependencies:

```powershell
pip3 install -r requirements.txt
pip3 install transformers accelerate huggingface_hub
```

---

## 🔧 Configuration

Edit `resources/config.ini`:

```ini
[jira]
base_url   = https://your-jira-instance.com
jira_token = YOUR_JIRA_PAT_TOKEN
project_key = YOUR_PROJECT_KEY
project_id  = YOUR_PROJECT_ID

[testcase]
owner_account_id = YOUR_ACCOUNT_ID
test_cycle       = YOUR_CYCLE_KEY

[ai]
hf_token    = YOUR_HUGGINGFACE_API_TOKEN
hf_endpoint = https://your-artifactory-host/artifactory/api/huggingfaceml/huggingface

[defaults]
type_of_test      = Regression Test
automation_status = Manual
```

> 🔒 `config.ini` is in `.gitignore` — credentials are never committed to git.

---

## 🧩 Setup — HuggingFace (Primary)

### Step 1 — Install dependencies

```powershell
pip3 install transformers accelerate huggingface_hub
```

### Step 2 — Configure HuggingFace Token

1. Get your HuggingFace or Artifactory identity token
2. Add it to `resources/config.ini`:

```ini
[ai]
hf_token = your-artifactory-identity-token
```

### Step 3 — First Run (model download ~300MB)

The model downloads automatically on first run from your configured `hf_endpoint`.

**Default model:** `google/flan-t5-base`

| Model | Size | Speed |
|---|---|---|
| `google/flan-t5-small` | ~80MB | Fastest |
| `google/flan-t5-base` | ~300MB | Default |
| `google/flan-t5-large` | ~800MB | Most accurate |

---

## 🦙 Setup — Ollama (Optional Fallback)

If HuggingFace is unavailable, install Ollama for fully local AI generation:

1. Download: **https://ollama.com/download/windows**
2. Run `OllamaSetup.exe`
3. Pull a model:

```powershell
ollama pull llama3.2
```

No code changes needed — the script detects Ollama automatically.

---

## 📝 Generating Test Cases

### From a folder of documents (`.docx` / `.pdf` / `.txt`)

```powershell
python -m test.generate_tc_via_AI --input-folder "C:\path\to\docs" --folder "/Generated" --count 10
```

### From a Jira issue key

```powershell
python -m test.generate_tc_via_AI --issue PROJECT-1234 --folder "/Generated" --count 5
```

### From a Jira issue + documents folder (combined)

```powershell
python -m test.generate_tc_via_AI --issue PROJECT-1234 --input-folder "C:\path\to\docs" --folder "/Generated" --count 10
```

### From a single file

```powershell
python -m test.generate_tc_via_AI --file resources/requirements.txt --folder "/Generated"
```

### From a Confluence page

```powershell
python -m test.generate_tc_via_AI --confluence 123456789 --folder "/Generated" --count 10
```

Pass the **page ID** (numeric) or the **full page URL**:

```powershell
python -m test.generate_tc_via_AI --confluence "https://your-jira.com/wiki/spaces/PROJECT/pages/123456789" --count 10
```

Multiple pages (comma-separated):

```powershell
python -m test.generate_tc_via_AI --confluence "123456789,987654321" --folder "/Generated"
```

> **How to get the Confluence page ID:**
> Open the page → click `...` (More actions) → **Page Information** → ID is in the URL: `.../pages/<pageId>/...`

### From inline text

```powershell
python -m test.generate_tc_via_AI --text "User should be able to login with valid credentials" --count 5
```

### All CLI Options

| Argument | Description | Default |
|---|---|---|
| `--issue` | Jira issue key (e.g. `PROJECT-1234`) | — |
| `--input-folder` | Folder containing `.docx` / `.pdf` / `.txt` files | — |
| `--file` | Single requirements file path | — |
| `--text` | Inline requirement text | — |
| `--confluence` | Confluence page ID or URL (comma-separated) | — |
| `--model` | HuggingFace or Ollama model name | `google/flan-t5-base` |
| `--count` | Number of test cases to generate per source | `5` |
| `--folder` | Zephyr Scale folder path | `/Generated` |
| `--output` | Output `.xlsx` file path | `resources/import_test_case.xlsx` |

---

## 📤 Import Generated Test Cases to Jira

```powershell
python -m test.import_testcases --file ..\resources\import_test_case.xlsx
```

---

## 🛠️ Other Scripts

### Update existing test cases

```powershell
python -m test.update_testcases --file ..\resources\update_test_case.xlsx
```

### Delete test cases — dry run first, then live

```powershell
python -m test.delete_testcases --file ..\resources\delete_Test_case.xlsx --dry-run
python -m test.delete_testcases --file ..\resources\delete_Test_case.xlsx
```

### Link test cases to a Jira issue

```powershell
python -m test.link_testcases_to_jira --file ..\resources\update_test_case.xlsx
```

### Update test cycle execution results

```powershell
python -m test.update_test_cycle --file ..\resources\update_test_cycle.xlsx
```

## 🤝 Contributing

Issues, feature requests, and PRs are welcome. If AITestCaseGen saves you time, consider giving it a ⭐ — it helps others discover the project.

## 📜 License

MIT License — see [LICENSE](LICENSE) for details.
