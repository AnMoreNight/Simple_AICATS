# Python Version - PM01/PM05 Diagnosis

Simplified Python implementation for Primary Diagnosis (PM01) and Secondary Diagnosis (PM05) with Q1-Q6.

## Features

- **PM01 (Primary Diagnosis)**: Scores Q1-Q6 with PRIMARY/SUB/PROCESS/AES scores
- **PM05 (Secondary Diagnosis)**: Reverse scoring validation and consistency checking
- **Scoring Engine**: Weighted scoring (PRIMARY 60%, SUB 20%, PROCESS 20%)
- **Google Sheets Integration**: Read Q1-Q6, write PM01/PM05 results
- **Local Execution**: Runs locally on Windows 11

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Google Sheets API Credentials

1. Create a service account in Google Cloud Console
2. Download the JSON credentials file
3. Set `GOOGLE_APPLICATION_CREDENTIALS_JSON` environment variable with the JSON content directly (copy the entire contents of the JSON file)

### 3. Spreadsheet Configuration

1. Share your Google Spreadsheet with the service account email
2. Set the `SPREADSHEET_ID` environment variable
3. Create sheets:
   - `Config`: Configuration sheet with key-value pairs (must include `llmApiKey`)
   - `Respondents`: Contains Q1-Q6 answers
   - `Questions`: Question metadata (Q1-Q6 with primary_skill, sub_skill, process_skill)
   - `PM01`: Will be created automatically for PM01 results
   - `PM05`: Will be created automatically for PM05 results

**Config Sheet Format**: Create a `Config` sheet with two columns:
- Column A: Configuration key (e.g., `llmApiKey`, `llmModel`, `llmProvider`, etc.)
- Column B: Configuration value (e.g., your OpenAI API key)

### 4. Environment Variables

Copy `.env_example` to `.env` and fill in:

```env
# Google Sheets API credentials (JSON content)
GOOGLE_APPLICATION_CREDENTIALS_JSON={"type":"service_account",...}

# Google Spreadsheet ID
SPREADSHEET_ID=your_spreadsheet_id_here
```

**Note**: The OpenAI API key should be configured in the `Config` sheet of your Google Spreadsheet (set `llmApiKey` in the Config sheet).

## Usage

```bash
python main.py
```

The script will:
1. Read all respondents from the Respondents sheet
2. Filter for pending respondents (not "pm01完了", "pm05完了", "診断完了")
3. Validate respondents (must have Q1-Q6 answers)
4. Process each respondent:
   - Run PM01 (Primary Diagnosis)
   - Run PM05 (Secondary Diagnosis)
   - Write results to PM01 and PM05 sheets

## PM01 Output Structure

```json
{
  "primary_scores": {"skill_name": <avg_score>},
  "sub_scores": {"skill_name": <avg_score>},
  "process_scores": {"skill_name": <avg_score>},
  "aes_scores": {"Q1": <score>, "Q2": <score>, ...},
  "total_score": <weighted_total>,
  "per_question": {
    "Q1": {
      "primary_score": 0-5,
      "sub_score": 0-5,
      "process_score": 0-5,
      "aes_score": 0-5,
      "comment": "<string>"
    },
    ...
  }
}
```

## PM05 Output Structure

```json
{
  "status": "valid | caution | re-evaluate",
  "consistency_score": 1-5,
  "issues": ["<issue1>", ...],
  "comment": "<string>"
}
```

## Scoring Rules

- **Weighted Total**: PRIMARY 60% + SUB 20% + PROCESS 20%
- **Score Range**: 0-5 for all scores
- **Consistency**: PM05 compares reverse-scored results with PM01
  - 4.5+ = "valid"
  - 3.5-4.4 = "caution"
  - <3.5 = "re-evaluate"

## Error Handling

- Validation errors → `ValidationLog` sheet
- API errors → `ErrorLog` sheet
- Run summary → `RunLog` sheet

## Project Structure

```
├── core/
│   ├── config.py          # Configuration management
│   └── utils.py           # Utility functions
├── services/
│   ├── json_parser.py    # JSON parsing (PM01/PM05)
│   ├── llm.py             # LLM service (PM01/PM05)
│   ├── scoring_engine.py  # Scoring calculations
│   ├── sheets.py          # Google Sheets integration
│   └── validation.py      # Data validation
├── main.py                # Main entry point
├── requirements.txt       # Python dependencies
└── README_PYTHON.md       # This file
```
