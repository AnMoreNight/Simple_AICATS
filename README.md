# Python Version - PM01/PM05 Diagnosis

4-step diagnostic system for Primary Diagnosis (PM01) and Secondary Diagnosis (PM05) with Q1-Q6.

## Features

- **4-Step Diagnostic Flow**: PM01 Raw → PM05 Raw → PM01 Final → PM05 Final
- **Individual Q-A Scoring**: Each question-answer pair is evaluated separately
- **Reverse Logic Validation**: PM05 uses reverse logic to validate PM01 results
- **Scoring Engine**: Weighted scoring (PRIMARY 60%, SUB 20%, PROCESS 20%)
- **LLM-Powered Analysis**: Configurable prompts for each diagnostic step
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
   - `Config`: Configuration sheet with key-value pairs (see Config Sheet Fields below)
   - `Respondents`: Contains Q1-Q6 answers (see Respondents Sheet Format below)
   - `Questions`: Question metadata (Q1-Q6, see Questions Sheet Format below)
   - `PM01`: Will be created automatically for PM01 results
   - `PM05`: Will be created automatically for PM05 results

**Config Sheet Format**: Create a `Config` sheet with two columns:
- Column A: Configuration key
- Column B: Configuration value

**Required Config Sheet Fields**:
- `respondentsSheet`: Sheet name for respondents (default: "Respondents")
- `validationLogSheet`: Sheet name for validation logs (default: "ValidationLog")
- `diagnosisDetailSheet`: Sheet name for diagnosis details
- `errorLogSheet`: Sheet name for error logs (default: "ErrorLog")
- `questionSheet`: Sheet name for questions (default: "Questions")
- `skillcriteriaSheet`: Sheet name for skill criteria
- `typecriteriaSheet`: Sheet name for type criteria
- `defaultTimeZone`: Default timezone (e.g., "Asia/Tokyo")
- `llmProvider`: LLM provider (e.g., "chatgpt")
- `llmApiUrl`: LLM API URL (e.g., "https://api.openai.com/v1/chat/completions")
- `llmModel`: LLM model name (e.g., "gpt-4o")
- `llmApiKey`: OpenAI API key (required)
- `maxRetries`: Maximum retry attempts (e.g., 3)
- `promptPM1Raw`: Prompt for STEP 1 - PM01 Raw Scoring (required)
- `promptPM1Final`: Prompt for STEP 3 - PM01 Final Analysis (required)
- `promptPM5Raw`: Prompt for STEP 2 - PM05 Raw Scoring (required)
- `promptPM5Final`: Prompt for STEP 4 - PM05 Final Consistency Check (required, or use `promptPM5Raw`)

**Respondents Sheet Format**:
- Column 0: No. (Respondent ID)
- Column 1: 作成日 (Creation Date)
- Column 2: お名前/姓 (Family Name)
- Column 3: お名前/名 (Given Name)
- Column 4: 所属部門（部署）名 (Department Name)
- Column 5: 会社名（法人名） (Company Name)
- Column 6: 年齢層 (Age Group)
- Column 7: Q1 Answer
- Column 8: Q1 Reason
- Column 9: Q2 Answer
- Column 10: Q2 Reason
- Column 11: Q3 Answer
- Column 12: Q3 Reason
- Column 13: Q4 Answer
- Column 14: Q4 Reason
- Column 15: Q5 Answer
- Column 16: Q5 Reason
- Column 17: Q6 Answer
- Column 18: Q6 Reason
- Column 19: Status

**Questions Sheet Format**:
- Column A: Question ID (Q1, Q2, Q3, Q4, Q5, Q6)
- Column B: Main question text (multiple choice)
- Column C: Follow-up question text
- Each question spans 2 rows:
  - Row 1: A=Q1, B=main question
  - Row 2: C=follow-up question

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
4. Process each respondent through 4 diagnostic steps:
   - **STEP 1**: PM01 Raw Scoring (Individual Q-A scoring)
   - **STEP 2**: PM05 Raw Scoring (Reverse logic validation)
   - **STEP 3**: PM01 Final (Aggregation + LLM Analysis)
   - **STEP 4**: PM05 Final (Consistency Check)
5. Write results to PM01 and PM05 sheets

## Diagnostic Flow

### STEP 1: PM01 Raw Scoring
- Processes each question (Q1-Q6) individually
- Outputs: `primary_score`, `sub_score`, `process_score`, `aes_clarity`, `aes_logic`, `aes_relevance`, `evidence`, `judgment_reason`
- Uses `promptPM1Raw` from Config sheet

### STEP 2: PM05 Raw Scoring
- Reverse logic scoring using PM01 raw results as reference
- Processes each question individually
- Compares with PM01 raw and notes differences
- Uses `promptPM5Raw` from Config sheet

### STEP 3: PM01 Final
- Aggregates PM01 raw scores by category
- Applies increase/decrease point adjustments
- Calls LLM for analysis (top strengths, weaknesses, summary, AI use level)
- Uses `promptPM1Final` from Config sheet

### STEP 4: PM05 Final
- Overall consistency check using PM01 Final
- Evaluates contradictions, excessive abstraction, superficiality
- Uses `promptPM5Final` (or `promptPM5Raw`) from Config sheet

## PM01 Final Output Structure

```json
{
  "scores_primary": {
    "問題理解": <avg_score>,
    "論理構成": <avg_score>,
    "仮説構築": <avg_score>,
    "AI指示": <avg_score>,
    "AI成果検証力": <avg_score>,
    "優先順位判断": <avg_score>
  },
  "scores_sub": {
    "情報整理": <avg_score>,
    "因果推論": <avg_score>,
    ...
  },
  "process": {
    "clarity": <avg_score>,
    "structure": <avg_score>,
    "hypothesis": <avg_score>,
    "prompt_clarity": <avg_score>,
    "quality_check": <avg_score>,
    "consistency": <avg_score>
  },
  "aes": {
    "Q1": <aes_score>,
    "Q2": <aes_score>,
    ...
  },
  "total_score": <weighted_total>,
  "per_question": {
    "Q1": {
      "primary_score": 0-5,
      "sub_score": 0-5,
      "process_score": 0-5,
      "aes_score": 0-5,
      "aes_clarity": 0-5,
      "aes_logic": 0-5,
      "aes_relevance": 0-5,
      "evidence": "<string>",
      "judgment_reason": "<string>"
    },
    ...
  },
  "top_strengths": [
    {"category": "primary|sub|process", "skill": "<skill_name>", "score": <float>, "reason": "<string>"}
  ],
  "top_weaknesses": [
    {"category": "primary|sub|process", "skill": "<skill_name>", "score": <float>, "reason": "<string>"}
  ],
  "overall_summary": "<comprehensive summary>",
  "ai_use_level": "基礎|標準|高度",
  "recommendations": ["<recommendation1>", ...],
  "debug_raw": {...}
}
```

## PM05 Final Output Structure

```json
{
  "status": "valid | caution | re-evaluate",
  "consistency_score": 1-5,
  "issues": ["<issue1>", "<issue2>", ...],
  "summary": "<overall consistency summary>"
}
```

## Scoring Rules

### Score Calculation
- **Weighted Total**: PRIMARY 60% + SUB 20% + PROCESS 20%
- **Score Range**: 0-5 for all scores
- **AES Score**: (clarity + logic + relevance) / 3 (not included in total score)

### Category Mapping (Q1-Q6)
- **Q1**: PRIMARY=問題理解, SUB=情報整理, PROCESS=clarity
- **Q2**: PRIMARY=論理構成, SUB=因果推論, PROCESS=structure
- **Q3**: PRIMARY=仮説構築, SUB=前提設定, PROCESS=hypothesis
- **Q4**: PRIMARY=AI指示, SUB=要件定義力, PROCESS=prompt_clarity
- **Q5**: PRIMARY=AI成果検証力, SUB=品質チェック力, PROCESS=quality_check
- **Q6**: PRIMARY=優先順位判断, SUB=意思決定, PROCESS=consistency

### Score Levels
- **強い**: 4.0+
- **標準**: 2.6-3.9
- **弱い**: 2.5以下

### Consistency Check (PM05 Final)
- **valid**: consistency_score >= 4.5
- **caution**: consistency_score >= 3.5
- **re-evaluate**: consistency_score < 3.5

## Error Handling

- Validation errors → `ValidationLog` sheet
- API errors → `ErrorLog` sheet
- Run summary → `RunLog` sheet

## Diagnostic Steps Details

### STEP 1: PM01 Raw Scoring
- **Purpose**: Individual Q-A scoring with evidence and judgment reasoning
- **Input**: One question-answer pair at a time
- **Output**: Scores, evidence, judgment reason, increase/decrease points
- **Prompt**: `promptPM1Raw` from Config sheet

### STEP 2: PM05 Raw Scoring
- **Purpose**: Reverse logic validation using PM01 raw as reference
- **Input**: One question-answer pair + PM01 raw result
- **Output**: Reverse-scored results with difference notes
- **Prompt**: `promptPM5Raw` from Config sheet

### STEP 3: PM01 Final
- **Purpose**: Aggregate raw scores and generate insights
- **Process**:
  1. Aggregate PM01 raw scores by category
  2. Apply increase/decrease point adjustments
  3. Calculate weighted total score
  4. Call LLM for analysis (strengths, weaknesses, summary, recommendations)
- **Prompt**: `promptPM1Final` from Config sheet

### STEP 4: PM05 Final
- **Purpose**: Overall consistency check and validation
- **Input**: PM01 Final result
- **Output**: Consistency score, status, issues, summary
- **Prompt**: `promptPM5Final` (or `promptPM5Raw`) from Config sheet

## Project Structure

```
├── core/
│   ├── config.py          # Configuration management (reads from Config sheet)
│   └── utils.py           # Utility functions
├── services/
│   ├── json_parser.py     # JSON parsing (PM01 Raw, PM05 Raw, PM01 Final, PM05 Final)
│   ├── llm.py             # LLM service (4-step diagnostic flow)
│   ├── scoring_engine.py  # Scoring calculations and aggregation
│   ├── sheets.py          # Google Sheets integration
│   └── validation.py      # Data validation
├── main.py                # Main entry point (4-step diagnostic flow)
├── requirements.txt       # Python dependencies
├── .env_example           # Environment variables template
└── README.md              # This file
```

## Important Notes

1. **Prompt Requirements**: All prompts (`promptPM1Raw`, `promptPM1Final`, `promptPM5Raw`, `promptPM5Final`) must be configured in the Config sheet
2. **Diagnostic Order**: The 4-step flow must be executed in order (PM01 raw → PM05 raw → PM01 Final → PM05 Final)
3. **Individual Q-A Processing**: Each question is processed separately in STEP 1 and STEP 2
4. **No Re-reading**: STEP 3 aggregates existing raw scores without re-reading Q-A pairs
5. **Consistency Validation**: STEP 4 validates overall consistency of the final diagnosis
