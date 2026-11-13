# Engineering Data Tools — Streamlit

One Streamlit website with two tabs:
- Tab 1: APC1 → OIT Converter (based on supplied APC1 v8 script)
- Tab 2: QA PSV Compare (based on supplied QApsv_compare_list_v4.py)

## Files
```text
APC1_OIT_QA_Streamlit/
├── app.py
├── converter.py
├── qa_compare.py
├── requirements.txt
├── README_Streamlit_部署教學.md
├── .gitignore
└── .streamlit/config.toml
```

## Local run
Python 3.10+ recommended:
```bash
python -m venv .venv
```
Windows PowerShell:
```powershell
.venv\Scripts\Activate.ps1
```
Then:
```bash
pip install -r requirements.txt
streamlit run app.py
```
Open `http://localhost:8501`.

## GitHub + Streamlit deployment
1. Create a GitHub repository. For source-code privacy, choose Private.
2. Upload all project files/folders to repository root (including `.streamlit/config.toml`).
3. Go to https://share.streamlit.io/ and sign in with GitHub.
4. Authorize Streamlit to access the repository, including private repositories if required.
5. Click Create app; choose repository, branch `main`, and main file `app.py`.
6. Deploy. Streamlit installs packages from `requirements.txt`.
7. If the app is Public, anyone with the URL may access it. Do not upload confidential company files unless your IT/InfoSec has approved the hosting and data handling.

## Tab 1 behavior
- APC1 `.xlsx` / `.xlsm` upload.
- Finds worksheet by the same required headers as v8.
- Preserves the embedded OIT2 template and v8 mapping/formatting.
- Downloads `MMDD_OIT_result_PTxxxx.xlsx` (Singapore local date).

## Tab 2 behavior
- Upload OIT1 `.xlsx` / `.xlsm`.
- Upload webDM2 raw `.psv`, PSV-export `.xlsx`/`.xlsm`, or normal Excel Device Listing.
- Compares Feature ID to Device Name, with original normalization and maturity/restriction rules.
- Downloads an Excel report with Summary, Differences, All_Comparisons, Unmatched, Read_Me sheets.

## Notes
- Original file selection prompts, local folder scanning, saving beside source files, Tkinter popups, and opening the output file are replaced with browser upload/download.
- The comparison/conversion business logic is retained.
- Validate both tabs against the original local scripts using representative, non-confidential test files before relying on the hosted version.
- This app processes uploads in memory and does not intentionally persist uploaded files. This is not a compliance guarantee; platform operations/logging are outside this code's control.
