# BTech CO-PO Attainment & Academic Analytics System

Streamlit web application for **Outcome-Based Education (OBE)** analytics aligned with **NBA** documentation.  
**Theory mode** follows the **BTech pattern**: five **CCE** internal assessments (10 marks each), **End Semester** university exam (50 marks), and **40% / 60%** internal–university weighting for final CO attainment.

---

## Features

- BTech theory: **CCE 1–CCE 5** + **End Semester**; configurable max marks and weightages (defaults 10 / 10 / … / 50 and 40:60)
- Default outcome structure: **5 COs**, **11 POs**, and **3 PSOs**
- **Practical mode** (optional): internal + configurable external components — unchanged workflow
- CO internal, university, and final attainment; **PO** and **PSO** contribution tables
- **Discrimination Index** with configurable top %, bottom %, and threshold
- Interactive charts (Plotly + Matplotlib in PDF); **no radar chart**
- **Excel** and **A4 PDF** exports with BTech-oriented cover text

---

## Tech stack

| Area | Library |
|------|---------|
| UI | Streamlit |
| Data | Pandas, NumPy |
| Charts | Plotly, Matplotlib |
| Reports | OpenPyXL, ReportLab |

---

## Project layout

```
.
├── app.py                  # Main application
├── requirements.txt        # Python dependencies
├── README.md
├── .streamlit/config.toml  # Streamlit theme / server defaults
├── run.bat / run.sh        # Local run helpers
└── .gitignore
```

---

## Setup & run

**Python:** 3.10 or 3.11 recommended.

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open the URL shown in the terminal (typically `http://localhost:8501`).

---

## Theory mode — Excel columns

| Column | Typical max (default) |
|--------|------------------------|
| Student Name | — |
| CCE 1 … CCE 5 | 10 each |
| End Semester | 50 |

Use **Download Sample Template** in the app for a correctly formatted file. Values above the configured maximum are **clipped** when you click Calculate.

---

## Deploy on Streamlit Community Cloud

1. Push this repository to GitHub (ensure `requirements.txt` and `app.py` are present).
2. Sign in at [Streamlit Community Cloud](https://share.streamlit.io/).
3. **New app** → select the repo → main file: **`app.py`** → Deploy.

No radar chart is included; dependencies match `requirements.txt`.

---

## Troubleshooting

| Issue | Suggestion |
|-------|------------|
| `ModuleNotFoundError` | Re-run `pip install -r requirements.txt` in the same environment used for Streamlit. |
| Upload errors | Column names must match exactly (**CCE 1**, …, **End Semester**). |
| Empty results | Upload marks and click **Calculate CO-PO Attainment**. |

---

## License

MIT License — use and modify with attribution.

---

## Author

**Harshvardhan Bharat Mali**

- LinkedIn: https://www.linkedin.com/in/harshvardhanmali910  
- GitHub: https://github.com/harsh9109  
