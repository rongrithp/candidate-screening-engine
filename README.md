# Candidate Screening Engine 🚀
> **Commercial-Grade Automated HR Candidate Screening & Resume Evaluation Engine**  
> Powered by Google Gemini GenAI (`gemini-2.5-flash`) & HR Natural Break / Gap Analysis Tiering

---

## 📌 1. System Overview

The **Candidate Screening Engine** is an automated recruitment assistant designed specifically for HR teams, Talent Acquisition specialists, and Hiring Managers. 

It reads candidate resumes (CVs) and Job Descriptions (JDs) in **PDF** (both standard text PDFs and scanned image PDFs), **Microsoft Word (.docx)**, and **Text (.txt)** formats, evaluates each candidate against key position requirements, applies **Harrow HR Natural Break & Gap Analysis Tiering**, and exports polished executive reports ready for HR decision-making.

### Key Capabilities
- **Multi-Format Document Parsing**: Direct processing of `.pdf`, `.docx`, `.doc`, `.txt`, and `.md` files.
- **Scanned PDF & OCR Fallback**: Native multimodal document analysis via Google Gemini 2.5 Flash for scanned or image-only PDFs.
- **Automated HR Tiering**: Classifies candidates into **🟢 Shortlist**, **🟡 Longlist**, and **🔴 Filtered Out / Rejected** based on quantitative fit scores (0–100) and gap analysis.
- **Triple Executive Reports**: Generates Microsoft Word (`.docx`), Microsoft Excel (`.xlsx`), and Markdown (`.md`) summary reports automatically.

---

## ⚡ 2. Quick 3-Step User Guide (For Non-Technical Users)

Running candidate screening requires **zero coding knowledge**. Follow these three simple steps:

```
workspace/
├── _JOB_TEMPLATE/               <-- Copy this folder for new jobs
│   ├── jd/                      <-- Drop Job Description here
│   └── cv/                      <-- Drop Candidate CVs here
└── run_screening.bat            <-- Double-click to execute
```

### Step 1: Copy `_JOB_TEMPLATE` Folder
- In File Explorer, right-click the `_JOB_TEMPLATE` folder.
- Select **Copy**, then **Paste** in the same workspace directory.
- Rename the newly created folder to match your target job position (for example: `Job-03_Head of Science`).

### Step 2: Place Job Description & Resumes
- Open your new job folder (e.g., `Job-03_Head of Science`).
- Open the `jd/` subfolder and place your **Job Description** file inside (`.pdf`, `.docx`, or `.txt`).
- Open the `cv/` subfolder and place all **Candidate Resumes / CVs** inside (`.pdf`, `.docx`, or `.txt`).

### Step 3: Double-Click `run_screening.bat`
- Inside your job folder, double-click the `run_screening.bat` file.
- A command window will open and perform candidate screening automatically.
- Once finished, all 3 formatted summary reports will be generated directly inside your job folder!

---

## 📊 3. Output Reports Explanation

Once screening is complete, 3 complementary executive report formats are generated:

| File Name | Format | Best Used For |
| :--- | :---: | :--- |
| **`Candidate_Summary.docx`** | Microsoft Word | **Executive & Panel Reviews**: Formatted in 1-page A4 layout following Harrow Brand Identity (Deep Navy & Gold styling). Contains fit scores, candidate strengths, skill gaps, executive recommendations, and clickable links to candidate CVs. Ideal for sharing directly with Head of Department / Panel Interviewers. |
| **`Candidate_Summary.xlsx`** | Microsoft Excel | **HR Candidate Tracking & Data Filtering**: Color-coded HR tiers (🟢 Shortlist, 🟡 Longlist, 🔴 Filtered Out), interactive hyperlink formulas for opening candidate CVs, candidate ranking, fit scores, and side-by-side experience/education breakdowns. Ideal for sorting and tracking. |
| **`Candidate_Summary.md`** | Markdown | **Quick Screen Preview & Documentation**: Lightweight markdown summary file for instant preview in text editors, web portals, or internal documentation systems. |

---

## 🔑 4. API Key Guide (Google Gemini AI)

### What is an API Key?
An **API Key** is a secure digital passphrase that allows the Candidate Screening Engine to connect to Google's Gemini AI service to analyze and evaluate candidate resumes.

### Step-by-step: How to get a Free Google Gemini API Key
1. Open your web browser and navigate to **[Google AI Studio](https://aistudio.google.com/app/apikey)** (`https://aistudio.google.com/app/apikey`).
2. Sign in using your Google / Gmail account.
3. Click the blue button labeled **"Create API key"** (or **"Get API key"**).
4. Choose **"Create API key in new project"**.
5. Copy your newly generated key string (a long string starting with `AIzaSy...`).

### How to enter your API Key on First Run
- The first time you double-click `run_screening.bat`, if no key is saved yet, the command window will prompt:
  ```text
  ============================================================
  🔑 GEMINI_API_KEY / GOOGLE_API_KEY not found in environment or .env file.
  ============================================================
  Please enter your GEMINI_API_KEY:
  ```
- **Paste your key**: Right-click inside the window or press `Ctrl + V`, then press `Enter`.
- **Automatic Self-Healing**: The engine automatically saves your key to a `.env` file in the project folder so you never have to enter it again!

---

## 🛡️ License & Safeguarding Notice

*Harrow International School Bangkok is committed to the safety and protection of children. Candidate data evaluated by this tool must be handled in accordance with institutional data protection and recruitment policies.*
