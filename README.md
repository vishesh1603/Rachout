# ReachOut App

ReachOut is a self-contained local web application that automates personalized recruiter outreach. It evaluates your current resume against a job description, calculates a baseline ATS compatibility score with gap analysis, tailors an ATS-optimized 1-page PDF using ReportLab with full editing capabilities and 1-click PDF download, and sends the outreach email via Gmail SMTP—all using free tiers and modern tools.

## Tech Stack
- **Backend:** Python + Flask
- **LLM SDK:** `google-genai` (Official modern Google GenAI SDK; model: `gemini-2.5-flash` / `gemini-2.0-flash` / `gemini-1.5-flash`)
- **Universal Input Ingestion:**
  - **PDF Documents:** Direct multimodal ingestion + `pypdf` text extraction
  - **Word Documents:** `python-docx` + `zipfile` fallback (.docx, .doc)
  - **Images / Screenshots:** Native Gemini Vision AI (JPG, PNG, WebP)
  - **Plain Text / Markdown:** Automatic multi-encoding text decoding (.txt, .md, .rtf)
- **PDF Engine:** ReportLab (free library)
- **Company Research:** DuckDuckGo Search (using `duckduckgo-search` library, free, no API key required)
- **Email Dispatch:** Gmail SMTP SSL using App Passwords (free)
- **Frontend:** Single HTML page with vanilla CSS & JS, featuring drag-and-drop file upload, ATS score gauge, inline resume editor, and live PDF preview

---

## Getting Started

### 1. Prerequisites & API Keys
Before running the application, configure your local environment or `.env` file:

- **Gemini API Key (Free Tier Available):**
  1. Visit [Google AI Studio](https://aistudio.google.com).
  2. Create a free API Key.
  3. Create a `.env` file in the project directory:
     ```env
     GEMINI_API_KEY=AIzaSy...
     # Optional: GEMINI_MODEL=gemini-2.5-flash
     ```
     Or expose it in PowerShell:
     ```powershell
     $env:GEMINI_API_KEY="AIzaSy..."
     ```

- **Gmail App Password (Free):**
  To send emails via Gmail SMTP, you must configure a secure 16-character App Password (not your primary Google password).
  1. Go to your [Google Account Dashboard](https://myaccount.google.com).
  2. Navigate to **Security** > **2-Step Verification** (must be enabled).
  3. Scroll to the bottom and select **App Passwords**.
  4. Create an app password (e.g. choose "Mail" and select your device).
  5. Copy the 16-character passcode shown.

---

### 2. Installation & Running

1. **Install python packages:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the server:**
   Double-click `run.bat` or run:
   ```bash
   python app.py
   ```
   *(Or using Anaconda: `C:\Users\DELL\anaconda3\python.exe app.py`)*

3. Open `http://localhost:5000` in your web browser.

---

## Universal Inputs Supported

You can provide both your **Resume** and the **Target Job Description** in any format:
- **Pasted Text:** Type or paste raw text directly into the text areas.
- **PDF Files (`.pdf`):** Drag and drop or browse. Automatically extracts text and passes native PDF binary to Gemini.
- **Word Documents (`.docx`, `.doc`):** Drag and drop Word resumes or job specifications; extracts full paragraphs and tabular data.
- **Images / Screenshots (`.jpg`, `.jpeg`, `.png`, `.webp`):** Take a screenshot of a job posting or upload an image resume; Gemini's multimodal vision model analyzes all text, layouts, and graphics directly.
- **Text Files (`.txt`, `.md`, `.rtf`):** Drag and drop plain text or markdown files.

---

## The 4-Stage End-to-End Workflow

```
[1. Import Profile & JD] ➔ [2. Current ATS Audit & Insights] ➔ [3. Tailored Resume & PDF Download] ➔ [4. Email Dispatch]
```

1. **Stage 1: Import Profile & Inputs**
   - Provide your name, target company, job role, recruiter email.
   - Enter or upload your current resume and job description (PDF, Word, image, or text).
   - Click **"Analyze ATS Match & Gaps"** (or use **"Load Sample Job & Profile"** for instant demonstration).

2. **Stage 2: Current ATS Compatibility Audit & Gap Analysis**
   - Displays a live **ATS Compatibility Gauge** (0–100 score).
   - Metric Breakdown: Skills Match %, Experience Relevance %, and Formatting/Readability %.
   - **Matched Keywords:** See which skills are already aligned.
   - **Missing Critical Skills:** Highlights gaps from the JD that need to be addressed.
   - **Insights & Actionable Improvements:** Specific advice on what to improve.
   - Click **"Generate ATS-Optimized Resume & Outreach"** to trigger synthesis.

3. **Stage 3: Review, Edit & PDF Download**
   - **ATS Score Boost Banner:** Highlights the jump from initial score to 95%+ optimized match.
   - **Live PDF Preview:** Inspect the single-page ATS-aligned ReportLab PDF.
   - **Inline Resume Editor:** Click "Edit Resume Fields" to tweak your name, summary, skills, or experience bullets, then click "Apply Changes & Recompile PDF" to update live.
   - **Download PDF Button:** Download the tailored PDF directly to your device.
   - **Personalized Email Editor:** Edit the custom cold outreach email with live word count.
   - Click **"Proceed to Email Recruiter"** to advance to dispatch.

4. **Stage 4: Secure Gmail SMTP Dispatch**
   - Enter your Gmail and 16-character App Password.
   - Inspect the recipient recruiter email, subject line, and attached PDF.
   - Click **"Send Email with PDF"** to deliver your application directly to the recruiter's inbox.
