# Local Email-to-Excel AI Prototype

No Microsoft Entra, Outlook API, or OpenAI API account required.

## Run locally
```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Open the local Streamlit URL printed in the terminal (usually http://localhost:8501).

## How to use
- Paste email body text, or upload an exported `.eml` file.
- Upload PDF attachments.
- Click Extract information.
- Inspect the result and save it to `extracted_emails.xlsx`.
- Unknown/uncertain cases are copied to the Human Review sheet.

## Optional local AI
Install Ollama from https://ollama.com/, then run:
```bash
ollama pull qwen2.5:3b
```
Enable "Use local AI (Ollama)" in the app. Ollama runs the language model locally; no OpenAI API key is needed.

## Important scope
This is a learning prototype. It does not monitor a live Outlook inbox. Direct inbox monitoring requires an authenticated mail protocol/API; this version avoids that setup by processing exported email files and uploaded PDFs. Scanned PDFs need OCR. Do not use confidential company mail without authorization and appropriate security controls.
# AI-Agent-Extract_info_pdf_from-mail
An AI-powered agent that extracts, processes, and organizes relevant information from PDF attachments received via email. Automates document information extraction to reduce manual effort and streamline email-based document processing workflows.
