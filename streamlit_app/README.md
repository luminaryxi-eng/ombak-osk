# OSK Ombak – Document Automation
## Deploy to Streamlit Cloud (free, ~5 minutes)

### Step 1 — GitHub
1. Create a free account at github.com
2. Create a new repository called `osk-ombak-automation` (set to **Private**)
3. Upload ALL files from this ZIP into the repository

### Step 2 — Streamlit Cloud
1. Go to share.streamlit.io → sign in with GitHub
2. Click "New app"
3. Select your repository `osk-ombak-automation`
4. Main file: `app.py`
5. Click "Deploy"

### Step 3 — Add API Key
1. In your app settings → **Secrets**
2. Paste:
   ```
   ANTHROPIC_API_KEY = "sk-ant-your-actual-key"
   ```
3. Save → app restarts

### Done!
Share the URL (e.g. `https://your-app.streamlit.app`) with your staff.

---

## Adding bank templates
1. Open your bank's Word document
2. Replace fields with placeholders (see sidebar in the app for full list):
   - `{{BORROWER_NAME_1}}` — Borrower 1 full name
   - `{{BORROWER_NRIC_1}}` — Borrower 1 NRIC
   - `{{BORROWER_NAME_2}}` — Borrower 2 full name (if applicable)
   - `{{UNIT_NO}}` — Unit number
   - `{{LOAN_AMOUNT}}` — Loan amount
   - `{{BANK_NAME}}` — Bank name
   - `{{DATE}}` — Document date
3. Save the file (e.g. `Maybank_Loan_Letter.docx`)
4. Upload via the sidebar in the app → instantly available

## Adding new lawyer forms
Drop any new `.docx` template into the `templates/lawyer/` folder in GitHub.
Use the same placeholder system or the existing fill_forms.py logic.
