# Email Outreach Automation

Automated, rate-limited email dispatch pipeline using Python, SMTP, and Excel queue management.

---

### **1. Requirements & Dependencies**

Install the required third-party libraries:

```bash
pip install pandas openpyxl

```

---

### **2. Google Account Setup (2-Step Verification & App Password)**

Gmail and Google Workspace block basic password logins over SMTP. You must generate a dedicated **16-character App Password**:

1. **Enable 2-Step Verification:**
* Go to **[myaccount.google.com/security](https://myaccount.google.com/security)**.
* Under *"How you sign in to Google"*, turn on **2-Step Verification**.


2. **Generate App Password:**
* Search for **"App passwords"** in the Google Account search bar (or navigate to Security --> 2-Step Verification --> App Passwords).
* Enter an app name (e.g., `OutreachScript`) and click **Create**.
* Copy the generated **16-character code** (e.g., `xxxx xxxx xxxx xxxx`).


3. **Configure Script:**
* Paste your address into `SENDER_EMAIL`.
* Paste the 16-character code into `SENDER_PASSWORD` (remove any spaces).



---

### **3. Running the Script**

* Set `EXCEL_FILE` to your target `.xlsx` path containing columns: `Organization`, `Type`, `Head Office Location`, `Official Website`, and `Official Email`.
* Run `main.py`. The script logs delivery status (`SENT` or `INVALID`) in real time, pauses between batches, and safely saves progress on manual termination (`Ctrl + C`).
