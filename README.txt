TEAMTRACK — FEATURE UPDATE

This package continues from the TeamTrack landing-page spacing update.

Included updates:
- Landing page social-proof row with three small human/profile icons and the exact text:
  "Used by 200+ companies".
- Existing landing-page spacing, footer placement, logo and illustration retained.
- Forgot password link on the Sign In view.
- Password-reset links expire after 45 minutes and can be used once.
- Company creation sends a registration confirmation email to the registered company contact (the first Admin email).
- Employee registration sends a confirmation email to the employee and a new-registration notification to the company's Admin email.
- SMTP settings can be read from Streamlit Community Cloud Secrets or environment variables.
- Existing role assignment rules are retained: Admin and IT Personnel can assign across departments; Team Leads can assign within their own department; Members can submit tickets but cannot assign work.
- Normal email addresses remain accepted for testing. Company-domain restrictions are not enforced.

EMAIL CONFIGURATION (Streamlit Community Cloud)

In the app's Settings > Secrets, add the following TOML values using credentials from your email provider:

TEAMTRACK_SMTP_HOST = "smtp.example.com"
TEAMTRACK_SMTP_PORT = "587"
TEAMTRACK_SMTP_USERNAME = "your-smtp-username"
TEAMTRACK_SMTP_PASSWORD = "your-smtp-password-or-app-password"
TEAMTRACK_FROM_EMAIL = "no-reply@yourdomain.com"
TEAMTRACK_APP_URL = "https://teamtrack.streamlit.app"

Replace the example values with real SMTP details. Do not commit real credentials to GitHub.
Email confirmations and password-reset messages will not be delivered until SMTP is configured. The app displays a notice if delivery is unavailable. For Gmail, use the provider's supported SMTP setup and an app password where required; do not use your normal account password.

RUN LOCALLY

  pip install -r requirements.txt
  streamlit run main.py

NOTES

- Password-reset tokens are stored as hashes, expire after 45 minutes and are invalidated after use.
- The company-domain-only registration rule is deliberately deferred for testing.
- The "Used by 200+ companies" line is the requested marketing copy and is not generated from verified customer analytics.
- Before production, configure SMTP, test all email delivery paths, review rate limiting and abuse protection for password resets, and validate access controls and tenant isolation.
- Back up teamtrack.db before replacing an existing deployment or project folder.
