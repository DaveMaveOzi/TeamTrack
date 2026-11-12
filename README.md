# TeamTrack

A multi-company internal ticket and issue tracker built with Streamlit, SQLite and bcrypt.

## Features
- Company workspace creation and invitation-code joining
- Email/password authentication and Admin, Team Lead, Member and Viewer roles
- Department management, ticket creation, assignment and status tracking
- Ticket comments, activity history, in-app notifications and optional SMTP email
- Ticket attachments and responsive workspace interface
- Light and dark mode
- Reports & Analytics with date, department, status, priority and employee filters
- Report exports to PDF, Word (.docx), Excel (.xlsx) and CSV

## Run locally
1. Use Python 3.10 or newer.
2. Open a terminal in this folder.
3. Create and activate a virtual environment.
4. Install dependencies: `pip install -r requirements.txt`
5. Start the app: `streamlit run main.py`

The SQLite database (`teamtrack.db`) and uploaded attachments (`uploads/`) are created automatically. Keep your existing `teamtrack.db` and `uploads/` folder when replacing application code; they contain your local data and attachments.

## Email notifications (optional)
Configure `TEAMTRACK_SMTP_HOST`, `TEAMTRACK_SMTP_PORT`, `TEAMTRACK_SMTP_USERNAME`, `TEAMTRACK_SMTP_PASSWORD`, and `TEAMTRACK_FROM_EMAIL` as environment variables.

## Deployment
For Streamlit Community Cloud, push `main.py`, `requirements.txt` and this README to the GitHub repository, then allow the app to rebuild. Do not commit local database files, uploaded attachments, secrets or `.env` files.

## Important MVP limitations
This remains an MVP, not production-ready for real company data. It uses SQLite and local file storage, and its invitation code is a simple shared code. Before production use, add persistent managed database storage (for example PostgreSQL), durable object storage, stronger invitations, password reset, email verification, rate limiting, comprehensive authorization tests, backups, audit retention and deployment secrets management.
