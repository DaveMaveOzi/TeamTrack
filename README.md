# TeamTrack MVP

A multi-company internal issue and ticket tracker built with Streamlit, SQLite and bcrypt.

## Run locally

1. Install Python 3.10 or newer.
2. Open a terminal in this folder.
3. Create and activate a virtual environment.
4. Install dependencies:

   `pip install -r requirements.txt`

5. Start the app:

   `streamlit run main.py`

The SQLite database (`teamtrack.db`) and uploaded attachments (`uploads/`) are created automatically.

## First use

- Select **Create company** to create a company workspace. The first user becomes Admin.
- Share the generated company invitation code with employees.
- Employees select **Join a company** to register.
- Members can create tickets, comment and follow updates.

## Email notifications (optional)

In-app notifications work without extra configuration. To enable email notifications, configure these environment variables before starting Streamlit:

- `TEAMTRACK_SMTP_HOST`
- `TEAMTRACK_SMTP_PORT` (usually `587`)
- `TEAMTRACK_SMTP_USERNAME`
- `TEAMTRACK_SMTP_PASSWORD`
- `TEAMTRACK_FROM_EMAIL`

Email delivery is best-effort. In-app notifications remain the source of truth.

## Important MVP limitations

This is a local prototype, not yet production-ready for real company data. It uses SQLite, local file storage and a simple company invitation code. Before public or multi-company production deployment, add PostgreSQL, managed object storage, verified email invitations, password reset, rate limiting, stronger authorization tests, CSRF/session hardening, backups, audit retention and deployment secrets management.
