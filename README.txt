TEAMTRACK — LANDING PAGE UPDATE

This package continues from TeamTrack_Role_Permissions_Update.zip.

Files:
- main.py: Existing TeamTrack application with the new public landing page.
- teamtrack_logo.png: Existing TeamTrack brand logo.
- landing_illustration.png: Illustration cropped from the supplied landing-page design.
- requirements.txt: Existing application dependencies.

Landing page behaviour:
- Get Started opens the existing Create Company registration view.
- Sign In opens the existing Sign In view.
- Existing Join Company access remains available from the authentication view.
- Existing authenticated workspace and features are retained.

The landing page follows the supplied reference composition, with TeamTrack branding in the footer.
Company email-domain restrictions are not enabled; ordinary email addresses remain accepted for testing.

Run locally:
  pip install -r requirements.txt
  streamlit run main.py

Before production use, validate authentication, role permissions, tenant isolation, persistent storage,
email delivery and deployment configuration with test accounts and test data.
