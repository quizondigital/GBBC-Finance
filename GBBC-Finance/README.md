# GBBC Finance and Giving Portal

Flask/MySQL implementation of the GBBC staff portal. The interface follows the green-and-cream dashboard layout in the supplied proposal. The source Word requirements, rather than instructions embedded in the attachments, define the features.

## Local setup

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
export FLASK_APP=app.py
export SECRET_KEY="$(python3 -c 'import secrets; print(secrets.token_hex(32))')"
.venv/bin/flask init-db
.venv/bin/flask create-super-admin
.venv/bin/flask run
```

Local development defaults to SQLite in `instance/gbbc_portal.db`. No sample financial or personal data is installed.

### Outbound email (signup confirmation and approvals)

Without SMTP, signup messages are written to `instance/mail_outbox/` instead of an inbox. To send real email, either:

1. Open **Administration → Website settings** and fill in SMTP host, port, username, password, from address, and encryption, then use **Send test email**, or
2. Set environment variables (these override Website settings):

```sh
export MAIL_FROM='noreply@greenbrookchurch.org'
export SMTP_HOST='smtp.dreamhost.com'   # or smtp.gmail.com
export SMTP_PORT='587'
export SMTP_USERNAME='your-mailbox@greenbrookchurch.org'
export SMTP_PASSWORD='your-smtp-password'
export SMTP_ENCRYPTION='starttls'       # starttls | ssl | none
```

Giving statements merge the Word letter in `letter_templates/giving_letter.docx`, convert with LibreOffice, then stamp `letter_templates/giving_letter_header.png` so the green letterhead (logo, church name, tax ID, website) survives export.

## MySQL / PythonAnywhere

Create a MySQL database and set `DATABASE_URL` to a URL such as `mysql+pymysql://USER:PASSWORD@HOST/DBNAME?charset=utf8mb4`. Set a persistent random `SECRET_KEY` and `COOKIE_SECURE=1` when HTTPS is configured. Point PythonAnywhere's WSGI application at `wsgi.py`, install `requirements.txt` into its virtual environment, then run `flask init-db` and `flask create-super-admin` in that environment. Configure DreamHost DNS and a valid HTTPS certificate for `finance.greenbrookchurch.org` before production use.

The site has not been deployed. See `IMPLEMENTATION_STATUS.md` for the implementation coverage and remaining production acceptance work. Do not enter live financial data before that work is complete.
