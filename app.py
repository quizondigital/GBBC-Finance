"""GBBC Finance and Giving Portal. Configure DATABASE_URL for MySQL in production."""
import csv
import hashlib
import io
import json
import os
import re
import secrets
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from functools import wraps
from pathlib import Path
from urllib.parse import urlparse

import click
from flask import Flask, abort, current_app, flash, g, redirect, render_template, request, send_file, session, url_for
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import simpleSplit
from reportlab.pdfgen import canvas
from werkzeug.security import check_password_hash, generate_password_hash

ROLES = ('Super Admin', 'System Administrator', 'Finance Staff', 'Church Leader')
US_STATES = (
    ('AL', 'Alabama'), ('AK', 'Alaska'), ('AZ', 'Arizona'), ('AR', 'Arkansas'), ('CA', 'California'),
    ('CO', 'Colorado'), ('CT', 'Connecticut'), ('DE', 'Delaware'), ('DC', 'District of Columbia'),
    ('FL', 'Florida'), ('GA', 'Georgia'), ('HI', 'Hawaii'), ('ID', 'Idaho'), ('IL', 'Illinois'),
    ('IN', 'Indiana'), ('IA', 'Iowa'), ('KS', 'Kansas'), ('KY', 'Kentucky'), ('LA', 'Louisiana'),
    ('ME', 'Maine'), ('MD', 'Maryland'), ('MA', 'Massachusetts'), ('MI', 'Michigan'), ('MN', 'Minnesota'),
    ('MS', 'Mississippi'), ('MO', 'Missouri'), ('MT', 'Montana'), ('NE', 'Nebraska'), ('NV', 'Nevada'),
    ('NH', 'New Hampshire'), ('NJ', 'New Jersey'), ('NM', 'New Mexico'), ('NY', 'New York'),
    ('NC', 'North Carolina'), ('ND', 'North Dakota'), ('OH', 'Ohio'), ('OK', 'Oklahoma'), ('OR', 'Oregon'),
    ('PA', 'Pennsylvania'), ('RI', 'Rhode Island'), ('SC', 'South Carolina'), ('SD', 'South Dakota'),
    ('TN', 'Tennessee'), ('TX', 'Texas'), ('UT', 'Utah'), ('VT', 'Vermont'), ('VA', 'Virginia'),
    ('WA', 'Washington'), ('WV', 'West Virginia'), ('WI', 'Wisconsin'), ('WY', 'Wyoming'),
)
US_STATE_CODES = {code for code, _ in US_STATES}
DEFAULT_FUNDS = ('General Fund', 'Missions Fund', 'Building Fund', 'Benevolence Fund')
FUND_ALIASES = {
    'Missions': 'Missions Fund',
    'Benevolence': 'Benevolence Fund',
}
PERMISSIONS = {
    'Super Admin': {'overview', 'people_view', 'people_write', 'gift_view', 'gift_write', 'reports', 'detail_reports', 'statements', 'users', 'roles', 'fiscal', 'feedback', 'settings', 'feedback_review', 'audit_view'},
    'System Administrator': {'overview', 'people_view', 'reports', 'users', 'roles', 'fiscal', 'feedback', 'settings', 'feedback_review', 'audit_view'},
    'Finance Staff': {'overview', 'people_view', 'people_write', 'gift_view', 'gift_write', 'reports', 'detail_reports', 'statements', 'feedback'},
    'Church Leader': {'overview', 'people_view', 'reports', 'feedback'},
}
PAGE_MATRIX = [
    ('Dashboard', 'overview'), ('People directory', 'people_view'), ('Create and edit people', 'people_write'),
    ('Contribution detail', 'gift_view'), ('Record and correct gifts', 'gift_write'), ('Aggregate giving reports', 'reports'),
    ('Individual giving reports', 'detail_reports'), ('Giving statements', 'statements'),
    ('User administration', 'users'), ('Role permissions', 'roles'), ('Fiscal-year administration', 'fiscal'),
    ('Feedback', 'feedback'), ('Website settings', 'settings'),
    ('Review submitted feedback', 'feedback_review'), ('Audit history', 'audit_view'),
]

db = SQLAlchemy()

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    first_name = db.Column(db.String(100), nullable=False)
    last_name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(255), unique=True, nullable=False)
    username = db.Column(db.String(100), unique=True)
    phone = db.Column(db.String(50), nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(40), nullable=False)
    active = db.Column(db.Boolean, default=True, nullable=False)
    last_login = db.Column(db.DateTime)
    @property
    def name(self): return f'{self.first_name} {self.last_name}'

class Person(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    first_name = db.Column(db.String(100), nullable=False)
    last_name = db.Column(db.String(100), nullable=False)
    member_type = db.Column(db.String(20), nullable=False)
    address_line1 = db.Column(db.String(120), default='')
    address_line2 = db.Column(db.String(120), default='')
    city = db.Column(db.String(100), default='')
    state = db.Column(db.String(2), default='')
    zip_code = db.Column(db.String(20), default='')
    email = db.Column(db.String(255), default='')
    phone = db.Column(db.String(50), default='')
    active = db.Column(db.Boolean, default=True, nullable=False)
    @property
    def name(self): return f'{self.first_name} {self.last_name}'
    @property
    def address_display(self):
        return ', '.join(person_address_lines(self)) or ''

class FiscalYear(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(30), unique=True, nullable=False)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    is_open = db.Column(db.Boolean, default=True, nullable=False)

class Fund(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)
    active = db.Column(db.Boolean, default=True, nullable=False)

class Gift(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    person_id = db.Column(db.Integer, db.ForeignKey('person.id'), nullable=False)
    person = db.relationship('Person')
    fiscal_year_id = db.Column(db.Integer, db.ForeignKey('fiscal_year.id'), nullable=False)
    fiscal_year = db.relationship('FiscalYear')
    fund_id = db.Column(db.Integer, db.ForeignKey('fund.id'), nullable=False)
    fund = db.relationship('Fund')
    date = db.Column(db.Date, nullable=False)
    amount = db.Column(db.Numeric(12, 2), nullable=False)
    method = db.Column(db.String(40), nullable=False)
    reference = db.Column(db.String(120), default='')
    note = db.Column(db.Text, default='')
    status = db.Column(db.String(20), default='active', nullable=False)
    reconciled = db.Column(db.Boolean, default=False, nullable=False)
    entered_by_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    entered_by = db.relationship('User')
    created_at = db.Column(db.DateTime, default=lambda: utcnow(), nullable=False)

class Feedback(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    user = db.relationship('User')
    category = db.Column(db.String(30), nullable=False)
    subject = db.Column(db.String(150), nullable=False)
    message = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(30), default='new')
    created_at = db.Column(db.DateTime, default=lambda: utcnow(), nullable=False)

class Audit(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    action = db.Column(db.String(100), nullable=False)
    entity = db.Column(db.String(60), nullable=False)
    entity_id = db.Column(db.Integer)
    timestamp = db.Column(db.DateTime, default=lambda: utcnow(), nullable=False)
    summary = db.Column(db.String(255), default='')

class Setting(db.Model):
    key = db.Column(db.String(80), primary_key=True)
    value = db.Column(db.Text, nullable=False)

class LoginAttempt(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(255), nullable=False, index=True)
    attempted_at = db.Column(db.DateTime, default=lambda: utcnow(), nullable=False)

def utcnow():
    return datetime.now(timezone.utc).replace(tzinfo=None)

def hash_password(password):
    method = 'scrypt' if hasattr(hashlib, 'scrypt') else 'pbkdf2:sha256'
    return generate_password_hash(password, method=method)

def excel_safe(value):
    text = str(value)
    return "'" + text if text.startswith(('=', '+', '-', '@', '\t', '\r')) else text

def audit(action, entity, entity_id=None, summary=''):
    db.session.add(Audit(user_id=g.user.id if g.user else None, action=action, entity=entity, entity_id=entity_id, summary=summary[:255]))

def valid_email(value):
    return bool(re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+', value))

def parse_amount(raw):
    try:
        value = Decimal(raw)
        if not value.is_finite() or value <= 0 or value.as_tuple().exponent < -2: raise ValueError
        return value.quantize(Decimal('0.01'))
    except (InvalidOperation, ValueError):
        raise ValueError('Enter a positive amount with no more than two decimal places.')

def person_address_lines(person):
    lines = []
    line1 = (person.address_line1 or '').strip()
    line2 = (person.address_line2 or '').strip()
    if line1:
        lines.append(line1)
    if line2:
        lines.append(line2)
    city = (person.city or '').strip()
    state = (person.state or '').strip()
    zip_code = (person.zip_code or '').strip()
    if city or state or zip_code:
        if city and state:
            locality = f'{city}, {state}'
        else:
            locality = city or state
        if zip_code:
            locality = f'{locality} {zip_code}'.strip()
        lines.append(locality)
    return lines

def normalize_state(value):
    code = (value or '').strip().upper()
    if not code:
        return ''
    if code in US_STATE_CODES:
        return code
    for state_code, state_name in US_STATES:
        if state_name.upper() == code:
            return state_code
    if re.fullmatch(r'[A-Z]{2}', code):
        return code
    raise ValueError('Enter a valid 2-letter state code.')

def normalize_zip(value):
    zip_code = (value or '').strip()
    if not zip_code:
        return ''
    if not re.fullmatch(r'\d{5}(-\d{4})?', zip_code):
        raise ValueError('Enter a ZIP code as 12345 or 12345-6789.')
    return zip_code

def read_person_address(source):
    """Read address fields from a form or CSV row dict."""
    get = source.get if hasattr(source, 'get') else lambda key, default='': source.get(key, default)
    line1 = (get('address_line1') or get('address') or '').strip()
    line2 = (get('address_line2') or '').strip()
    city = (get('city') or '').strip()
    state = normalize_state(get('state') or '')
    zip_code = normalize_zip(get('zip_code') or get('zip') or '')
    return line1, line2, city, state, zip_code

def ensure_schema():
    """Add columns introduced after the first release; migrate legacy person.address."""
    from sqlalchemy import inspect, text
    inspector = inspect(db.engine)
    tables = set(inspector.get_table_names())
    if 'user' in tables:
        user_cols = {col['name'] for col in inspector.get_columns('user')}
        if 'username' not in user_cols:
            db.session.execute(text('ALTER TABLE user ADD COLUMN username VARCHAR(100)'))
            db.session.commit()
    if 'person' not in tables:
        return
    cols = {col['name'] for col in inspector.get_columns('person')}
    additions = [
        ('address_line1', 'VARCHAR(120) DEFAULT \'\''),
        ('address_line2', 'VARCHAR(120) DEFAULT \'\''),
        ('city', 'VARCHAR(100) DEFAULT \'\''),
        ('state', 'VARCHAR(2) DEFAULT \'\''),
        ('zip_code', 'VARCHAR(20) DEFAULT \'\''),
    ]
    for name, definition in additions:
        if name not in cols:
            db.session.execute(text(f'ALTER TABLE person ADD COLUMN {name} {definition}'))
            db.session.commit()
            cols.add(name)
    if 'address' in cols:
        db.session.execute(text(
            "UPDATE person SET address_line1 = address "
            "WHERE (address_line1 IS NULL OR address_line1 = '') "
            "AND address IS NOT NULL AND address != ''"
        ))
        db.session.commit()
    if 'fund' in tables:
        for old_name, new_name in FUND_ALIASES.items():
            existing_old = db.session.scalar(db.select(Fund).filter_by(name=old_name))
            existing_new = db.session.scalar(db.select(Fund).filter_by(name=new_name))
            if existing_old and not existing_new:
                existing_old.name = new_name
                db.session.commit()
            elif existing_old and existing_new:
                db.session.execute(text('UPDATE gift SET fund_id = :new_id WHERE fund_id = :old_id'), {'new_id': existing_new.id, 'old_id': existing_old.id})
                db.session.delete(existing_old)
                db.session.commit()
        for name in DEFAULT_FUNDS:
            if not db.session.scalar(db.select(Fund.id).filter_by(name=name)):
                db.session.add(Fund(name=name))
        db.session.commit()

def funds_in_display_order(funds):
    rank = {name: index for index, name in enumerate(DEFAULT_FUNDS)}
    return sorted(funds, key=lambda fund: (rank.get(fund.name, len(DEFAULT_FUNDS)), fund.name.lower()))

def resolve_fund_name(name):
    cleaned = (name or '').strip()
    if not cleaned:
        raise ValueError('fund is required')
    mapped = FUND_ALIASES.get(cleaned, cleaned)
    fund = db.session.scalar(db.select(Fund).where(Fund.active.is_(True), Fund.name == mapped))
    if fund:
        return fund
    if not mapped.endswith(' Fund'):
        fund = db.session.scalar(db.select(Fund).where(Fund.active.is_(True), Fund.name == f'{mapped} Fund'))
        if fund:
            return fund
    raise ValueError(f'unknown fund "{cleaned}"')

def resolve_import_person(row):
    first = (row.get('first_name') or '').strip()
    last = (row.get('last_name') or '').strip()
    email = (row.get('email') or '').strip().lower()
    if not first or not last:
        raise ValueError('first_name and last_name are required')
    matches = db.session.scalars(
        db.select(Person).where(
            Person.active.is_(True),
            Person.first_name.ilike(first),
            Person.last_name.ilike(last),
        )
    ).all()
    if email:
        email_matches = [person for person in matches if (person.email or '').lower() == email]
        if email_matches:
            matches = email_matches
        elif not matches:
            matches = db.session.scalars(
                db.select(Person).where(Person.active.is_(True), Person.email == email)
            ).all()
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise ValueError(f'no active person matched {first} {last}')
    raise ValueError(f'multiple people matched {first} {last}; include email to disambiguate')

def parse_gift_import_row(row, cache):
    person = resolve_import_person(row)
    fund = resolve_fund_name(row.get('fund') or '')
    method = (row.get('method') or '').strip()
    if method not in ('Check', 'Cash', 'Online', 'Other'):
        raise ValueError('method must be Check, Cash, Online, or Other')
    try:
        gift_date = date.fromisoformat((row.get('date') or '').strip())
    except ValueError as exc:
        raise ValueError('date must be YYYY-MM-DD') from exc
    raw_amount = (row.get('amount') or '').strip()
    if not raw_amount:
        raise ValueError('amount is required')
    try:
        amount = parse_amount(raw_amount)
    except ValueError as exc:
        raise ValueError('amount must be a positive number (for example 40.00)') from exc
    year_name = (row.get('fiscal_year') or '').strip()
    if year_name:
        year = cache['years'].get(year_name) or db.session.scalar(db.select(FiscalYear).filter_by(name=year_name))
        if year:
            cache['years'][year_name] = year
    else:
        year = next((item for item in cache['open_years'] if item.start_date <= gift_date <= item.end_date), None)
    if not year:
        raise ValueError('fiscal year not found' if year_name else 'date does not fall in an open fiscal year')
    if not year.is_open:
        raise ValueError(f'fiscal year {year.name} is closed')
    if not (year.start_date <= gift_date <= year.end_date):
        raise ValueError(f'date must fall within fiscal year {year.name}')
    return {
        'person': person,
        'fund': fund,
        'year': year,
        'date': gift_date,
        'amount': amount,
        'method': method,
        'reference': (row.get('reference') or '').strip(),
        'note': (row.get('note') or '').strip(),
    }

def csv_row_blank(row):
    return not any(str(value or '').strip() for value in row.values())

def csv_row_preview(row, fieldnames):
    if not row:
        return ''
    names = fieldnames or list(row.keys())
    return ','.join(str(row.get(name, '') or '').replace('\n', ' ') for name in names)

def read_csv_records(file_storage, required_columns):
    raw = file_storage.read()
    if not raw or not raw.strip():
        raise ValueError('The file is empty.')
    try:
        text = raw.decode('utf-8-sig')
    except UnicodeError as exc:
        raise ValueError('The file could not be read. Save it again as a CSV from Excel or Google Sheets.') from exc
    lines = [line for line in text.splitlines() if line.strip()]
    if not lines:
        raise ValueError('The file is empty.')
    if ',' not in lines[0]:
        raise ValueError('Columns must be separated by commas. Save the file as CSV and try again.')
    try:
        reader = csv.DictReader(io.StringIO(text))
    except csv.Error as exc:
        raise ValueError('The file is not a valid CSV. Check that commas separate each column.') from exc
    if not reader.fieldnames:
        raise ValueError('The file is missing a header row with column names.')
    original_fields = list(reader.fieldnames)
    normalized = [(name or '').strip() for name in original_fields]
    if len(normalized) != len(set(normalized)):
        raise ValueError('The header row has duplicate column names.')
    if not required_columns.issubset(set(normalized)):
        missing = ', '.join(sorted(required_columns - set(normalized)))
        raise ValueError(f'The CSV is missing required columns: {missing}.')
    key_map = dict(zip(original_fields, normalized))
    try:
        records = []
        for raw_row in reader:
            row = {key_map[key]: (raw_row.get(key) or '') for key in original_fields}
            records.append(row)
    except csv.Error as exc:
        raise ValueError('A row could not be read. Check that commas separate each column.') from exc
    data_rows = [row for row in records if not csv_row_blank(row)]
    if not data_rows:
        raise ValueError('The file has no data rows to import.')
    if len(data_rows) > 10000:
        raise ValueError('CSV exceeds 10,000 rows.')
    return normalized, data_rows

def find_duplicate_person(first, last, email):
    email_l = email.lower()
    matches = db.session.scalars(
        db.select(Person).where(Person.first_name.ilike(first), Person.last_name.ilike(last))
    ).all()
    for person in matches:
        person_email = (person.email or '').strip().lower()
        if email_l and person_email == email_l:
            return person
        if not email_l and not person_email:
            return person
    if email_l:
        return db.session.scalar(db.select(Person).where(Person.email == email))
    return None

def find_duplicate_gift(person_id, gift_date, amount, fund_id, method, reference):
    return db.session.scalar(
        db.select(Gift).where(
            Gift.status == 'active',
            Gift.person_id == person_id,
            Gift.date == gift_date,
            Gift.amount == amount,
            Gift.fund_id == fund_id,
            Gift.method == method,
            Gift.reference == reference,
        )
    )

def import_report_path():
    folder = Path(current_app.instance_path)
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f'import_report_{g.user.id}.json'

def store_import_report(report):
    import_report_path().write_text(json.dumps(report), encoding='utf-8')

def load_import_report():
    path = import_report_path()
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except (OSError, json.JSONDecodeError):
        return None

def build_failure_csv(fieldnames, failures):
    out = io.StringIO()
    writer = csv.DictWriter(out, fieldnames=list(fieldnames) + ['failure_reason'], extrasaction='ignore')
    writer.writeheader()
    for item in failures:
        row = dict(item.get('row_data') or {})
        row['failure_reason'] = item['reason']
        writer.writerow({name: row.get(name, '') for name in writer.fieldnames})
    return out.getvalue()

def statement_year_label(gifts):
    year_id = request.args.get('year_id', type=int)
    if year_id:
        fiscal = db.session.get(FiscalYear, year_id)
        if fiscal:
            return fiscal.name
    years = sorted({gift.date.year for gift in gifts})
    if len(years) == 1:
        return str(years[0])
    if len(years) > 1:
        return f'{years[0]}–{years[-1]}'
    return str(date.today().year)

def draw_paragraphs(pdf, paragraphs, x, y, width, font='Helvetica', size=10, leading=14, gap=10):
    pdf.setFont(font, size)
    for paragraph in paragraphs:
        lines = simpleSplit(paragraph, font, size, width) or ['']
        for line in lines:
            if y < 72:
                pdf.showPage()
                pdf.setFont(font, size)
                y = 740
            pdf.drawString(x, y, line)
            y -= leading
        y -= gap
    return y

def create_app(test_config=None):
    if os.environ.get('PORTAL_PRODUCTION') == '1' and not os.environ.get('SECRET_KEY') and not test_config:
        raise RuntimeError('SECRET_KEY must be set in production.')
    app = Flask(__name__)
    app.config.update(
        SECRET_KEY=os.environ.get('SECRET_KEY', secrets.token_hex(32)),
        SQLALCHEMY_DATABASE_URI=os.environ.get('DATABASE_URL', 'sqlite:///gbbc_portal.db'),
        SQLALCHEMY_TRACK_MODIFICATIONS=False,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE='Lax',
        SESSION_COOKIE_SECURE=os.environ.get('COOKIE_SECURE', '0') == '1' or os.environ.get('PORTAL_PRODUCTION') == '1',
        PERMANENT_SESSION_LIFETIME=1800,
        MAX_CONTENT_LENGTH=2 * 1024 * 1024,
    )
    if test_config: app.config.update(test_config)
    db.init_app(app)
    with app.app_context():
        try:
            ensure_schema()
        except Exception:
            pass

    @app.after_request
    def response_headers(response):
        response.headers['X-Content-Type-Options'] = 'nosniff'
        response.headers['X-Frame-Options'] = 'DENY'
        response.headers['Referrer-Policy'] = 'same-origin'
        if request.path != '/login': response.headers['Cache-Control'] = 'no-store'
        return response

    @app.before_request
    def load_user_and_csrf():
        g.user = db.session.get(User, session.get('user_id')) if session.get('user_id') else None
        if g.user and not g.user.active:
            session.clear(); g.user = None
        if 'csrf' not in session: session['csrf'] = secrets.token_urlsafe(32)
        if request.method in ('POST', 'PUT', 'PATCH', 'DELETE'):
            if not secrets.compare_digest(session['csrf'], request.form.get('csrf_token', '')): abort(400)
        session.permanent = True

    @app.context_processor
    def context():
        def can(permission): return bool(g.user and permission in PERMISSIONS[g.user.role])
        return {'current_user': g.user, 'can': can, 'csrf_token': lambda: session['csrf'], 'roles': ROLES, 'today': date.today(), 'us_states': US_STATES}

    def require(permission):
        def deco(fn):
            @wraps(fn)
            def wrapper(*args, **kwargs):
                if not g.user: return redirect(url_for('login', next=request.path))
                if permission not in PERMISSIONS[g.user.role]: abort(403)
                return fn(*args, **kwargs)
            return wrapper
        return deco

    @app.route('/login', methods=['GET', 'POST'])
    def login():
        if request.method == 'POST':
            identifier = request.form.get('email', '').strip().lower()
            recent = db.session.scalar(db.select(db.func.count(LoginAttempt.id)).where(LoginAttempt.email == identifier, LoginAttempt.attempted_at >= utcnow() - timedelta(minutes=15)))
            if recent >= 5:
                flash('Too many sign-in attempts. Try again in 15 minutes.', 'error')
                return render_template('login.html'), 429
            user = db.session.scalar(db.select(User).where(db.or_(User.email == identifier, User.username == identifier), User.active.is_(True)))
            if user and check_password_hash(user.password_hash, request.form.get('password', '')):
                session.clear(); session['user_id'] = user.id; session['csrf'] = secrets.token_urlsafe(32)
                user.last_login = utcnow(); g.user = user; audit('login', 'user', user.id); db.session.commit()
                target = request.args.get('next', '/')
                return redirect(target if target.startswith('/') and not target.startswith('//') else '/')
            db.session.add(LoginAttempt(email=identifier)); audit('failed_login', 'user', user.id if user else None); db.session.commit()
            flash('Invalid email or password.', 'error')
        return render_template('login.html')

    @app.post('/logout')
    def logout():
        if g.user: audit('logout', 'user', g.user.id); db.session.commit()
        session.clear(); return redirect(url_for('login'))

    @app.route('/')
    @require('overview')
    def dashboard():
        fy = db.session.scalar(db.select(FiscalYear).order_by(FiscalYear.start_date.desc()))
        active_gifts = db.session.scalars(db.select(Gift).filter_by(status='active').order_by(Gift.date.desc()).limit(5)).all() if 'gift_view' in PERMISSIONS[g.user.role] else []
        total_query = db.select(db.func.sum(Gift.amount)).where(Gift.status == 'active')
        if fy: total_query = total_query.where(Gift.fiscal_year_id == fy.id)
        total = db.session.scalar(total_query) or Decimal('0')
        count = db.session.scalar(db.select(db.func.count(Person.id))) or 0
        return render_template('dashboard.html', gifts=active_gifts, total=total, people_count=count, fy=fy)

    @app.route('/people')
    @require('people_view')
    def people():
        term = request.args.get('q', '').strip()
        query = db.select(Person).order_by(Person.last_name, Person.first_name)
        if term: query = query.where(db.or_(Person.first_name.ilike(f'%{term}%'), Person.last_name.ilike(f'%{term}%'), Person.email.ilike(f'%{term}%')))
        return render_template('people.html', people=db.session.scalars(query.limit(200)).all(), q=term)

    @app.route('/people/new', methods=['GET', 'POST'])
    @require('people_write')
    def person_new(): return person_editor()

    @app.route('/people/<int:person_id>/edit', methods=['GET', 'POST'])
    @require('people_write')
    def person_edit(person_id): return person_editor(db.get_or_404(Person, person_id))

    def person_editor(person=None):
        if request.method == 'POST':
            first = request.form.get('first_name', '').strip(); last = request.form.get('last_name', '').strip()
            kind = request.form.get('member_type', '')
            if not first or not last or kind not in ('Member', 'Non-member'):
                flash('First name, last name, and member type are required.', 'error')
            else:
                try:
                    line1, line2, city, state, zip_code = read_person_address(request.form)
                except ValueError as exc:
                    flash(str(exc), 'error')
                else:
                    if person is None: person = Person(); db.session.add(person)
                    person.first_name = first; person.last_name = last; person.member_type = kind
                    person.email = request.form.get('email', '').strip(); person.phone = request.form.get('phone', '').strip()
                    person.address_line1 = line1; person.address_line2 = line2
                    person.city = city; person.state = state; person.zip_code = zip_code
                    db.session.flush(); audit('save', 'person', person.id); db.session.commit()
                    flash('Person saved.', 'success'); return redirect(url_for('people'))
        return render_template('person_form.html', person=person)

    @app.route('/people/import', methods=['GET', 'POST'])
    @require('people_write')
    def people_import():
        if request.method == 'POST':
            file = request.files.get('file')
            if not file or not file.filename.lower().endswith('.csv'):
                flash('Choose a CSV file.', 'error')
            else:
                try:
                    fieldnames, records = read_csv_records(file, {'first_name', 'last_name', 'member_type'})
                except ValueError as exc:
                    store_import_report({
                        'kind': 'people',
                        'title': 'People import report',
                        'success_count': 0,
                        'failed_count': 0,
                        'file_error': str(exc),
                        'failures': [],
                        'fieldnames': [],
                        'back_endpoint': 'people_import',
                        'list_endpoint': 'people',
                        'list_label': 'View people',
                    })
                    return redirect(url_for('import_report'))
                successes = 0
                failures = []
                seen = set()
                for index, row in enumerate(records, start=2):
                    first = (row.get('first_name') or '').strip()
                    last = (row.get('last_name') or '').strip()
                    kind = (row.get('member_type') or '').strip()
                    email = (row.get('email') or '').strip()
                    preview = csv_row_preview(row, fieldnames)
                    try:
                        if not first or not last:
                            raise ValueError('first name and last name are required')
                        if kind not in ('Member', 'Non-member'):
                            raise ValueError('member_type must be Member or Non-member')
                        line1, line2, city, state, zip_code = read_person_address(row)
                        dup_key = (first.lower(), last.lower(), email.lower())
                        if dup_key in seen:
                            raise ValueError('duplicate of another row in this file')
                        if find_duplicate_person(first, last, email):
                            raise ValueError('duplicate of an existing person record')
                        seen.add(dup_key)
                        db.session.add(Person(
                            first_name=first, last_name=last, member_type=kind, email=email,
                            phone=(row.get('phone') or '').strip(), address_line1=line1, address_line2=line2,
                            city=city, state=state, zip_code=zip_code,
                        ))
                        successes += 1
                    except ValueError as exc:
                        failures.append({'row': index, 'reason': str(exc), 'preview': preview, 'row_data': row})
                if successes:
                    audit('import', 'person', summary=f'{successes} loaded, {len(failures)} failed')
                    db.session.commit()
                else:
                    db.session.rollback()
                store_import_report({
                    'kind': 'people',
                    'title': 'People import report',
                    'success_count': successes,
                    'failed_count': len(failures),
                    'file_error': '',
                    'failures': failures,
                    'fieldnames': fieldnames,
                    'back_endpoint': 'people_import',
                    'list_endpoint': 'people',
                    'list_label': 'View people',
                })
                return redirect(url_for('import_report'))
        return render_template('import.html')

    @app.route('/people/template.csv')
    @require('people_write')
    def people_template():
        header = 'first_name,last_name,member_type,email,phone,address_line1,address_line2,city,state,zip_code\n'
        content = io.BytesIO(header.encode('utf-8'))
        return send_file(content, as_attachment=True, download_name='GBBC_people_template.csv', mimetype='text/csv')

    @app.route('/contributions')
    @require('gift_view')
    def contributions():
        per_page = 25
        sort_keys = {
            'date': Gift.date,
            'contributor': (Person.last_name, Person.first_name),
            'fund': Fund.name,
            'method': Gift.method,
            'amount': Gift.amount,
            'reconciled': Gift.reconciled,
            'status': Gift.status,
        }
        sort = request.args.get('sort', 'date')
        if sort not in sort_keys:
            sort = 'date'
        direction = request.args.get('dir', 'desc' if sort == 'date' else 'asc')
        if direction not in ('asc', 'desc'):
            direction = 'desc' if sort == 'date' else 'asc'
        total = db.session.scalar(db.select(db.func.count()).select_from(Gift)) or 0
        total_pages = max(1, (total + per_page - 1) // per_page)
        page = request.args.get('page', 1, type=int) or 1
        page = min(max(page, 1), total_pages)
        query = db.select(Gift)
        if sort == 'contributor':
            query = query.join(Person, Gift.person_id == Person.id)
        elif sort == 'fund':
            query = query.join(Fund, Gift.fund_id == Fund.id)
        columns = sort_keys[sort]
        if not isinstance(columns, tuple):
            columns = (columns,)
        order = [col.desc() if direction == 'desc' else col.asc() for col in columns]
        order.append(Gift.id.desc() if direction == 'desc' else Gift.id.asc())
        gifts = db.session.scalars(query.order_by(*order).offset((page - 1) * per_page).limit(per_page)).all()
        return render_template(
            'contributions.html',
            gifts=gifts,
            sort=sort,
            direction=direction,
            page=page,
            total_pages=total_pages,
            total=total,
            per_page=per_page,
        )

    @app.route('/contributions/import', methods=['GET', 'POST'])
    @require('gift_write')
    def gifts_import():
        if request.method == 'POST':
            file = request.files.get('file')
            if not file or not file.filename.lower().endswith('.csv'):
                flash('Choose a CSV file.', 'error')
            else:
                try:
                    fieldnames, records = read_csv_records(file, {'date', 'first_name', 'last_name', 'fund', 'amount', 'method'})
                except ValueError as exc:
                    store_import_report({
                        'kind': 'contributions',
                        'title': 'Contributions import report',
                        'success_count': 0,
                        'failed_count': 0,
                        'file_error': str(exc),
                        'failures': [],
                        'fieldnames': [],
                        'back_endpoint': 'gifts_import',
                        'list_endpoint': 'contributions',
                        'list_label': 'View contributions',
                    })
                    return redirect(url_for('import_report'))
                cache = {
                    'years': {},
                    'open_years': db.session.scalars(db.select(FiscalYear).where(FiscalYear.is_open.is_(True))).all(),
                }
                successes = 0
                failures = []
                seen = set()
                for index, row in enumerate(records, start=2):
                    preview = csv_row_preview(row, fieldnames)
                    try:
                        item = parse_gift_import_row(row, cache)
                        dup_key = (
                            item['person'].id, item['date'].isoformat(), str(item['amount']),
                            item['fund'].id, item['method'], item['reference'],
                        )
                        if dup_key in seen:
                            raise ValueError('duplicate of another row in this file')
                        if find_duplicate_gift(item['person'].id, item['date'], item['amount'], item['fund'].id, item['method'], item['reference']):
                            raise ValueError('duplicate of an existing contribution')
                        seen.add(dup_key)
                        db.session.add(Gift(
                            person=item['person'], fund=item['fund'], fiscal_year=item['year'],
                            date=item['date'], amount=item['amount'], method=item['method'],
                            reference=item['reference'], note=item['note'], entered_by_id=g.user.id,
                        ))
                        successes += 1
                    except ValueError as exc:
                        failures.append({'row': index, 'reason': str(exc), 'preview': preview, 'row_data': row})
                if successes:
                    audit('import', 'gift', summary=f'{successes} loaded, {len(failures)} failed')
                    db.session.commit()
                else:
                    db.session.rollback()
                store_import_report({
                    'kind': 'contributions',
                    'title': 'Contributions import report',
                    'success_count': successes,
                    'failed_count': len(failures),
                    'file_error': '',
                    'failures': failures,
                    'fieldnames': fieldnames,
                    'back_endpoint': 'gifts_import',
                    'list_endpoint': 'contributions',
                    'list_label': 'View contributions',
                })
                return redirect(url_for('import_report'))
        return render_template('contributions_import.html')

    @app.route('/import/report')
    def import_report():
        if not g.user:
            return redirect(url_for('login', next=request.path))
        report = load_import_report()
        if not report:
            flash('No recent import report is available.', 'error')
            return redirect(url_for('dashboard'))
        if report.get('kind') == 'people' and 'people_write' not in PERMISSIONS[g.user.role]:
            abort(403)
        if report.get('kind') == 'contributions' and 'gift_write' not in PERMISSIONS[g.user.role]:
            abort(403)
        return render_template('import_report.html', report=report)

    @app.route('/import/report/failures.csv')
    def import_report_failures():
        if not g.user:
            return redirect(url_for('login', next=request.path))
        report = load_import_report()
        if not report or not report.get('failures'):
            abort(404)
        if report.get('kind') == 'people' and 'people_write' not in PERMISSIONS[g.user.role]:
            abort(403)
        if report.get('kind') == 'contributions' and 'gift_write' not in PERMISSIONS[g.user.role]:
            abort(403)
        content = build_failure_csv(report.get('fieldnames') or [], report['failures'])
        return send_file(
            io.BytesIO(content.encode('utf-8')),
            as_attachment=True,
            download_name=f"GBBC_{report.get('kind', 'import')}_failed_rows.csv",
            mimetype='text/csv',
        )

    @app.route('/contributions/template.csv')
    @require('gift_write')
    def gifts_template():
        header = 'date,first_name,last_name,email,fund,amount,method,fiscal_year,reference,note\n'
        example = '2026-09-01,Ada,Member,ada@example.org,General Fund,25.00,Check,2026,1001,\n'
        content = io.BytesIO((header + example).encode('utf-8'))
        return send_file(content, as_attachment=True, download_name='GBBC_contributions_template.csv', mimetype='text/csv')

    @app.route('/contributions/new', methods=['GET', 'POST'])
    @require('gift_write')
    def gift_new(): return gift_editor()

    @app.route('/contributions/<int:gift_id>/edit', methods=['GET', 'POST'])
    @require('gift_write')
    def gift_edit(gift_id): return gift_editor(db.get_or_404(Gift, gift_id))

    def gift_editor(gift=None):
        years = db.session.scalars(db.select(FiscalYear).order_by(FiscalYear.start_date.desc())).all()
        funds = funds_in_display_order(db.session.scalars(db.select(Fund).filter_by(active=True)).all())
        persons = db.session.scalars(db.select(Person).filter_by(active=True).order_by(Person.last_name)).all()
        default_fund = next((fund for fund in funds if fund.name == 'General Fund'), funds[0] if funds else None)
        if request.method == 'POST':
            try:
                year = db.session.get(FiscalYear, int(request.form['fiscal_year_id']))
                person = db.session.get(Person, int(request.form['person_id']))
                fund = db.session.get(Fund, int(request.form['fund_id']))
                gift_date = date.fromisoformat(request.form['date']); amount = parse_amount(request.form['amount'])
                if not year or not person or not fund or not year.is_open: raise ValueError('Choose an active person, fund, and open fiscal year.')
                if not (year.start_date <= gift_date <= year.end_date): raise ValueError('Date must fall within the selected fiscal year.')
                if gift and not gift.fiscal_year.is_open: raise ValueError('Closed-year gifts cannot be edited.')
                if gift is None: gift = Gift(entered_by_id=g.user.id); db.session.add(gift)
                gift.person = person; gift.fund = fund; gift.fiscal_year = year; gift.date = gift_date; gift.amount = amount
                gift.method = request.form.get('method', '').strip(); gift.reference = request.form.get('reference', '').strip()
                gift.note = request.form.get('note', '').strip()
                if gift.method not in ('Check', 'Cash', 'Online', 'Other'): raise ValueError('Choose a payment method.')
                db.session.flush(); audit('save', 'gift', gift.id); db.session.commit()
                flash('Contribution saved.', 'success'); return redirect(url_for('contributions'))
            except (ValueError, KeyError) as exc: flash(str(exc), 'error')
        return render_template('gift_form.html', gift=gift, years=years, funds=funds, people=persons, default_fund_id=default_fund.id if default_fund else None)

    @app.post('/contributions/<int:gift_id>/void')
    @require('gift_write')
    def gift_void(gift_id):
        gift = db.get_or_404(Gift, gift_id)
        if not gift.fiscal_year.is_open: abort(403)
        gift.status = 'void'; audit('void', 'gift', gift.id); db.session.commit(); flash('Contribution voided.', 'success')
        return redirect(url_for('contributions'))

    @app.post('/contributions/<int:gift_id>/delete')
    @require('gift_write')
    def gift_delete(gift_id):
        gift = db.get_or_404(Gift, gift_id)
        if not gift.fiscal_year.is_open:
            flash('Closed-year contributions cannot be deleted.', 'error')
            return redirect(url_for('contributions'))
        summary = f'{gift.person.name} ${gift.amount} on {gift.date.isoformat()}'
        db.session.delete(gift)
        audit('delete', 'gift', gift_id, summary=summary)
        db.session.commit()
        flash('Contribution deleted.', 'success')
        return redirect(url_for('contributions'))

    @app.post('/contributions/<int:gift_id>/reconcile')
    @require('gift_write')
    def gift_reconcile(gift_id):
        gift = db.get_or_404(Gift, gift_id)
        if gift.status != 'active' or not gift.fiscal_year.is_open: abort(403)
        gift.reconciled = not gift.reconciled
        audit('reconcile', 'gift', gift.id, str(gift.reconciled)); db.session.commit()
        flash('Reconciliation status updated.', 'success')
        return redirect(url_for('contributions'))

    def filtered_gifts():
        year_id = request.args.get('year_id', type=int)
        person_id = request.args.get('person_id', type=int)
        query = db.select(Gift).where(Gift.status == 'active')
        if year_id: query = query.where(Gift.fiscal_year_id == year_id)
        if person_id: query = query.where(Gift.person_id == person_id)
        for field, bound in (('start_date', 'start'), ('end_date', 'end')):
            raw = request.args.get(field)
            if raw:
                try:
                    parsed = date.fromisoformat(raw)
                except ValueError:
                    abort(400)
                query = query.where(Gift.date >= parsed if bound == 'start' else Gift.date <= parsed)
        return db.session.scalars(query.order_by(Gift.date, Gift.id)).all()

    def report_gifts():
        gifts = filtered_gifts()
        minimum = request.args.get('minimum_total', '').strip()
        if not minimum: return gifts
        try:
            threshold = Decimal(minimum)
            if threshold < 0: raise ValueError
        except (InvalidOperation, ValueError): abort(400)
        person_totals = {}
        for gift in gifts: person_totals[gift.person_id] = person_totals.get(gift.person_id, Decimal('0')) + gift.amount
        return [gift for gift in gifts if person_totals[gift.person_id] >= threshold]

    @app.route('/reports')
    @require('reports')
    def reports():
        if request.args.get('person_id') and 'detail_reports' not in PERMISSIONS[g.user.role]: abort(403)
        if request.args.get('minimum_total') and 'detail_reports' not in PERMISSIONS[g.user.role]: abort(403)
        years = db.session.scalars(db.select(FiscalYear).order_by(FiscalYear.start_date.desc())).all()
        persons = db.session.scalars(db.select(Person).order_by(Person.last_name)).all() if 'detail_reports' in PERMISSIONS[g.user.role] else []
        gifts = report_gifts()
        total = sum((x.amount for x in gifts), Decimal('0'))
        funds = {}
        for gift in gifts: funds[gift.fund.name] = funds.get(gift.fund.name, Decimal('0')) + gift.amount
        monthly = {}
        for gift in gifts:
            period = gift.date.strftime('%b %Y')
            monthly[period] = monthly.get(period, Decimal('0')) + gift.amount
        audit('view', 'report', summary='aggregate' if not persons else 'detail'); db.session.commit()
        return render_template('reports.html', years=years, people=persons, gifts=gifts if 'detail_reports' in PERMISSIONS[g.user.role] else [], total=total, funds=funds, monthly=monthly, selected_year=request.args.get('year_id', type=int))

    @app.route('/reports/export.xlsx')
    @require('detail_reports')
    def report_xlsx():
        gifts = report_gifts(); wb = Workbook(); ws = wb.active; ws.title = 'Giving'
        ws.append(['Date', 'Contributor', 'Address line 1', 'Address line 2', 'City', 'State', 'ZIP', 'Fund', 'Payment method', 'Reference', 'Amount'])
        for gift in gifts:
            person = gift.person
            ws.append([
                gift.date.isoformat(), excel_safe(person.name),
                excel_safe(person.address_line1 or ''), excel_safe(person.address_line2 or ''),
                excel_safe(person.city or ''), excel_safe(person.state or ''), excel_safe(person.zip_code or ''),
                excel_safe(gift.fund.name), excel_safe(gift.method), excel_safe(gift.reference), float(gift.amount),
            ])
        ws.append(['', '', '', '', '', '', '', '', '', 'Total', float(sum((x.amount for x in gifts), Decimal('0')))])
        out = io.BytesIO(); wb.save(out); out.seek(0)
        audit('export', 'report', summary='xlsx'); db.session.commit()
        return send_file(out, download_name='GBBC_giving_report.xlsx', as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    @app.route('/reports/statements/<int:person_id>.pdf')
    @require('statements')
    def statement(person_id):
        person = db.get_or_404(Person, person_id)
        gifts = [x for x in filtered_gifts() if x.person_id == person_id]
        church_setting = db.session.get(Setting, 'church_name')
        church_name = church_setting.value if church_setting else 'Green Brook Baptist Church'
        year_label = statement_year_label(gifts)
        total = sum((x.amount for x in gifts), Decimal('0'))
        amount_text = f'${total:,.2f}'
        out = io.BytesIO()
        pdf = canvas.Canvas(out, pagesize=letter, pageCompression=0 if app.config.get('TESTING') else 1)
        left, width, y = 54, 504, 740

        pdf.setFont('Helvetica', 11)
        pdf.drawString(left, y, f'{person.first_name} {person.last_name}'); y -= 16
        for line in person_address_lines(person):
            pdf.drawString(left, y, line); y -= 16
        y -= 18

        pdf.drawString(left, y, f'Dear {person.first_name},'); y -= 24
        y = draw_paragraphs(pdf, [
            f'We gratefully acknowledge your gift to {church_name} throughout {year_label} in the amount of {amount_text}. Through your generous giving, our congregation has been able to support the work of Jesus Christ; locally and globally through our supported missionaries, organizations, and Bible schools.',
            f'Attached is an itemized statement of your contributions for {year_label} based on our records. If you have any concerns about the accuracy of this information, please feel free to let us know.',
            'For income tax purposes, it is important for us to state here that you did not receive any goods or services in return for any of these contributions other than intangible religious benefits. You made this gift out of your own generosity and commitment to Jesus Christ.',
            'Once again, thank you for your generosity and commitment to the work of Jesus Christ through this church.',
        ], left, y, width, size=11, leading=15, gap=12)

        y -= 6
        pdf.setFont('Helvetica', 11)
        pdf.drawString(left, y, 'Sincerely,'); y -= 28
        pdf.drawString(left, y, 'Romella Seepaul'); y -= 15
        pdf.drawString(left, y, 'Financial Secretary'); y -= 15
        pdf.drawString(left, y, church_name); y -= 36

        scripture = (
            '"Whoever sows sparingly will also reap sparingly, and whoever sows bountifully will also '
            'reap bountifully. Each one must give as he has decided in his heart, not reluctantly or '
            'under compulsion, for God loves a cheerful giver. And God is able to make all grace abound '
            'to you, so that having all sufficiency in all things at all times, you may abound in every good work."'
        )
        y = draw_paragraphs(pdf, [scripture], left, y, width, font='Helvetica-Oblique', size=9, leading=12, gap=6)
        pdf.setFont('Helvetica', 9)
        pdf.drawRightString(left + width, y, '2 Corinthians 9:6-8')

        pdf.showPage()
        pdf.setFillColor(colors.HexColor('#174e3b'))
        pdf.setFont('Helvetica-Bold', 14)
        pdf.drawString(left, 740, church_name)
        pdf.setFillColor(colors.black)
        pdf.setFont('Helvetica', 10)
        pdf.drawString(left, 720, f'Itemized contributions for {person.name} — {year_label}')
        y = 690
        pdf.setFont('Helvetica-Bold', 10)
        pdf.drawString(left, y, 'Date'); pdf.drawString(140, y, 'Fund'); pdf.drawRightString(left + width, y, 'Amount'); y -= 20
        pdf.setFont('Helvetica', 10)
        for gift in gifts:
            if y < 72:
                pdf.showPage(); y = 740
                pdf.setFont('Helvetica', 10)
            pdf.drawString(left, y, gift.date.isoformat())
            pdf.drawString(140, y, gift.fund.name[:50])
            pdf.drawRightString(left + width, y, f'${gift.amount:,.2f}')
            y -= 18
        pdf.setFont('Helvetica-Bold', 11)
        pdf.drawRightString(left + width, max(56, y - 10), f'Total: {amount_text}')
        pdf.save(); out.seek(0)
        audit('export', 'statement', person_id); db.session.commit()
        return send_file(out, download_name=f'GBBC_statement_{person_id}.pdf', as_attachment=True, mimetype='application/pdf')

    @app.route('/admin/users', methods=['GET', 'POST'])
    @require('users')
    def users():
        if request.method == 'POST':
            first = request.form.get('first_name', '').strip(); last = request.form.get('last_name', '').strip()
            email = request.form.get('email', '').strip().lower(); phone = request.form.get('phone', '').strip()
            role = request.form.get('role', ''); password = request.form.get('password', '')
            username = request.form.get('username', '').strip().lower() or None
            if not all((first, last, phone, password)) or len(password) < 12 or not valid_email(email) or role not in ROLES:
                flash('First name, last name, valid email, phone, role, and password are required.', 'error')
            elif db.session.scalar(db.select(User).filter_by(email=email)):
                flash('That email address is already used.', 'error')
            elif username and db.session.scalar(db.select(User).filter_by(username=username)):
                flash('That username is already used.', 'error')
            elif role == 'Super Admin' and g.user.role != 'Super Admin':
                abort(403)
            else:
                user = User(first_name=first, last_name=last, email=email, username=username, phone=phone, role=role, password_hash=hash_password(password))
                db.session.add(user); db.session.flush(); audit('create', 'user', user.id); db.session.commit()
                flash('User created.', 'success'); return redirect(url_for('users'))
        all_users = db.session.scalars(db.select(User).order_by(User.last_name)).all()
        return render_template('users.html', users=all_users)

    @app.post('/admin/users/<int:user_id>/toggle')
    @require('users')
    def user_toggle(user_id):
        user = db.get_or_404(User, user_id)
        if user.id == g.user.id or (user.role == 'Super Admin' and g.user.role != 'Super Admin'): abort(403)
        if user.role == 'Super Admin' and user.active:
            count = db.session.scalar(db.select(db.func.count(User.id)).filter_by(role='Super Admin', active=True))
            if count <= 1: abort(403)
        user.active = not user.active; audit('status', 'user', user.id, str(user.active)); db.session.commit()
        return redirect(url_for('users'))

    @app.route('/admin/users/<int:user_id>/edit', methods=['GET', 'POST'])
    @require('users')
    def user_edit(user_id):
        user = db.get_or_404(User, user_id)
        if user.role == 'Super Admin' and g.user.role != 'Super Admin': abort(403)
        if request.method == 'POST':
            first = request.form.get('first_name', '').strip()
            last = request.form.get('last_name', '').strip()
            email = request.form.get('email', '').strip().lower()
            username = request.form.get('username', '').strip().lower() or None
            phone = request.form.get('phone', '').strip()
            role = request.form.get('role', '')
            other = db.session.scalar(db.select(User).where(User.email == email, User.id != user.id))
            other_username = db.session.scalar(db.select(User).where(User.username == username, User.id != user.id)) if username else None
            if not first or not last or not phone or not valid_email(email) or role not in ROLES:
                flash('First name, last name, valid email, phone, and role are required.', 'error')
            elif other:
                flash('That email address is already used.', 'error')
            elif other_username:
                flash('That username is already used.', 'error')
            elif role == 'Super Admin' and g.user.role != 'Super Admin':
                abort(403)
            elif user.role == 'Super Admin' and role != 'Super Admin' and db.session.scalar(db.select(db.func.count(User.id)).filter_by(role='Super Admin', active=True)) <= 1:
                flash('The last active Super Admin cannot change role.', 'error')
            else:
                user.first_name = first; user.last_name = last; user.email = email; user.username = username; user.phone = phone; user.role = role
                password = request.form.get('password', '')
                if password:
                    if len(password) < 12:
                        flash('A replacement password must contain at least 12 characters.', 'error')
                        return render_template('user_edit.html', user=user)
                    user.password_hash = hash_password(password)
                audit('update', 'user', user.id); db.session.commit()
                flash('User updated.', 'success'); return redirect(url_for('users'))
        return render_template('user_edit.html', user=user)

    @app.route('/admin/roles')
    @require('roles')
    def role_matrix(): return render_template('roles.html', matrix=PAGE_MATRIX, permissions=PERMISSIONS)

    @app.route('/admin/settings', methods=['GET', 'POST'])
    @require('settings')
    def settings_page():
        if request.method == 'POST':
            for key in ('church_name', 'church_address'):
                value = request.form.get(key, '').strip()
                if not value:
                    flash('Church name and mailing address are required.', 'error')
                    return redirect(url_for('settings_page'))
                setting = db.session.get(Setting, key)
                if setting is None:
                    setting = Setting(key=key, value=value); db.session.add(setting)
                else: setting.value = value
            audit('update', 'settings'); db.session.commit()
            flash('Website settings saved.', 'success'); return redirect(url_for('settings_page'))
        values = {item.key: item.value for item in db.session.scalars(db.select(Setting)).all()}
        return render_template('settings.html', values=values)

    @app.route('/admin/fiscal', methods=['GET', 'POST'])
    @require('fiscal')
    def fiscal():
        if request.method == 'POST':
            action = request.form.get('action')
            if action == 'create':
                try:
                    start = date.fromisoformat(request.form['start_date']); end = date.fromisoformat(request.form['end_date'])
                    if start >= end: raise ValueError('End date must follow start date.')
                    for existing in db.session.scalars(db.select(FiscalYear)).all():
                        if start <= existing.end_date and end >= existing.start_date: raise ValueError('Fiscal years cannot overlap.')
                    year = FiscalYear(name=request.form['name'].strip(), start_date=start, end_date=end)
                    db.session.add(year); db.session.flush(); audit('create', 'fiscal_year', year.id); db.session.commit()
                except (ValueError, KeyError) as exc: flash(str(exc), 'error')
            elif action == 'toggle':
                year = db.get_or_404(FiscalYear, request.form.get('year_id', type=int)); year.is_open = not year.is_open
                audit('status', 'fiscal_year', year.id, str(year.is_open)); db.session.commit()
            return redirect(url_for('fiscal'))
        years = db.session.scalars(db.select(FiscalYear).order_by(FiscalYear.start_date.desc())).all()
        return render_template('fiscal.html', years=years)

    @app.route('/help', methods=['GET', 'POST'])
    @require('feedback')
    def help_page():
        if request.method == 'POST':
            subject = request.form.get('subject', '').strip(); message = request.form.get('message', '').strip()
            if not subject or not message: flash('Subject and description are required.', 'error')
            else:
                feedback = Feedback(user_id=g.user.id, category=request.form.get('category', 'Suggestion'), subject=subject, message=message)
                db.session.add(feedback); db.session.flush(); audit('create', 'feedback', feedback.id); db.session.commit()
                flash('Feedback sent.', 'success'); return redirect(url_for('help_page'))
        return render_template('help.html')

    @app.route('/admin/feedback', methods=['GET', 'POST'])
    @require('feedback_review')
    def feedback_review():
        if request.method == 'POST':
            item = db.get_or_404(Feedback, request.form.get('feedback_id', type=int))
            status = request.form.get('status', '')
            if status not in ('new', 'under review', 'resolved', 'declined'): abort(400)
            item.status = status; audit('status', 'feedback', item.id, status); db.session.commit()
            return redirect(url_for('feedback_review'))
        items = db.session.scalars(db.select(Feedback).order_by(Feedback.created_at.desc()).limit(200)).all()
        return render_template('feedback_review.html', items=items)

    @app.route('/admin/audit')
    @require('audit_view')
    def audit_view():
        events = db.session.scalars(db.select(Audit).order_by(Audit.timestamp.desc()).limit(200)).all()
        user_ids = {event.user_id for event in events if event.user_id}
        names = {user_id: db.session.get(User, user_id).name for user_id in user_ids if db.session.get(User, user_id)}
        return render_template('audit.html', events=events, names=names)

    @app.cli.command('init-db')
    def init_db():
        db.create_all()
        ensure_schema()
        if not db.session.scalar(db.select(Fund.id).limit(1)):
            db.session.add_all([Fund(name=n) for n in DEFAULT_FUNDS])
        if not db.session.scalar(db.select(FiscalYear.id).limit(1)):
            y = date.today().year; db.session.add(FiscalYear(name=str(y), start_date=date(y, 1, 1), end_date=date(y, 12, 31)))
        db.session.commit(); click.echo('Database initialized.')

    @app.cli.command('create-super-admin')
    @click.option('--email', prompt=True)
    @click.option('--first-name', prompt=True)
    @click.option('--last-name', prompt=True)
    @click.option('--phone', prompt=True)
    @click.password_option()
    def create_super_admin(email, first_name, last_name, phone, password):
        if not valid_email(email) or not all((first_name, last_name, phone)): raise click.ClickException('Provide a valid email, names, and phone.')
        if db.session.scalar(db.select(User).filter_by(email=email.lower())): raise click.ClickException('Email already exists.')
        user = User(email=email.lower(), first_name=first_name, last_name=last_name, phone=phone, role='Super Admin', password_hash=hash_password(password))
        db.session.add(user); db.session.commit(); click.echo('Super Admin created.')

    @app.cli.command('list-users')
    def list_users():
        users = db.session.scalars(db.select(User).order_by(User.last_name, User.first_name)).all()
        if not users: raise click.ClickException('No users exist. Run create-super-admin first.')
        for user in users:
            click.echo(f'{user.email}\t{user.name}\t{user.role}\t{"active" if user.active else "inactive"}')

    @app.cli.command('reset-password')
    @click.option('--email', prompt=True)
    @click.password_option()
    def reset_password(email, password):
        user = db.session.scalar(db.select(User).filter_by(email=email.strip().lower()))
        if not user: raise click.ClickException('No user has that email address.')
        if len(password) < 12: raise click.ClickException('A replacement password must contain at least 12 characters.')
        user.password_hash = hash_password(password)
        user.active = True
        db.session.execute(db.delete(LoginAttempt).where(LoginAttempt.email == user.email))
        db.session.add(Audit(user_id=user.id, action='password_reset_cli', entity='user', entity_id=user.id, summary='reset from command line'))
        db.session.commit()
        click.echo(f'Password reset for {user.email}.')

    return app

app = create_app()
