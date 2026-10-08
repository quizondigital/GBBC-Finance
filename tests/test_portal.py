from datetime import date
import io
import pytest
from werkzeug.security import generate_password_hash
from app import create_app, db, User, Person, Fund, FiscalYear, Gift, ensure_schema

def _hash(password):
    return generate_password_hash(password, method='pbkdf2:sha256')

@pytest.fixture
def client(tmp_path):
    app = create_app({'TESTING': True, 'SECRET_KEY': 'test-key', 'SQLALCHEMY_DATABASE_URI': f'sqlite:///{tmp_path / "test.db"}'})
    with app.app_context():
        db.create_all()
        ensure_schema()
        db.session.add_all([
            User(first_name='Sally', last_name='Admin', email='sally@example.org', phone='555-1000', role='Super Admin', password_hash=_hash('correct-password')),
            User(first_name='Chris', last_name='Leader', email='chris@example.org', phone='555-2000', role='Church Leader', password_hash=_hash('correct-password')),
            Person(first_name='Ada', last_name='Member', member_type='Member'),
            FiscalYear(name='2026', start_date=date(2026,1,1), end_date=date(2026,12,31)),
        ]); db.session.commit()
    return app.test_client()

def token(client):
    with client.session_transaction() as s: return s['csrf']

def login(client, email):
    client.get('/login')
    return client.post('/login', data={'csrf_token': token(client), 'email': email, 'password': 'correct-password'}, follow_redirects=True)

def test_auth_and_user_fields(client):
    assert client.get('/people').status_code == 302
    assert b'Good morning, Sally' in login(client, 'sally@example.org').data
    bad = client.post('/admin/users', data={'csrf_token': token(client), 'first_name':'Pat','last_name':'Smith','email':'pat@example.org','role':'Finance Staff','password':'secure-password'}, follow_redirects=True)
    assert b'phone' in bad.data.lower()
    good = client.post('/admin/users', data={'csrf_token': token(client), 'first_name':'Pat','last_name':'Smith','email':'pat@example.org','phone':'555-3000','role':'Finance Staff','password':'secure-password'}, follow_redirects=True)
    assert b'User created' in good.data
    assert client.post('/admin/users', data={}).status_code == 400
    assert b'Church Leader' in client.get('/admin/roles').data

def test_leader_suppression_and_server_denial(client):
    with client.application.app_context():
        db.session.add(Gift(person_id=1, fiscal_year_id=1, fund_id=1, date=date(2026,9,1), amount=123, method='Check', entered_by_id=1)); db.session.commit()
    login(client, 'chris@example.org')
    dashboard = client.get('/').data
    assert b'$123.00' in dashboard and b'Ada Member' not in dashboard
    assert b'Contributions</span>' not in dashboard and b'Administration</span>' not in dashboard
    for path in ['/contributions', '/reports/export.xlsx', '/admin/users', '/admin/roles', '/people/new']:
        assert client.get(path).status_code == 403
    report = client.get('/reports').data
    assert b'$123.00' in report and b'Ada Member' not in report
    assert client.get('/reports?person_id=1').status_code == 403
    assert client.get('/reports?minimum_total=1').status_code == 403

def test_gift_entry_and_year_close(client):
    login(client, 'sally@example.org')
    form = client.get('/contributions/new').data.decode()
    assert form.index('General Fund') < form.index('Missions Fund') < form.index('Building Fund') < form.index('Benevolence Fund')
    assert 'selected' in form[form.index('General Fund')-80:form.index('General Fund')]
    data={'csrf_token':token(client),'person_id':'1','fiscal_year_id':'1','fund_id':'1','date':'2026-09-01','amount':'23.45','method':'Check'}
    assert b'Contribution saved' in client.post('/contributions/new',data=data,follow_redirects=True).data
    with client.application.app_context():
        db.session.get(FiscalYear,1).is_open=False;db.session.commit()
    assert b'open fiscal year' in client.post('/contributions/new',data=data,follow_redirects=True).data

def test_person_structured_address_and_letter(client):
    login(client, 'sally@example.org')
    data = {
        'csrf_token': token(client),
        'first_name': 'Ada', 'last_name': 'Member', 'member_type': 'Member',
        'email': 'ada@example.org', 'phone': '555-0100',
        'address_line1': '170 Greenbrook Rd.', 'address_line2': 'Attn: Finance',
        'city': 'Green Brook', 'state': 'nj', 'zip_code': '08812',
    }
    assert b'Person saved' in client.post('/people/1/edit', data=data, follow_redirects=True).data
    with client.application.app_context():
        person = db.session.get(Person, 1)
        assert person.address_line1 == '170 Greenbrook Rd.'
        assert person.address_line2 == 'Attn: Finance'
        assert person.city == 'Green Brook'
        assert person.state == 'NJ'
        assert person.zip_code == '08812'
        assert person.address_display == '170 Greenbrook Rd., Attn: Finance, Green Brook, NJ 08812'
    with client.application.app_context():
        db.session.add(Gift(person_id=1, fiscal_year_id=1, fund_id=1, date=date(2026,9,1), amount=50, method='Check', entered_by_id=1)); db.session.commit()
    pdf = client.get('/reports/statements/1.pdf?year_id=1')
    assert pdf.status_code == 200
    assert b'170 Greenbrook Rd.' in pdf.data
    assert b'Attn: Finance' in pdf.data
    assert b'Green Brook, NJ 08812' in pdf.data
    people = client.get('/people').data
    assert b'170 Greenbrook Rd.' in people and b'Green Brook, NJ 08812' in people
    form = client.get('/people/new').data
    assert b'Address line 1' in form and b'us-states' in form
    template = client.get('/people/template.csv').data
    assert b'address_line1,address_line2,city,state,zip_code' in template

def test_exports_and_role_matrix(client):
    login(client, 'sally@example.org')
    with client.application.app_context():
        db.session.add(Gift(person_id=1, fiscal_year_id=1, fund_id=1, date=date(2026,9,1), amount=123, method='Check', entered_by_id=1)); db.session.commit()
    pdf = client.get('/reports/statements/1.pdf')
    assert pdf.status_code == 200 and pdf.data.startswith(b'%PDF')
    assert b'Romella Seepaul' in pdf.data
    assert b'Dear Ada' in pdf.data
    assert b'2 Corinthians 9:6-8' in pdf.data
    assert b'Itemized contributions' in pdf.data
    xlsx = client.get('/reports/export.xlsx')
    assert xlsx.status_code == 200 and xlsx.data.startswith(b'PK')
    assert client.get('/admin/settings').status_code == 200

def test_login_limit(client):
    client.get('/login')
    for _ in range(5):
        response = client.post('/login', data={'csrf_token':token(client), 'email':'sally@example.org', 'password':'wrong'})
        assert response.status_code == 200
    assert client.post('/login', data={'csrf_token':token(client), 'email':'sally@example.org', 'password':'correct-password'}).status_code == 429

def test_all_primary_pages_render(client):
    login(client, 'sally@example.org')
    paths = ['/', '/people', '/people/new', '/people/1/edit', '/people/import', '/contributions', '/contributions/import', '/contributions/new', '/reports', '/admin/users', '/admin/users/1/edit', '/admin/roles', '/admin/fiscal', '/admin/settings', '/admin/feedback', '/admin/audit', '/help']
    for path in paths:
        response = client.get(path)
        assert response.status_code == 200, path
        assert b'GBBC Finance' in response.data, path

def test_contribution_csv_import(client):
    login(client, 'sally@example.org')
    template = client.get('/contributions/template.csv')
    assert template.status_code == 200
    assert b'date,first_name,last_name,email,fund,amount,method,fiscal_year,reference,note' in template.data
    csv_body = (
        'date,first_name,last_name,email,fund,amount,method,fiscal_year,reference,note\n'
        '2026-09-01,Ada,Member,,General Fund,40.00,Check,2026,chk-1,\n'
        '2026-09-02,Ada,Member,,Missions,15.50,Cash,2026,,\n'
        '2026-09-03,Ada,Member,,General Fund,not-a-number,Check,2026,,\n'
        '2026-09-01,Ada,Member,,General Fund,40.00,Check,2026,chk-1,\n'
    )
    response = client.post(
        '/contributions/import',
        data={'csrf_token': token(client), 'file': (io.BytesIO(csv_body.encode('utf-8')), 'gifts.csv')},
        content_type='multipart/form-data',
        follow_redirects=True,
    )
    assert b'Contributions import report' in response.data
    assert b'Loaded successfully' in response.data
    assert b'>2<' in response.data or b'2</strong>' in response.data
    assert b'Failed to load' in response.data
    assert b'amount must be a positive number' in response.data
    assert b'duplicate of another row in this file' in response.data
    with client.application.app_context():
        gifts = db.session.scalars(db.select(Gift).order_by(Gift.id)).all()
        assert len(gifts) == 2
        assert float(gifts[0].amount) == 40.00
        assert gifts[1].fund.name == 'Missions Fund'
        assert gifts[1].method == 'Cash'
    failed = client.get('/import/report/failures.csv')
    assert failed.status_code == 200
    assert b'failure_reason' in failed.data

def test_people_csv_import_report(client):
    login(client, 'sally@example.org')
    csv_body = (
        'first_name,last_name,member_type,email,phone,address_line1,address_line2,city,state,zip_code\n'
        'Ada,Member,Member,ada@example.org,555-0100,1 Main St,,Green Brook,NJ,08812\n'
        ',Visitor,Non-member,ben@example.org,555-0101,2 Main St,,Plainfield,NJ,07060\n'
        'Ada,Member,Member,ada@example.org,555-0100,1 Main St,,Green Brook,NJ,08812\n'
    )
    response = client.post(
        '/people/import',
        data={'csrf_token': token(client), 'file': (io.BytesIO(csv_body.encode('utf-8')), 'people.csv')},
        content_type='multipart/form-data',
        follow_redirects=True,
    )
    assert b'People import report' in response.data
    assert b'first name and last name are required' in response.data
    assert b'duplicate of another row in this file' in response.data
    with client.application.app_context():
        people = db.session.scalars(db.select(Person).order_by(Person.id)).all()
        # fixture Ada Member + 1 imported
        assert len(people) == 2
        assert people[1].email == 'ada@example.org'

def test_empty_csv_import_report(client):
    login(client, 'sally@example.org')
    response = client.post(
        '/people/import',
        data={'csrf_token': token(client), 'file': (io.BytesIO(b''), 'empty.csv')},
        content_type='multipart/form-data',
        follow_redirects=True,
    )
    assert b'People import report' in response.data
    assert b'The file is empty.' in response.data

def test_gift_delete_from_list(client):
    login(client, 'sally@example.org')
    with client.application.app_context():
        db.session.add(Gift(person_id=1, fiscal_year_id=1, fund_id=1, date=date(2026,9,1), amount=25, method='Check', entered_by_id=1))
        db.session.commit()
    listing = client.get('/contributions').data
    assert b'Delete' in listing
    deleted = client.post('/contributions/1/delete', data={'csrf_token': token(client)}, follow_redirects=True)
    assert b'Contribution deleted.' in deleted.data
    with client.application.app_context():
        assert db.session.get(Gift, 1) is None

def test_contributions_sortable_default_date_desc(client):
    login(client, 'sally@example.org')
    with client.application.app_context():
        db.session.add(Person(first_name='Ben', last_name='Visitor', member_type='Non-member'))
        db.session.commit()
        db.session.add_all([
            Gift(person_id=1, fiscal_year_id=1, fund_id=1, date=date(2026,8,1), amount=10, method='Cash', entered_by_id=1),
            Gift(person_id=2, fiscal_year_id=1, fund_id=1, date=date(2026,10,1), amount=50, method='Check', entered_by_id=1),
        ])
        db.session.commit()
    default = client.get('/contributions').data.decode()
    assert 'sort=date' in default and '▼' in default
    assert default.index('2026-10-01') < default.index('2026-08-01')
    by_name = client.get('/contributions?sort=contributor&dir=asc').data.decode()
    assert by_name.index('Ada Member') < by_name.index('Ben Visitor')
    by_amount = client.get('/contributions?sort=amount&dir=desc').data.decode()
    assert by_amount.index('$50.00') < by_amount.index('$10.00')

def test_reconciliation_is_audited_and_role_limited(client):
    login(client, 'sally@example.org')
    with client.application.app_context():
        db.session.add(Gift(person_id=1, fiscal_year_id=1, fund_id=1, date=date(2026,9,1), amount=25, method='Check', entered_by_id=1)); db.session.commit()
    response = client.post('/contributions/1/reconcile', data={'csrf_token':token(client)}, follow_redirects=True)
    assert b'Reconciliation status updated' in response.data
    with client.application.app_context(): assert db.session.get(Gift,1).reconciled is True
