import json, os, secrets
from datetime import datetime, date, timedelta, timezone
from io import BytesIO
from zoneinfo import ZoneInfo
from functools import wraps
from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import inspect, text, case, or_
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', secrets.token_hex(32))
db_url = os.environ.get('DATABASE_URL', f"sqlite:///{BASE/'rimel.db'}")
if db_url.startswith('postgres://'):
    db_url = db_url.replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_DATABASE_URI'] = db_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024
db = SQLAlchemy(app)

ALLOWED_RESULTS = [
    'PEDIDO CONCRETADO', 'SIN NECESIDAD DE REPOSICION', 'CLIENTE AUSENTE',
    'CLIENTE CERRADO', 'VISITA REPROGRAMADA', 'CLIENTE INACTIVO'
]
DAYS = ['LUNES','MARTES','MIERCOLES','JUEVES','VIERNES','SABADO','DOMINGO']
APP_VERSION = '2026.07.20.7'
RIVERA_TZ = ZoneInfo('America/Montevideo')

def utc_now_naive():
    return datetime.now(timezone.utc).replace(tzinfo=None)

def rivera_now():
    return datetime.now(RIVERA_TZ)

def to_rivera(value):
    if value is None:
        return None
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(RIVERA_TZ)

def local_day_utc_bounds(day_value):
    local_start = datetime.combine(day_value, datetime.min.time(), tzinfo=RIVERA_TZ)
    local_end = local_start + timedelta(days=1)
    return (local_start.astimezone(timezone.utc).replace(tzinfo=None),
            local_end.astimezone(timezone.utc).replace(tzinfo=None))

def local_range_utc_bounds(date_from, date_to):
    start, _ = local_day_utc_bounds(date_from)
    _, end = local_day_utc_bounds(date_to)
    return start, end

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(120), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='seller')
    seller_name = db.Column(db.String(120), nullable=True)
    active = db.Column(db.Boolean, default=True)
    must_change_password = db.Column(db.Boolean, default=True)

class Client(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(30), nullable=False, index=True)
    name = db.Column(db.String(200), nullable=False)
    address = db.Column(db.String(300), nullable=False)
    seller = db.Column(db.String(120), nullable=False, index=True)
    day = db.Column(db.String(20), nullable=False, index=True)
    original_order = db.Column(db.Integer, default=0)
    optimized_order = db.Column(db.Integer, nullable=True)
    latitude = db.Column(db.Float, nullable=True)
    longitude = db.Column(db.Float, nullable=True)
    active = db.Column(db.Boolean, default=True)

class Visit(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    client_id = db.Column(db.Integer, db.ForeignKey('client.id'), nullable=False)
    seller = db.Column(db.String(120), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    result = db.Column(db.String(80), nullable=False)
    notes = db.Column(db.Text, nullable=True)
    latitude = db.Column(db.Float, nullable=True)
    longitude = db.Column(db.Float, nullable=True)
    created_at = db.Column(db.DateTime, default=utc_now_naive, nullable=False, index=True)
    client = db.relationship('Client')
    user = db.relationship('User')


class LocationPing(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    seller = db.Column(db.String(120), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    latitude = db.Column(db.Float, nullable=False)
    longitude = db.Column(db.Float, nullable=False)
    accuracy = db.Column(db.Float, nullable=True)
    speed = db.Column(db.Float, nullable=True)
    heading = db.Column(db.Float, nullable=True)
    battery = db.Column(db.Float, nullable=True)
    created_at = db.Column(db.DateTime, default=utc_now_naive, nullable=False, index=True)
    user = db.relationship('User')

class Complaint(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    visit_id = db.Column(db.Integer, db.ForeignKey('visit.id'), nullable=False)
    client_id = db.Column(db.Integer, db.ForeignKey('client.id'), nullable=False)
    seller = db.Column(db.String(120), nullable=False, index=True)
    text = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(30), default='PENDIENTE')
    category = db.Column(db.String(60), default='OTRA')
    priority = db.Column(db.String(20), default='MEDIA')
    resolution = db.Column(db.Text, nullable=True)
    resolved_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, default=utc_now_naive, nullable=False)
    visit = db.relationship('Visit')
    client = db.relationship('Client')


def current_user():
    uid = session.get('user_id')
    return db.session.get(User, uid) if uid else None

@app.context_processor
def inject_globals():
    return {'current_user': current_user(), 'allowed_results': ALLOWED_RESULTS, 'app_version': APP_VERSION}

@app.template_filter('rivera_dt')
def rivera_dt_filter(value, fmt='%d/%m/%Y %H:%M'):
    converted = to_rivera(value)
    return converted.strftime(fmt) if converted else ''

def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not current_user(): return redirect(url_for('login'))
        return fn(*args, **kwargs)
    return wrapper

def manager_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        u=current_user()
        if not u or u.role != 'manager':
            flash('Acceso restringido al perfil gerencial.', 'error')
            return redirect(url_for('dashboard'))
        return fn(*args, **kwargs)
    return wrapper

def normalize_day(v):
    return str(v or '').strip().upper().replace('É','E').replace('Á','A').replace('Í','I').replace('Ó','O').replace('Ú','U')


def ensure_schema_updates():
    """Aplica cambios de esquema simples sin borrar la base existente."""
    inspector = inspect(db.engine)
    if 'client' not in inspector.get_table_names():
        return
    statements = []
    client_columns = {c['name'] for c in inspector.get_columns('client')}
    if 'latitude' not in client_columns:
        statements.append('ALTER TABLE client ADD COLUMN latitude FLOAT')
    if 'longitude' not in client_columns:
        statements.append('ALTER TABLE client ADD COLUMN longitude FLOAT')
    if 'complaint' in inspector.get_table_names():
        complaint_columns = {c['name'] for c in inspector.get_columns('complaint')}
        if 'category' not in complaint_columns:
            statements.append("ALTER TABLE complaint ADD COLUMN category VARCHAR(60) DEFAULT 'OTRA'")
        if 'priority' not in complaint_columns:
            statements.append("ALTER TABLE complaint ADD COLUMN priority VARCHAR(20) DEFAULT 'MEDIA'")
        if 'resolution' not in complaint_columns:
            statements.append('ALTER TABLE complaint ADD COLUMN resolution TEXT')
        if 'resolved_at' not in complaint_columns:
            statements.append('ALTER TABLE complaint ADD COLUMN resolved_at TIMESTAMP')
    for statement in statements:
        db.session.execute(text(statement))
    if statements:
        db.session.commit()

def route_order_expression():
    return case({day: idx for idx, day in enumerate(DAYS, start=1)}, value=Client.day, else_=99)

def next_client_code():
    max_id = db.session.query(db.func.max(Client.id)).scalar() or 0
    return f'CLI-{max_id + 1:06d}'

def next_route_order(seller, day):
    current = db.session.query(db.func.max(db.func.coalesce(Client.optimized_order, Client.original_order))).filter(
        Client.seller == seller, Client.day == day
    ).scalar() or 0
    return int(current) + 1

def seed_data():
    if Client.query.count() == 0:
        with open(BASE/'routes_embedded.json', encoding='utf-8') as f: data=json.load(f)
        for r in data['routes']:
            db.session.add(Client(code=r['provisional_code'], name=r['name'], address=r['address'], seller=r['seller'].title(), day=normalize_day(r['day']), original_order=r.get('approved_order') or 0))
    if User.query.count() == 0:
        defaults = [
            ('gerencia','Gerencia RIMEL','manager',None,'Rimel2026!'),
            ('gerson','Gerson','seller','Gerson','Gerson2026!'),
            ('eduardo','Eduardo','seller','Eduardo','Eduardo2026!'),
            ('victoria','Victoria','seller','Victoria','Victoria2026!')]
        for username,full,role,seller,pw in defaults:
            db.session.add(User(username=username,password_hash=generate_password_hash(pw),full_name=full,role=role,seller_name=seller,must_change_password=False))
    db.session.commit()

@app.before_request
def ensure_db():
    if not getattr(app, '_initialized', False):
        db.create_all(); ensure_schema_updates(); seed_data(); app._initialized=True

@app.get('/health')
def health():
    return {'status':'ok','version':APP_VERSION,'timezone':'America/Montevideo','server_time_rivera':rivera_now().isoformat()}

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method=='POST':
        u=User.query.filter_by(username=request.form.get('username','').strip().lower(), active=True).first()
        if not u or not check_password_hash(u.password_hash, request.form.get('password','')):
            flash('Usuario o contraseña incorrectos.', 'error')
        else:
            session.clear(); session['user_id']=u.id
            return redirect(url_for('change_password') if u.must_change_password else url_for('dashboard'))
    return render_template('login.html')

@app.get('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

@app.route('/change-password', methods=['GET','POST'])
@login_required
def change_password():
    u=current_user()
    if request.method=='POST':
        p=request.form.get('password',''); c=request.form.get('confirm','')
        if len(p)<8: flash('La contraseña debe tener al menos 8 caracteres.','error')
        elif p!=c: flash('Las contraseñas no coinciden.','error')
        else:
            u.password_hash=generate_password_hash(p); u.must_change_password=False; db.session.commit(); flash('Contraseña actualizada.','success'); return redirect(url_for('dashboard'))
    return render_template('change_password.html')

@app.get('/')
@login_required
def dashboard():
    u=current_user()
    seller_filter=request.args.get('seller','').strip()
    day_filter=normalize_day(request.args.get('day',''))
    date_from_raw=request.args.get('date_from','').strip()
    date_to_raw=request.args.get('date_to','').strip()
    try: date_from=date.fromisoformat(date_from_raw) if date_from_raw else date.today()
    except ValueError: date_from=date.today()
    try: date_to=date.fromisoformat(date_to_raw) if date_to_raw else date.today()
    except ValueError: date_to=date.today()
    if date_to < date_from: date_from, date_to = date_to, date_from
    start_dt, end_dt = local_range_utc_bounds(date_from, date_to)

    cq=Client.query.filter_by(active=True)
    if u.role=='seller':
        seller_filter=u.seller_name
        cq=cq.filter_by(seller=u.seller_name)
    elif seller_filter:
        cq=cq.filter_by(seller=seller_filter)
    if day_filter: cq=cq.filter_by(day=day_filter)
    clients=cq.order_by(route_order_expression(), db.func.coalesce(Client.optimized_order, Client.original_order), Client.name).all()
    client_ids=[c.id for c in clients]

    vq=Visit.query.filter(Visit.created_at>=start_dt, Visit.created_at<end_dt)
    if u.role=='seller': vq=vq.filter_by(seller=u.seller_name)
    elif seller_filter: vq=vq.filter_by(seller=seller_filter)
    if day_filter and client_ids: vq=vq.filter(Visit.client_id.in_(client_ids))
    elif day_filter and not client_ids: vq=vq.filter(db.text('1=0'))
    period_visits=vq.all()

    total=len(clients)
    visits=len(period_visits)
    unique=len({v.client_id for v in period_visits})
    orders=sum(v.result=='PEDIDO CONCRETADO' for v in period_visits)
    no_restock=sum(v.result=='SIN NECESIDAD DE REPOSICION' for v in period_visits)
    absent=sum(v.result=='CLIENTE AUSENTE' for v in period_visits)
    closed=sum(v.result=='CLIENTE CERRADO' for v in period_visits)
    reprogrammed=sum(v.result=='VISITA REPROGRAMADA' for v in period_visits)
    inactive=sum(v.result=='CLIENTE INACTIVO' for v in period_visits)
    effective=orders+no_restock
    contacted=effective+reprogrammed
    gps_count=sum(v.latitude is not None and v.longitude is not None for v in period_visits)
    notes_count=sum(bool(v.notes) for v in period_visits)
    repeat_visits=max(visits-unique,0)
    period_days=max((date_to-date_from).days+1,1)

    complaints_q=Complaint.query.filter(Complaint.created_at>=start_dt, Complaint.created_at<end_dt)
    if u.role=='seller': complaints_q=complaints_q.filter_by(seller=u.seller_name)
    elif seller_filter: complaints_q=complaints_q.filter_by(seller=seller_filter)
    complaints_total=complaints_q.count()
    complaints_pending=complaints_q.filter_by(status='PENDIENTE').count()

    active_sellers_q=db.session.query(Visit.seller).filter(Visit.created_at>=start_dt, Visit.created_at<end_dt).distinct()
    if seller_filter: active_sellers_q=active_sellers_q.filter(Visit.seller==seller_filter)
    active_sellers=active_sellers_q.count()

    metrics={
      'planned':total,'visits':visits,'unique':unique,
      'compliance':round(unique/total*100,1) if total else 0,
      'orders':orders,'conversion':round(orders/effective*100,1) if effective else 0,
      'contact_rate':round(contacted/visits*100,1) if visits else 0,
      'no_restock':no_restock,'absent':absent,'closed':closed,'reprogrammed':reprogrammed,'inactive':inactive,
      'gps_rate':round(gps_count/visits*100,1) if visits else 0,
      'notes_rate':round(notes_count/visits*100,1) if visits else 0,
      'repeat_visits':repeat_visits,'avg_daily':round(visits/period_days,1),
      'complaints':complaints_total,'complaints_pending':complaints_pending,
      'complaint_rate':round(complaints_total/visits*100,1) if visits else 0,
      'active_sellers':active_sellers
    }
    sellers=sorted([x[0] for x in db.session.query(Client.seller).distinct().all()])

    seller_stats=[]
    if u.role=='manager':
        for seller in sellers:
            sq=Visit.query.filter(Visit.seller==seller, Visit.created_at>=start_dt, Visit.created_at<end_dt)
            sv=sq.all(); su=len({v.client_id for v in sv}); so=sum(v.result=='PEDIDO CONCRETADO' for v in sv)
            se=sum(v.result in ('PEDIDO CONCRETADO','SIN NECESIDAD DE REPOSICION') for v in sv)
            sp=Client.query.filter_by(active=True,seller=seller)
            if day_filter: sp=sp.filter_by(day=day_filter)
            planned=sp.count()
            latest=LocationPing.query.filter_by(seller=seller).order_by(LocationPing.created_at.desc()).first()
            seller_stats.append({'seller':seller,'planned':planned,'visits':len(sv),'unique':su,
              'compliance':round(su/planned*100,1) if planned else 0,'orders':so,
              'conversion':round(so/se*100,1) if se else 0,'last_seen':latest.created_at if latest else None})

    return render_template('dashboard.html', clients=clients, metrics=metrics, sellers=sellers,
      selected_seller=seller_filter if u.role=='manager' else '', selected_day=day_filter,
      date_from=date_from.isoformat(),date_to=date_to.isoformat(),seller_stats=seller_stats)

@app.post('/visits')
@login_required
def create_visit():
    u=current_user(); client=db.session.get(Client, int(request.form['client_id']))
    if not client or (u.role=='seller' and client.seller!=u.seller_name): return ('No autorizado',403)
    result=request.form.get('result','').strip().upper()
    if result not in ALLOWED_RESULTS: flash('Resultado inválido.','error'); return redirect(url_for('dashboard'))
    def fnum(v):
        try:return float(v)
        except:return None
    visit=Visit(client_id=client.id,seller=client.seller,user_id=u.id,result=result,notes=request.form.get('notes','').strip() or None,latitude=fnum(request.form.get('latitude')),longitude=fnum(request.form.get('longitude')))
    db.session.add(visit); db.session.flush()
    complaint=request.form.get('complaint','').strip()
    if complaint:
        db.session.add(Complaint(visit_id=visit.id, client_id=client.id, seller=client.seller, text=complaint,
                                 category=request.form.get('complaint_category','OTRA'),
                                 priority=request.form.get('complaint_priority','MEDIA')))
    db.session.commit(); flash('Visita registrada y sincronizada con Gerencia.','success'); return redirect(url_for('dashboard'))

@app.get('/history')
@login_required
def history():
    u = current_user()
    seller = request.args.get('seller','').strip()
    result = request.args.get('result','').strip()
    date_from_raw = request.args.get('date_from','').strip()
    date_to_raw = request.args.get('date_to','').strip()
    q = Visit.query
    if u.role == 'seller':
        seller = u.seller_name
        q = q.filter_by(seller=u.seller_name)
    elif seller:
        q = q.filter_by(seller=seller)
    if result:
        q = q.filter_by(result=result)
    if date_from_raw:
        try:
            start_dt, _ = local_day_utc_bounds(date.fromisoformat(date_from_raw)); q=q.filter(Visit.created_at>=start_dt)
        except ValueError: pass
    if date_to_raw:
        try:
            _, end_dt = local_day_utc_bounds(date.fromisoformat(date_to_raw)); q=q.filter(Visit.created_at<end_dt)
        except ValueError: pass
    visits=q.order_by(Visit.created_at.desc()).limit(2000).all()
    sellers=sorted([x[0] for x in db.session.query(Visit.seller).distinct().all()])
    return render_template('history.html', visits=visits, sellers=sellers, selected_seller=seller,
                           selected_result=result, date_from=date_from_raw, date_to=date_to_raw)

@app.get('/complaints')
@login_required
def complaints():
    u=current_user(); seller=request.args.get('seller','').strip(); status=request.args.get('status','').strip()
    priority=request.args.get('priority','').strip(); q=Complaint.query
    if u.role=='seller': seller=u.seller_name; q=q.filter_by(seller=u.seller_name)
    elif seller: q=q.filter_by(seller=seller)
    if status: q=q.filter_by(status=status)
    if priority: q=q.filter_by(priority=priority)
    rows=q.order_by(Complaint.created_at.desc()).limit(2000).all()
    sellers=sorted([x[0] for x in db.session.query(Complaint.seller).distinct().all()])
    return render_template('complaints.html', complaints=rows, sellers=sellers,
                           selected_seller=seller, selected_status=status, selected_priority=priority)

@app.post('/complaints/<int:cid>/status')
@manager_required
def complaint_status(cid):
    c=db.session.get(Complaint,cid)
    if c:
        c.status=request.form.get('status','PENDIENTE')
        c.category=request.form.get('category',c.category or 'OTRA')
        c.priority=request.form.get('priority',c.priority or 'MEDIA')
        c.resolution=request.form.get('resolution','').strip() or None
        c.resolved_at=utc_now_naive() if c.status=='RESUELTA' else None
        db.session.commit()
    return redirect(url_for('complaints'))


def excel_response(filename, sheet_title, headers, rows):
    wb=Workbook(); ws=wb.active; ws.title=sheet_title[:31]
    ws.append(headers)
    fill=PatternFill('solid', fgColor='0876A8')
    for cell in ws[1]:
        cell.font=Font(color='FFFFFF',bold=True); cell.fill=fill; cell.alignment=Alignment(horizontal='center')
    for row in rows: ws.append(row)
    ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
    for col in range(1, ws.max_column+1):
        max_len=max(len(str(ws.cell(r,col).value or '')) for r in range(1, min(ws.max_row,1000)+1))
        ws.column_dimensions[get_column_letter(col)].width=min(max(max_len+2,10),55)
    output=BytesIO(); wb.save(output); output.seek(0)
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.get('/export/routes.xlsx')
@login_required
def export_routes():
    u=current_user(); seller=request.args.get('seller','').strip(); day_filter=normalize_day(request.args.get('day',''))
    q=Client.query.filter_by(active=True)
    if u.role=='seller': q=q.filter_by(seller=u.seller_name)
    elif seller: q=q.filter_by(seller=seller)
    if day_filter: q=q.filter_by(day=day_filter)
    clients=q.order_by(Client.seller,route_order_expression(),db.func.coalesce(Client.optimized_order,Client.original_order)).all()
    rows=[[c.seller,c.day,c.optimized_order or c.original_order,c.code,c.name,c.address,c.latitude,c.longitude] for c in clients]
    return excel_response('rutas_rimel.xlsx','Rutas',['Vendedor','Día','Orden','Código','Cliente','Dirección','Latitud','Longitud'],rows)

@app.get('/export/history.xlsx')
@login_required
def export_history():
    u=current_user(); seller=request.args.get('seller','').strip(); result=request.args.get('result','').strip()
    q=Visit.query
    if u.role=='seller': q=q.filter_by(seller=u.seller_name)
    elif seller: q=q.filter_by(seller=seller)
    if result: q=q.filter_by(result=result)
    df=request.args.get('date_from','').strip(); dt=request.args.get('date_to','').strip()
    if df:
        try: start,_=local_day_utc_bounds(date.fromisoformat(df)); q=q.filter(Visit.created_at>=start)
        except ValueError: pass
    if dt:
        try: _,end=local_day_utc_bounds(date.fromisoformat(dt)); q=q.filter(Visit.created_at<end)
        except ValueError: pass
    visits=q.order_by(Visit.created_at.desc()).all()
    rows=[]
    for v in visits:
        local=to_rivera(v.created_at)
        rows.append([v.id,local.strftime('%d/%m/%Y') if local else '',local.strftime('%H:%M:%S') if local else '',v.seller,v.client.code,v.client.name,v.client.address,v.result,v.notes or '',v.latitude,v.longitude])
    return excel_response('historial_visitas_rimel.xlsx','Historial',['ID','Fecha Rivera','Hora Rivera','Vendedor','Código','Cliente','Dirección','Resultado','Observaciones','Latitud','Longitud'],rows)

@app.get('/export/complaints.xlsx')
@login_required
def export_complaints():
    u=current_user(); seller=request.args.get('seller','').strip(); status=request.args.get('status','').strip(); priority=request.args.get('priority','').strip()
    q=Complaint.query
    if u.role=='seller': q=q.filter_by(seller=u.seller_name)
    elif seller: q=q.filter_by(seller=seller)
    if status: q=q.filter_by(status=status)
    if priority: q=q.filter_by(priority=priority)
    rows=[]
    for c in q.order_by(Complaint.created_at.desc()).all():
        created=to_rivera(c.created_at); resolved=to_rivera(c.resolved_at)
        rows.append([c.id,created.strftime('%d/%m/%Y %H:%M') if created else '',c.seller,c.client.code,c.client.name,c.category or 'OTRA',c.priority or 'MEDIA',c.text,c.status,c.resolution or '',resolved.strftime('%d/%m/%Y %H:%M') if resolved else ''])
    return excel_response('quejas_rimel.xlsx','Quejas',['ID','Fecha y hora Rivera','Vendedor','Código','Cliente','Categoría','Prioridad','Queja','Estado','Resolución','Cierre'],rows)


@app.get('/clients')
@manager_required
def clients_admin():
    seller = request.args.get('seller', '').strip()
    day = normalize_day(request.args.get('day', ''))
    status = request.args.get('status', 'active').strip()
    search = request.args.get('q', '').strip()
    q = Client.query
    if seller:
        q = q.filter(Client.seller == seller)
    if day:
        q = q.filter(Client.day == day)
    if status == 'active':
        q = q.filter(Client.active.is_(True))
    elif status == 'inactive':
        q = q.filter(Client.active.is_(False))
    if search:
        like = f'%{search}%'
        q = q.filter(or_(Client.name.ilike(like), Client.address.ilike(like), Client.code.ilike(like)))
    rows = q.order_by(route_order_expression(), Client.seller, db.func.coalesce(Client.optimized_order, Client.original_order), Client.name).all()
    sellers = sorted([x[0] for x in db.session.query(Client.seller).distinct().all()])
    return render_template('clients.html', clients=rows, sellers=sellers, selected_seller=seller,
                           selected_day=day, selected_status=status, search=search)

@app.post('/clients/create')
@manager_required
def client_create():
    seller = request.form.get('seller', '').strip().title()
    day = normalize_day(request.form.get('day', ''))
    name = request.form.get('name', '').strip()
    address = request.form.get('address', '').strip()
    code = request.form.get('code', '').strip().upper() or next_client_code()
    if not seller or day not in DAYS or not name or not address:
        flash('Completa vendedor, día, nombre y dirección.', 'error')
        return redirect(url_for('clients_admin'))
    if Client.query.filter_by(code=code).first():
        flash('Ya existe un cliente con ese código.', 'error')
        return redirect(url_for('clients_admin'))
    order = next_route_order(seller, day)
    c = Client(code=code, name=name, address=address, seller=seller, day=day,
               original_order=order, optimized_order=order, active=True)
    db.session.add(c); db.session.commit()
    flash('Cliente agregado. Quedó al final de la ruta; conviene volver a optimizar ese día.', 'success')
    return redirect(url_for('clients_admin', seller=seller, day=day))

@app.route('/clients/<int:client_id>/edit', methods=['GET', 'POST'])
@manager_required
def client_edit(client_id):
    c = db.session.get(Client, client_id)
    if not c:
        return ('Cliente no encontrado', 404)
    if request.method == 'POST':
        seller = request.form.get('seller', '').strip().title()
        day = normalize_day(request.form.get('day', ''))
        name = request.form.get('name', '').strip()
        address = request.form.get('address', '').strip()
        code = request.form.get('code', '').strip().upper()
        if not seller or day not in DAYS or not name or not address or not code:
            flash('Todos los campos son obligatorios.', 'error')
        elif Client.query.filter(Client.code == code, Client.id != c.id).first():
            flash('Ese código ya pertenece a otro cliente.', 'error')
        else:
            route_changed = c.seller != seller or c.day != day or c.address != address
            c.code, c.name, c.address, c.seller, c.day = code, name, address, seller, day
            if route_changed:
                c.optimized_order = next_route_order(seller, day)
                c.latitude = None; c.longitude = None
            db.session.commit()
            flash('Cliente actualizado.' + (' Reoptimiza la ruta modificada.' if route_changed else ''), 'success')
            return redirect(url_for('clients_admin', seller=seller, day=day))
    sellers = sorted([x[0] for x in db.session.query(Client.seller).distinct().all()])
    return render_template('client_edit.html', client=c, sellers=sellers)

@app.post('/clients/<int:client_id>/toggle')
@manager_required
def client_toggle(client_id):
    c = db.session.get(Client, client_id)
    if c:
        c.active = not c.active
        db.session.commit()
        flash('Cliente activado.' if c.active else 'Cliente dado de baja. Su historial se conserva.', 'success')
    return redirect(request.referrer or url_for('clients_admin'))

@app.get('/api/routes/<seller>/<day>')
@manager_required
def route_clients_api(seller, day):
    day = normalize_day(day)
    rows = Client.query.filter_by(seller=seller, day=day, active=True).order_by(
        db.func.coalesce(Client.optimized_order, Client.original_order), Client.name
    ).all()
    return jsonify([{'id': c.id, 'name': c.name, 'address': c.address, 'latitude': c.latitude,
                     'longitude': c.longitude, 'order': c.optimized_order or c.original_order} for c in rows])

@app.post('/api/routes/save-optimized')
@manager_required
def save_optimized_route():
    data = request.get_json(silent=True) or {}
    seller = str(data.get('seller', '')).strip()
    day = normalize_day(data.get('day', ''))
    ordered = data.get('clients') or []
    if not seller or day not in DAYS or not ordered:
        return jsonify({'error': 'Datos incompletos'}), 400
    valid = {c.id: c for c in Client.query.filter_by(seller=seller, day=day, active=True).all()}
    received_ids = []
    for position, item in enumerate(ordered, start=1):
        try:
            cid = int(item.get('id'))
        except (TypeError, ValueError, AttributeError):
            continue
        c = valid.get(cid)
        if not c:
            continue
        c.optimized_order = position
        received_ids.append(cid)
        try:
            c.latitude = float(item['latitude']) if item.get('latitude') is not None else c.latitude
            c.longitude = float(item['longitude']) if item.get('longitude') is not None else c.longitude
        except (TypeError, ValueError):
            pass
    # Los no geocodificados quedan al final, conservando su orden anterior.
    remaining = [c for cid, c in valid.items() if cid not in received_ids]
    remaining.sort(key=lambda c: (c.optimized_order or c.original_order or 999999, c.name))
    pos = len(received_ids)
    for c in remaining:
        pos += 1; c.optimized_order = pos
    db.session.commit()
    return jsonify({'ok': True, 'saved': len(received_ids), 'total': len(valid)})

@app.route('/users', methods=['GET','POST'])
@manager_required
def users():
    if request.method=='POST':
        username=request.form['username'].strip().lower()
        if User.query.filter_by(username=username).first(): flash('El usuario ya existe.','error')
        else:
            role=request.form['role']; seller=request.form.get('seller_name','').strip() or None
            db.session.add(User(username=username,full_name=request.form['full_name'].strip(),role=role,seller_name=seller,password_hash=generate_password_hash(request.form['password']),must_change_password=False)); db.session.commit(); flash('Usuario creado.','success')
    sellers=sorted([x[0] for x in db.session.query(Client.seller).distinct().all()])
    return render_template('users.html', users=User.query.order_by(User.role,User.full_name).all(), sellers=sellers)

@app.post('/users/<int:uid>/reset')
@manager_required
def reset_user(uid):
    u=db.session.get(User,uid); pw=request.form.get('password','Rimel2026!')
    if u: u.password_hash=generate_password_hash(pw); u.must_change_password=True; db.session.commit(); flash('Contraseña restablecida.','success')
    return redirect(url_for('users'))

@app.get('/api/summary')
@login_required
def api_summary():
    u=current_user(); since,_=local_day_utc_bounds(rivera_now().date()); q=Visit.query.filter(Visit.created_at>=since)
    if u.role=='seller': q=q.filter_by(seller=u.seller_name)
    rows=q.all(); return jsonify({'visits':len(rows),'orders':sum(v.result=='PEDIDO CONCRETADO' for v in rows),'last_update':rivera_now().isoformat()})


@app.post('/api/location')
@login_required
def save_location():
    u=current_user()
    if u.role!='seller' or not u.seller_name: return jsonify({'error':'Solo vendedores'}),403
    data=request.get_json(silent=True) or {}
    try:
        lat=float(data['latitude']); lon=float(data['longitude'])
    except (KeyError,TypeError,ValueError): return jsonify({'error':'Coordenadas inválidas'}),400
    def val(name):
        try: return float(data[name]) if data.get(name) is not None else None
        except (TypeError,ValueError): return None
    ping=LocationPing(seller=u.seller_name,user_id=u.id,latitude=lat,longitude=lon,
        accuracy=val('accuracy'),speed=val('speed'),heading=val('heading'),battery=val('battery'))
    db.session.add(ping); db.session.commit()
    return jsonify({'ok':True,'saved_at':ping.created_at.isoformat()})

@app.get('/tracking')
@manager_required
def tracking():
    sellers=sorted([x[0] for x in db.session.query(Client.seller).distinct().all()])
    return render_template('tracking.html', sellers=sellers)

@app.get('/api/locations')
@manager_required
def api_locations():
    sellers=[x[0] for x in db.session.query(Client.seller).distinct().all()]
    result=[]
    now=utc_now_naive()
    for seller in sellers:
        p=LocationPing.query.filter_by(seller=seller).order_by(LocationPing.created_at.desc()).first()
        if p:
            age=(now-p.created_at).total_seconds()
            result.append({'seller':seller,'latitude':p.latitude,'longitude':p.longitude,
              'accuracy':p.accuracy,'speed':p.speed,'heading':p.heading,
              'created_at':p.created_at.isoformat()+'Z',
              'local_time':to_rivera(p.created_at).strftime('%d/%m/%Y %H:%M:%S'),
              'age_seconds':round(age),'online':age<=180})
        else:
            result.append({'seller':seller,'online':False})
    return jsonify(result)

@app.get('/client/<int:client_id>/map')
@login_required
def client_map(client_id):
    u=current_user(); c=db.session.get(Client,client_id)
    if not c or (u.role=='seller' and c.seller!=u.seller_name): return ('No autorizado',403)
    return redirect('https://www.google.com/maps/search/?api=1&query='+__import__('urllib.parse').parse.quote_plus(f"{c.address}, Rivera, Uruguay"))

@app.get('/service-worker.js')
def service_worker(): return send_file(BASE/'static'/'sw.js', mimetype='application/javascript')

if __name__=='__main__':
    with app.app_context(): db.create_all(); ensure_schema_updates(); seed_data()
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT',8000)), debug=False, use_reloader=False)
