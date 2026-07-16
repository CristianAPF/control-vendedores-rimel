import json, os, secrets
from datetime import datetime, date
from functools import wraps
from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash

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
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)
    client = db.relationship('Client')
    user = db.relationship('User')

class Complaint(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    visit_id = db.Column(db.Integer, db.ForeignKey('visit.id'), nullable=False)
    client_id = db.Column(db.Integer, db.ForeignKey('client.id'), nullable=False)
    seller = db.Column(db.String(120), nullable=False, index=True)
    text = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(30), default='PENDIENTE')
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    visit = db.relationship('Visit')
    client = db.relationship('Client')


def current_user():
    uid = session.get('user_id')
    return db.session.get(User, uid) if uid else None

@app.context_processor
def inject_globals():
    return {'current_user': current_user(), 'allowed_results': ALLOWED_RESULTS}

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
        db.create_all(); seed_data(); app._initialized=True

@app.get('/health')
def health(): return {'status':'ok'}

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
    u=current_user(); seller_filter=request.args.get('seller','').strip(); day_filter=normalize_day(request.args.get('day',''))
    cq=Client.query.filter_by(active=True)
    vq=Visit.query
    if u.role=='seller':
        cq=cq.filter_by(seller=u.seller_name); vq=vq.filter_by(seller=u.seller_name)
    elif seller_filter:
        cq=cq.filter_by(seller=seller_filter); vq=vq.filter_by(seller=seller_filter)
    if day_filter: cq=cq.filter_by(day=day_filter)
    clients=cq.order_by(Client.day, db.func.coalesce(Client.optimized_order, Client.original_order), Client.name).all()
    today_start=datetime.combine(date.today(), datetime.min.time())
    today_visits=vq.filter(Visit.created_at>=today_start).all()
    total=len(clients); visits=len(today_visits); unique=len({v.client_id for v in today_visits}); orders=sum(v.result=='PEDIDO CONCRETADO' for v in today_visits)
    effective=sum(v.result in ('PEDIDO CONCRETADO','SIN NECESIDAD DE REPOSICION') for v in today_visits)
    complaints=Complaint.query.filter_by(status='PENDIENTE')
    if u.role=='seller': complaints=complaints.filter_by(seller=u.seller_name)
    metrics={'planned':total,'visits':visits,'unique':unique,'compliance':round(unique/total*100,1) if total else 0,'orders':orders,'conversion':round(orders/effective*100,1) if effective else 0,'contact_rate':round(effective/visits*100,1) if visits else 0,'complaints':complaints.count()}
    sellers=sorted([x[0] for x in db.session.query(Client.seller).distinct().all()])
    return render_template('dashboard.html', clients=clients, metrics=metrics, sellers=sellers, selected_seller=seller_filter, selected_day=day_filter)

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
    if complaint: db.session.add(Complaint(visit_id=visit.id,client_id=client.id,seller=client.seller,text=complaint))
    db.session.commit(); flash('Visita registrada y sincronizada con Gerencia.','success'); return redirect(url_for('dashboard'))

@app.get('/history')
@login_required
def history():
    u=current_user(); q=Visit.query.order_by(Visit.created_at.desc())
    if u.role=='seller': q=q.filter_by(seller=u.seller_name)
    visits=q.limit(500).all(); return render_template('history.html', visits=visits)

@app.get('/complaints')
@login_required
def complaints():
    u=current_user(); q=Complaint.query.order_by(Complaint.created_at.desc())
    if u.role=='seller': q=q.filter_by(seller=u.seller_name)
    return render_template('complaints.html', complaints=q.limit(500).all())

@app.post('/complaints/<int:cid>/status')
@manager_required
def complaint_status(cid):
    c=db.session.get(Complaint,cid)
    if c: c.status=request.form.get('status','PENDIENTE'); db.session.commit()
    return redirect(url_for('complaints'))

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
    u=current_user(); since=datetime.combine(date.today(), datetime.min.time()); q=Visit.query.filter(Visit.created_at>=since)
    if u.role=='seller': q=q.filter_by(seller=u.seller_name)
    rows=q.all(); return jsonify({'visits':len(rows),'orders':sum(v.result=='PEDIDO CONCRETADO' for v in rows),'last_update':datetime.utcnow().isoformat()})

@app.get('/service-worker.js')
def service_worker(): return send_file(BASE/'static'/'sw.js', mimetype='application/javascript')

if __name__=='__main__':
    with app.app_context(): db.create_all(); seed_data()
    app.run(host='0.0.0.0',port=int(os.environ.get('PORT',8000)),debug=True)
