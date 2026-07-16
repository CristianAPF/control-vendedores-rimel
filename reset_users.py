from app import app, db, User, seed_data
from werkzeug.security import generate_password_hash

DEFAULTS = {
    'gerencia': ('Gerencia RIMEL', 'manager', None, 'Rimel2026!'),
    'gerson': ('Gerson', 'seller', 'Gerson', 'Gerson2026!'),
    'eduardo': ('Eduardo', 'seller', 'Eduardo', 'Eduardo2026!'),
    'victoria': ('Victoria', 'seller', 'Victoria', 'Victoria2026!'),
}

with app.app_context():
    db.create_all()
    seed_data()
    for username, (full_name, role, seller_name, password) in DEFAULTS.items():
        user = User.query.filter_by(username=username).first()
        if user is None:
            user = User(username=username)
            db.session.add(user)
        user.full_name = full_name
        user.role = role
        user.seller_name = seller_name
        user.active = True
        user.password_hash = generate_password_hash(password)
        user.must_change_password = False
    db.session.commit()
    print('Usuarios y contraseñas restablecidos correctamente.')
