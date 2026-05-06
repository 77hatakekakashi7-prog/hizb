"""
server.py — الحزب الاشتراكى Finance Server
قاعدة بيانات SQLite دائمة — بيانات كل يوزر محفوظة بشكل منفصل
"""
from flask import Flask, request, send_file, jsonify, make_response, send_from_directory
import io, os, json, sqlite3, hashlib, secrets
from datetime import datetime, timedelta
from functools import wraps

try:
    from gen_excel import build_workbook
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

app = Flask(__name__, static_folder='.')

# ─── إعداد قاعدة البيانات ───────────────────────────────────────────
DB_PATH = os.environ.get('DB_PATH', 'hizb_finance.db')

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn

def init_db():
    with get_db() as db:
        db.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            username  TEXT UNIQUE NOT NULL,
            password  TEXT NOT NULL,
            role      TEXT DEFAULT 'user',
            created   TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS sessions (
            token     TEXT PRIMARY KEY,
            user_id   INTEGER NOT NULL,
            expires   TEXT NOT NULL,
            FOREIGN KEY(user_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS transactions (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id   INTEGER NOT NULL,
            type      TEXT NOT NULL,
            amt       REAL NOT NULL,
            dsc       TEXT NOT NULL,
            cat       TEXT,
            dt        TEXT,
            note      TEXT,
            created   TEXT DEFAULT (datetime('now')),
            FOREIGN KEY(user_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS periods (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id   INTEGER NOT NULL,
            label     TEXT NOT NULL,
            inc       REAL DEFAULT 0,
            exp       REAL DEFAULT 0,
            bal       REAL DEFAULT 0,
            closed_at TEXT DEFAULT (datetime('now')),
            FOREIGN KEY(user_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS period_transactions (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id   INTEGER NOT NULL,
            type        TEXT,
            amt         REAL,
            dsc         TEXT,
            cat         TEXT,
            dt          TEXT,
            note        TEXT,
            FOREIGN KEY(period_id) REFERENCES periods(id)
        );

        CREATE TABLE IF NOT EXISTS goals (
            user_id   INTEGER PRIMARY KEY,
            amount    REAL DEFAULT 0,
            FOREIGN KEY(user_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS vault_folders (
            id        TEXT PRIMARY KEY,
            user_id   INTEGER NOT NULL,
            name      TEXT NOT NULL,
            icon      TEXT DEFAULT '📁',
            FOREIGN KEY(user_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS vault_accounts (
            id        TEXT PRIMARY KEY,
            user_id   INTEGER NOT NULL,
            folder_id TEXT,
            title     TEXT NOT NULL,
            username  TEXT,
            password  TEXT,
            email     TEXT,
            pin       TEXT,
            url       TEXT,
            note      TEXT,
            created_at INTEGER,
            FOREIGN KEY(user_id) REFERENCES users(id)
        );
        """)
        # إنشاء مستخدمين افتراضيين إن لم يكونوا موجودين
        for uname, upass, role in [('admin','hizb2024','admin'), ('محاسب','hizb2024','user')]:
            existing = db.execute("SELECT id FROM users WHERE username=?", (uname,)).fetchone()
            if not existing:
                hashed = hash_pass(upass)
                db.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)", (uname, hashed, role))
        db.commit()
    print(f"✓ قاعدة البيانات جاهزة: {DB_PATH}")

def hash_pass(p):
    return hashlib.sha256(p.encode()).hexdigest()

# ─── CORS ────────────────────────────────────────────────────────────
@app.after_request
def add_cors(response):
    origin = request.headers.get('Origin', '*')
    response.headers['Access-Control-Allow-Origin']      = origin
    response.headers['Access-Control-Allow-Methods']     = 'GET, POST, PUT, DELETE, OPTIONS'
    response.headers['Access-Control-Allow-Headers']     = 'Content-Type, Authorization, X-Auth-Token'
    response.headers['Access-Control-Allow-Credentials'] = 'true'
    return response

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve_static(path):
    if path and os.path.exists(os.path.join('.', path)):
        return send_from_directory('.', path)
    return send_from_directory('.', 'hizb_finance.html')

# ─── Auth helpers ─────────────────────────────────────────────────────
def require_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if request.method == 'OPTIONS':
            return make_response('', 200)
        token = request.headers.get('X-Auth-Token') or request.cookies.get('hizb_token')
        if not token:
            return jsonify({'error': 'غير مصرح'}), 401
        with get_db() as db:
            row = db.execute(
                "SELECT s.user_id, u.username, u.role FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token=? AND s.expires > datetime('now')",
                (token,)
            ).fetchone()
        if not row:
            return jsonify({'error': 'انتهت الجلسة، سجل دخولك مجدداً'}), 401
        request.user_id  = row['user_id']
        request.username = row['username']
        request.role     = row['role']
        return f(*args, **kwargs)
    return decorated

# ─── AUTH ENDPOINTS ───────────────────────────────────────────────────
@app.route('/api/login', methods=['POST', 'OPTIONS'])
def login():
    if request.method == 'OPTIONS': return make_response('', 200)
    data = request.get_json(force=True)
    uname = data.get('username','').strip()
    upass = data.get('password','')
    if not uname or not upass:
        return jsonify({'error': 'أدخل اسم المستخدم وكلمة المرور'}), 400
    with get_db() as db:
        user = db.execute("SELECT * FROM users WHERE username=?", (uname,)).fetchone()
    if not user or user['password'] != hash_pass(upass):
        return jsonify({'error': 'اسم المستخدم أو كلمة المرور غير صحيحة'}), 401
    token   = secrets.token_hex(32)
    expires = (datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d %H:%M:%S')
    with get_db() as db:
        db.execute("INSERT INTO sessions(token,user_id,expires) VALUES(?,?,?)", (token, user['id'], expires))
        db.commit()
    resp = jsonify({'token': token, 'username': uname, 'role': user['role']})
    resp.set_cookie('hizb_token', token, max_age=604800, httponly=True, samesite='Lax')
    return resp

@app.route('/api/logout', methods=['POST', 'OPTIONS'])
@require_auth
def logout():
    token = request.headers.get('X-Auth-Token') or request.cookies.get('hizb_token')
    with get_db() as db:
        db.execute("DELETE FROM sessions WHERE token=?", (token,))
        db.commit()
    resp = jsonify({'ok': True})
    resp.delete_cookie('hizb_token')
    return resp

# ─── TRANSACTIONS ─────────────────────────────────────────────────────
@app.route('/api/transactions', methods=['GET', 'OPTIONS'])
@require_auth
def get_transactions():
    with get_db() as db:
        rows = db.execute(
            "SELECT * FROM transactions WHERE user_id=? ORDER BY created DESC",
            (request.user_id,)
        ).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/transactions', methods=['POST'])
@require_auth
def add_transaction():
    d = request.get_json(force=True)
    with get_db() as db:
        cur = db.execute(
            "INSERT INTO transactions(user_id,type,amt,dsc,cat,dt,note) VALUES(?,?,?,?,?,?,?)",
            (request.user_id, d['type'], d['amt'], d['dsc'], d.get('cat','Other'), d.get('dt',''), d.get('note',''))
        )
        db.commit()
        tx_id = cur.lastrowid
        row = db.execute("SELECT * FROM transactions WHERE id=?", (tx_id,)).fetchone()
    return jsonify(dict(row)), 201

@app.route('/api/transactions/<int:tx_id>', methods=['PUT', 'OPTIONS'])
@require_auth
def update_transaction(tx_id):
    if request.method == 'OPTIONS': return make_response('', 200)
    d = request.get_json(force=True)
    with get_db() as db:
        db.execute(
            "UPDATE transactions SET dsc=?,amt=?,cat=?,note=? WHERE id=? AND user_id=?",
            (d['dsc'], d['amt'], d['cat'], d.get('note',''), tx_id, request.user_id)
        )
        db.commit()
        row = db.execute("SELECT * FROM transactions WHERE id=?", (tx_id,)).fetchone()
    return jsonify(dict(row))

@app.route('/api/transactions/<int:tx_id>', methods=['DELETE', 'OPTIONS'])
@require_auth
def delete_transaction(tx_id):
    if request.method == 'OPTIONS': return make_response('', 200)
    with get_db() as db:
        db.execute("DELETE FROM transactions WHERE id=? AND user_id=?", (tx_id, request.user_id))
        db.commit()
    return jsonify({'ok': True})

# ─── PERIODS (تاريخ الفترات) ──────────────────────────────────────────
@app.route('/api/periods', methods=['GET', 'OPTIONS'])
@require_auth
def get_periods():
    with get_db() as db:
        periods = db.execute(
            "SELECT * FROM periods WHERE user_id=? ORDER BY closed_at DESC",
            (request.user_id,)
        ).fetchall()
        result = []
        for p in periods:
            txs = db.execute(
                "SELECT * FROM period_transactions WHERE period_id=?", (p['id'],)
            ).fetchall()
            result.append({**dict(p), 'txs': [dict(t) for t in txs]})
    return jsonify(result)

@app.route('/api/periods', methods=['POST'])
@require_auth
def close_period():
    d = request.get_json(force=True)
    label = d.get('label', 'فترة')
    with get_db() as db:
        txs = db.execute(
            "SELECT * FROM transactions WHERE user_id=?", (request.user_id,)
        ).fetchall()
        txs_list = [dict(t) for t in txs]
        inc = sum(t['amt'] for t in txs_list if t['type']=='income')
        exp = sum(t['amt'] for t in txs_list if t['type']=='expense')
        cur = db.execute(
            "INSERT INTO periods(user_id,label,inc,exp,bal) VALUES(?,?,?,?,?)",
            (request.user_id, label, inc, exp, inc-exp)
        )
        period_id = cur.lastrowid
        for tx in txs_list:
            db.execute(
                "INSERT INTO period_transactions(period_id,type,amt,dsc,cat,dt,note) VALUES(?,?,?,?,?,?,?)",
                (period_id, tx['type'], tx['amt'], tx['dsc'], tx.get('cat',''), tx.get('dt',''), tx.get('note',''))
            )
        db.execute("DELETE FROM transactions WHERE user_id=?", (request.user_id,))
        db.commit()
    return jsonify({'ok': True, 'period_id': period_id}), 201

# ─── GOALS ────────────────────────────────────────────────────────────
@app.route('/api/goal', methods=['GET', 'OPTIONS'])
@require_auth
def get_goal():
    with get_db() as db:
        row = db.execute("SELECT amount FROM goals WHERE user_id=?", (request.user_id,)).fetchone()
    return jsonify({'amount': row['amount'] if row else 0})

@app.route('/api/goal', methods=['POST'])
@require_auth
def set_goal():
    d = request.get_json(force=True)
    amt = float(d.get('amount', 0))
    with get_db() as db:
        db.execute("INSERT INTO goals(user_id,amount) VALUES(?,?) ON CONFLICT(user_id) DO UPDATE SET amount=excluded.amount", (request.user_id, amt))
        db.commit()
    return jsonify({'ok': True, 'amount': amt})

# ─── VAULT ────────────────────────────────────────────────────────────
@app.route('/api/vault/folders', methods=['GET', 'OPTIONS'])
@require_auth
def get_vault_folders():
    with get_db() as db:
        rows = db.execute("SELECT * FROM vault_folders WHERE user_id=?", (request.user_id,)).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/vault/folders', methods=['POST'])
@require_auth
def add_vault_folder():
    d = request.get_json(force=True)
    fid = 'f' + secrets.token_hex(6)
    with get_db() as db:
        db.execute("INSERT INTO vault_folders(id,user_id,name,icon) VALUES(?,?,?,?)", (fid, request.user_id, d['name'], d.get('icon','📁')))
        db.commit()
    return jsonify({'id': fid, 'name': d['name'], 'icon': d.get('icon','📁')}), 201

@app.route('/api/vault/folders/<folder_id>', methods=['DELETE', 'OPTIONS'])
@require_auth
def delete_vault_folder(folder_id):
    if request.method == 'OPTIONS': return make_response('', 200)
    with get_db() as db:
        db.execute("DELETE FROM vault_folders WHERE id=? AND user_id=?", (folder_id, request.user_id))
        db.execute("UPDATE vault_accounts SET folder_id='' WHERE folder_id=? AND user_id=?", (folder_id, request.user_id))
        db.commit()
    return jsonify({'ok': True})

@app.route('/api/vault/accounts', methods=['GET', 'OPTIONS'])
@require_auth
def get_vault_accounts():
    with get_db() as db:
        rows = db.execute("SELECT * FROM vault_accounts WHERE user_id=? ORDER BY created_at DESC", (request.user_id,)).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/vault/accounts', methods=['POST'])
@require_auth
def add_vault_account():
    d = request.get_json(force=True)
    aid = 'va' + secrets.token_hex(8)
    with get_db() as db:
        db.execute(
            "INSERT INTO vault_accounts(id,user_id,folder_id,title,username,password,email,pin,url,note,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
            (aid, request.user_id, d.get('folderId',''), d['title'], d.get('username',''), d.get('password',''), d.get('email',''), d.get('pin',''), d.get('url',''), d.get('note',''), int(datetime.now().timestamp()*1000))
        )
        db.commit()
        row = db.execute("SELECT * FROM vault_accounts WHERE id=?", (aid,)).fetchone()
    return jsonify(dict(row)), 201

@app.route('/api/vault/accounts/<acc_id>', methods=['PUT', 'OPTIONS'])
@require_auth
def update_vault_account(acc_id):
    if request.method == 'OPTIONS': return make_response('', 200)
    d = request.get_json(force=True)
    with get_db() as db:
        db.execute(
            "UPDATE vault_accounts SET folder_id=?,title=?,username=?,password=?,email=?,pin=?,url=?,note=? WHERE id=? AND user_id=?",
            (d.get('folderId',''), d['title'], d.get('username',''), d.get('password',''), d.get('email',''), d.get('pin',''), d.get('url',''), d.get('note',''), acc_id, request.user_id)
        )
        db.commit()
        row = db.execute("SELECT * FROM vault_accounts WHERE id=?", (acc_id,)).fetchone()
    return jsonify(dict(row))

@app.route('/api/vault/accounts/<acc_id>', methods=['DELETE', 'OPTIONS'])
@require_auth
def delete_vault_account(acc_id):
    if request.method == 'OPTIONS': return make_response('', 200)
    with get_db() as db:
        db.execute("DELETE FROM vault_accounts WHERE id=? AND user_id=?", (acc_id, request.user_id))
        db.commit()
    return jsonify({'ok': True})

# ─── EXCEL EXPORT ─────────────────────────────────────────────────────
@app.route('/api/export-excel', methods=['POST', 'OPTIONS'])
@require_auth
def export_excel():
    if request.method == 'OPTIONS': return make_response('', 200)
    if not EXCEL_AVAILABLE:
        return jsonify({'error': 'مكتبة openpyxl غير مثبتة'}), 500
    d    = request.get_json(force=True)
    txs  = d.get('txs', [])
    label = d.get('label', 'تقرير مالي')
    wb_bytes = build_workbook(txs, label)
    buf = io.BytesIO(wb_bytes); buf.seek(0)
    filename = label.replace(' ', '_').replace('/', '-') + '.xlsx'
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ─── BACKUP ───────────────────────────────────────────────────────────
@app.route('/api/backup', methods=['GET', 'OPTIONS'])
@require_auth
def backup_db():
    if request.method == 'OPTIONS': return make_response('', 200)
    if request.role != 'admin':
        return jsonify({'error': 'ممنوع — admins only'}), 403
    if not os.path.exists(DB_PATH):
        return jsonify({'error': 'قاعدة البيانات غير موجودة'}), 404
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(
        DB_PATH,
        as_attachment=True,
        download_name=f'hizb_backup_{timestamp}.db',
        mimetype='application/octet-stream'
    )

# ─── CHANGE USERNAME ──────────────────────────────────────────────────
@app.route('/api/change-username', methods=['POST', 'OPTIONS'])
@require_auth
def change_username():
    if request.method == 'OPTIONS': return make_response('', 200)
    d = request.get_json(force=True)
    new_username = d.get('new_username', '').strip()
    password     = d.get('password', '')
    if not new_username or not password:
        return jsonify({'error': 'أدخل اسم المستخدم الجديد وكلمة المرور'}), 400
    if len(new_username) < 3:
        return jsonify({'error': 'اسم المستخدم يجب أن يكون 3 أحرف على الأقل'}), 400
    with get_db() as db:
        user = db.execute("SELECT * FROM users WHERE id=?", (request.user_id,)).fetchone()
        if not user or user['password'] != hash_pass(password):
            return jsonify({'error': 'كلمة المرور غير صحيحة'}), 401
        existing = db.execute("SELECT id FROM users WHERE username=? AND id!=?", (new_username, request.user_id)).fetchone()
        if existing:
            return jsonify({'error': 'اسم المستخدم موجود بالفعل'}), 409
        db.execute("UPDATE users SET username=? WHERE id=?", (new_username, request.user_id))
        db.commit()
    return jsonify({'ok': True, 'username': new_username, 'message': 'تم تغيير اسم المستخدم بنجاح ✓'})

# ─── CHANGE PASSWORD ──────────────────────────────────────────────────
@app.route('/api/change-password', methods=['POST', 'OPTIONS'])
@require_auth
def change_password():
    if request.method == 'OPTIONS': return make_response('', 200)
    d = request.get_json(force=True)
    old_pass = d.get('old_password', '')
    new_pass = d.get('new_password', '')
    if not old_pass or not new_pass:
        return jsonify({'error': 'أدخل كلمة المرور القديمة والجديدة'}), 400
    if len(new_pass) < 6:
        return jsonify({'error': 'كلمة المرور الجديدة يجب أن تكون 6 أحرف على الأقل'}), 400
    with get_db() as db:
        user = db.execute("SELECT * FROM users WHERE id=?", (request.user_id,)).fetchone()
        if not user or user['password'] != hash_pass(old_pass):
            return jsonify({'error': 'كلمة المرور القديمة غير صحيحة'}), 401
        db.execute("UPDATE users SET password=? WHERE id=?", (hash_pass(new_pass), request.user_id))
        # حذف كل الجلسات الأخرى عشان الأمان
        db.execute("DELETE FROM sessions WHERE user_id=? AND token!=?",
                   (request.user_id, request.headers.get('X-Auth-Token','')))
        db.commit()
    return jsonify({'ok': True, 'message': 'تم تغيير كلمة المرور بنجاح ✓'})

# ─── ADMIN: إدارة المستخدمين ──────────────────────────────────────────
@app.route('/api/admin/users', methods=['GET', 'OPTIONS'])
@require_auth
def admin_list_users():
    if request.role != 'admin':
        return jsonify({'error': 'ممنوع'}), 403
    with get_db() as db:
        rows = db.execute("SELECT id, username, role, created FROM users").fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/admin/users', methods=['POST'])
@require_auth
def admin_add_user():
    if request.role != 'admin':
        return jsonify({'error': 'ممنوع'}), 403
    d = request.get_json(force=True)
    uname = d.get('username','').strip()
    upass = d.get('password','')
    role  = d.get('role','user')
    if not uname or not upass:
        return jsonify({'error': 'اسم المستخدم وكلمة المرور مطلوبان'}), 400
    try:
        with get_db() as db:
            db.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)", (uname, hash_pass(upass), role))
            db.commit()
        return jsonify({'ok': True, 'username': uname}), 201
    except sqlite3.IntegrityError:
        return jsonify({'error': 'اسم المستخدم موجود بالفعل'}), 409

@app.route('/api/admin/users/<int:uid>', methods=['DELETE', 'OPTIONS'])
@require_auth
def admin_delete_user(uid):
    if request.method == 'OPTIONS': return make_response('', 200)
    if request.role != 'admin':
        return jsonify({'error': 'ممنوع'}), 403
    if uid == request.user_id:
        return jsonify({'error': 'لا يمكنك حذف نفسك'}), 400
    with get_db() as db:
        db.execute("DELETE FROM users WHERE id=?", (uid,))
        db.execute("DELETE FROM sessions WHERE user_id=?", (uid,))
        db.commit()
    return jsonify({'ok': True})

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('DEBUG', 'true').lower() == 'true'
    print(f'\n★  الحزب الاشتراكى Finance Server v2.0')
    print(f'★  قاعدة البيانات: {DB_PATH}')
    print(f'★  افتح المتصفح على: http://localhost:{port}\n')
    app.run(debug=debug, port=port, host='0.0.0.0')
