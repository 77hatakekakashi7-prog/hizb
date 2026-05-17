"""
server.py — الحزب الاشتراكى Finance Server v2.1
نظام المجموعات: يوزرات في نفس المجموعة يشاركوا نفس البيانات المالية
"""
from flask import Flask, request, send_file, jsonify, make_response, send_from_directory
import io, os, sqlite3, hashlib, secrets
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

        -- ─── المجموعات ────────────────────────────────────────────
        -- كل يوزر إما في مجموعة أو يعمل solo (group_id = NULL)
        CREATE TABLE IF NOT EXISTS groups (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            name      TEXT NOT NULL,
            created   TEXT DEFAULT (datetime('now'))
        );

        -- ربط اليوزرات بالمجموعات (يوزر واحد في مجموعة واحدة بس)
        CREATE TABLE IF NOT EXISTS group_members (
            user_id   INTEGER PRIMARY KEY,
            group_id  INTEGER NOT NULL,
            FOREIGN KEY(user_id)  REFERENCES users(id),
            FOREIGN KEY(group_id) REFERENCES groups(id)
        );

        -- ─── العمليات: owner_id = من أضافها، scope = group_id أو user_id ──
        CREATE TABLE IF NOT EXISTS transactions (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id   INTEGER NOT NULL,
            group_id  INTEGER,
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
            group_id  INTEGER,
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
            scope_key  TEXT PRIMARY KEY,
            amount     REAL DEFAULT 0
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

        # ── Migration: إنشاء جداول المجموعات لو مش موجودة ──
        db.executescript("""
        CREATE TABLE IF NOT EXISTS groups (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            name    TEXT NOT NULL,
            created TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS group_members (
            user_id  INTEGER PRIMARY KEY,
            group_id INTEGER NOT NULL,
            FOREIGN KEY(user_id)  REFERENCES users(id),
            FOREIGN KEY(group_id) REFERENCES groups(id)
        );
        """)

        # ── Migrations: إضافة عمود group_id لو مش موجود ──
        for tbl, col, coldef in [
            ('transactions',  'group_id', 'INTEGER'),
            ('periods',       'group_id', 'INTEGER'),
            ('vault_folders', 'group_id', 'INTEGER'),
            ('vault_accounts','group_id', 'INTEGER'),
        ]:
            cols = [r['name'] for r in db.execute(f"PRAGMA table_info({tbl})").fetchall()]
            if col not in cols:
                db.execute(f"ALTER TABLE {tbl} ADD COLUMN {col} {coldef}")

        # ── migrate goals table if old schema ──
        goal_cols = [r['name'] for r in db.execute("PRAGMA table_info(goals)").fetchall()]
        if 'user_id' in goal_cols and 'scope_key' not in goal_cols:
            db.execute("DROP TABLE goals")
            db.execute("CREATE TABLE goals (scope_key TEXT PRIMARY KEY, amount REAL DEFAULT 0)")

        # ── إنشاء مستخدمين افتراضيين ──
        for uname, upass, role in [('admin','hizb2024','admin'), ('محاسب','hizb2024','user')]:
            if not db.execute("SELECT id FROM users WHERE username=?", (uname,)).fetchone():
                db.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                           (uname, hash_pass(upass), role))
        db.commit()
    print(f"✓ قاعدة البيانات جاهزة: {DB_PATH}")

def hash_pass(p):
    return hashlib.sha256(p.encode()).hexdigest()

# ─── helper: جيب الـ scope (group_id أو None) للـ user ──────────────
def get_scope(db, user_id):
    """Returns (group_id_or_None, scope_key_str)"""
    row = db.execute("SELECT group_id FROM group_members WHERE user_id=?", (user_id,)).fetchone()
    if row:
        gid = row['group_id']
        return gid, f"g{gid}"
    return None, f"u{user_id}"

def tx_filter(db, user_id):
    """Returns SQL WHERE clause and params to filter transactions by scope."""
    gid, _ = get_scope(db, user_id)
    if gid:
        return "group_id = ?", (gid,)
    return "user_id = ? AND (group_id IS NULL OR group_id = 0)", (user_id,)

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

# ─── Auth decorator ──────────────────────────────────────────────────
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
                "SELECT s.user_id, u.username, u.role FROM sessions s "
                "JOIN users u ON u.id=s.user_id "
                "WHERE s.token=? AND s.expires > datetime('now')",
                (token,)
            ).fetchone()
        if not row:
            return jsonify({'error': 'انتهت الجلسة'}), 401
        request.user_id  = row['user_id']
        request.username = row['username']
        request.role     = row['role']
        return f(*args, **kwargs)
    return decorated

# ═══════════════════════════════════════════════════════════════════════
# AUTH
# ═══════════════════════════════════════════════════════════════════════
@app.route('/api/login', methods=['POST', 'OPTIONS'])
def login():
    if request.method == 'OPTIONS': return make_response('', 200)
    data  = request.get_json(force=True)
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
        db.execute("INSERT INTO sessions(token,user_id,expires) VALUES(?,?,?)",
                   (token, user['id'], expires))
        db.commit()
        # جيب المجموعة لو موجودة
        gm = db.execute("SELECT g.id, g.name FROM group_members gm JOIN groups g ON g.id=gm.group_id WHERE gm.user_id=?", (user['id'],)).fetchone()
    group_info = {'id': gm['id'], 'name': gm['name']} if gm else None
    resp = jsonify({'token': token, 'username': uname, 'role': user['role'], 'group': group_info})
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

# ═══════════════════════════════════════════════════════════════════════
# TRANSACTIONS  — scope-aware
# ═══════════════════════════════════════════════════════════════════════
@app.route('/api/transactions', methods=['GET', 'OPTIONS'])
@require_auth
def get_transactions():
    with get_db() as db:
        where, params = tx_filter(db, request.user_id)
        rows = db.execute(
            f"SELECT t.*, u.username as added_by FROM transactions t "
            f"LEFT JOIN users u ON u.id=t.user_id "
            f"WHERE {where} ORDER BY t.created DESC",
            params
        ).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/transactions', methods=['POST'])
@require_auth
def add_transaction():
    d = request.get_json(force=True)
    with get_db() as db:
        gid, _ = get_scope(db, request.user_id)
        cur = db.execute(
            "INSERT INTO transactions(user_id,group_id,type,amt,dsc,cat,dt,note) VALUES(?,?,?,?,?,?,?,?)",
            (request.user_id, gid, d['type'], d['amt'], d['dsc'],
             d.get('cat','Other'), d.get('dt',''), d.get('note',''))
        )
        db.commit()
        row = db.execute(
            "SELECT t.*, u.username as added_by FROM transactions t "
            "LEFT JOIN users u ON u.id=t.user_id WHERE t.id=?", (cur.lastrowid,)
        ).fetchone()
    return jsonify(dict(row)), 201

@app.route('/api/transactions/<int:tx_id>', methods=['PUT', 'OPTIONS'])
@require_auth
def update_transaction(tx_id):
    if request.method == 'OPTIONS': return make_response('', 200)
    d = request.get_json(force=True)
    with get_db() as db:
        where, params = tx_filter(db, request.user_id)
        # تأكد إن العملية في نطاق اليوزر/المجموعة
        existing = db.execute(f"SELECT id FROM transactions WHERE id=? AND {where}", (tx_id, *params)).fetchone()
        if not existing:
            return jsonify({'error': 'غير مصرح'}), 403
        db.execute(
            "UPDATE transactions SET dsc=?,amt=?,cat=?,note=? WHERE id=?",
            (d['dsc'], d['amt'], d['cat'], d.get('note',''), tx_id)
        )
        db.commit()
        row = db.execute(
            "SELECT t.*, u.username as added_by FROM transactions t "
            "LEFT JOIN users u ON u.id=t.user_id WHERE t.id=?", (tx_id,)
        ).fetchone()
    return jsonify(dict(row))

@app.route('/api/transactions/<int:tx_id>', methods=['DELETE', 'OPTIONS'])
@require_auth
def delete_transaction(tx_id):
    if request.method == 'OPTIONS': return make_response('', 200)
    with get_db() as db:
        where, params = tx_filter(db, request.user_id)
        existing = db.execute(f"SELECT id FROM transactions WHERE id=? AND {where}", (tx_id, *params)).fetchone()
        if not existing:
            return jsonify({'error': 'غير مصرح'}), 403
        db.execute("DELETE FROM transactions WHERE id=?", (tx_id,))
        db.commit()
    return jsonify({'ok': True})

# ═══════════════════════════════════════════════════════════════════════
# PERIODS
# ═══════════════════════════════════════════════════════════════════════
@app.route('/api/periods', methods=['GET', 'OPTIONS'])
@require_auth
def get_periods():
    with get_db() as db:
        gid, _ = get_scope(db, request.user_id)
        if gid:
            periods = db.execute(
                "SELECT * FROM periods WHERE group_id=? ORDER BY closed_at DESC", (gid,)
            ).fetchall()
        else:
            periods = db.execute(
                "SELECT * FROM periods WHERE user_id=? AND (group_id IS NULL OR group_id=0) ORDER BY closed_at DESC",
                (request.user_id,)
            ).fetchall()
        result = []
        for p in periods:
            txs = db.execute("SELECT * FROM period_transactions WHERE period_id=?", (p['id'],)).fetchall()
            result.append({**dict(p), 'txs': [dict(t) for t in txs]})
    return jsonify(result)

@app.route('/api/periods', methods=['POST'])
@require_auth
def close_period():
    d     = request.get_json(force=True)
    label = d.get('label', 'فترة')
    with get_db() as db:
        gid, _ = get_scope(db, request.user_id)
        where, params = tx_filter(db, request.user_id)
        txs_list = [dict(t) for t in db.execute(
            f"SELECT * FROM transactions WHERE {where}", params
        ).fetchall()]
        inc = sum(t['amt'] for t in txs_list if t['type']=='income')
        exp = sum(t['amt'] for t in txs_list if t['type']=='expense')
        cur = db.execute(
            "INSERT INTO periods(user_id,group_id,label,inc,exp,bal) VALUES(?,?,?,?,?,?)",
            (request.user_id, gid, label, inc, exp, inc-exp)
        )
        pid = cur.lastrowid
        for tx in txs_list:
            db.execute(
                "INSERT INTO period_transactions(period_id,type,amt,dsc,cat,dt,note) VALUES(?,?,?,?,?,?,?)",
                (pid, tx['type'], tx['amt'], tx['dsc'], tx.get('cat',''), tx.get('dt',''), tx.get('note',''))
            )
        db.execute(f"DELETE FROM transactions WHERE {where}", params)
        db.commit()
    return jsonify({'ok': True, 'period_id': pid}), 201

# ═══════════════════════════════════════════════════════════════════════
# GOALS
# ═══════════════════════════════════════════════════════════════════════
@app.route('/api/goal', methods=['GET', 'OPTIONS'])
@require_auth
def get_goal():
    with get_db() as db:
        _, scope_key = get_scope(db, request.user_id)
        row = db.execute("SELECT amount FROM goals WHERE scope_key=?", (scope_key,)).fetchone()
    return jsonify({'amount': row['amount'] if row else 0})

@app.route('/api/goal', methods=['POST'])
@require_auth
def set_goal():
    d   = request.get_json(force=True)
    amt = float(d.get('amount', 0))
    with get_db() as db:
        _, scope_key = get_scope(db, request.user_id)
        db.execute(
            "INSERT INTO goals(scope_key,amount) VALUES(?,?) ON CONFLICT(scope_key) DO UPDATE SET amount=excluded.amount",
            (scope_key, amt)
        )
        db.commit()
    return jsonify({'ok': True, 'amount': amt})

# ═══════════════════════════════════════════════════════════════════════
# VAULT — مشترك في المجموعة
# ═══════════════════════════════════════════════════════════════════════
def vault_filter(db, user_id):
    gid, _ = get_scope(db, user_id)
    if gid:
        return "group_id = ?", (gid,)
    return "user_id = ? AND (group_id IS NULL OR group_id = 0)", (user_id,)

@app.route('/api/vault/folders', methods=['GET', 'OPTIONS'])
@require_auth
def get_vault_folders():
    with get_db() as db:
        where, params = vault_filter(db, request.user_id)
        rows = db.execute(f"SELECT * FROM vault_folders WHERE {where}", params).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/vault/folders', methods=['POST'])
@require_auth
def add_vault_folder():
    d   = request.get_json(force=True)
    fid = 'f' + secrets.token_hex(6)
    with get_db() as db:
        gid, _ = get_scope(db, request.user_id)
        db.execute("INSERT INTO vault_folders(id,user_id,group_id,name,icon) VALUES(?,?,?,?,?)",
                   (fid, request.user_id, gid, d['name'], d.get('icon','📁')))
        db.commit()
    return jsonify({'id': fid, 'name': d['name'], 'icon': d.get('icon','📁')}), 201

@app.route('/api/vault/folders/<folder_id>', methods=['DELETE', 'OPTIONS'])
@require_auth
def delete_vault_folder(folder_id):
    if request.method == 'OPTIONS': return make_response('', 200)
    with get_db() as db:
        where, params = vault_filter(db, request.user_id)
        existing = db.execute(f"SELECT id FROM vault_folders WHERE id=? AND {where}", (folder_id, *params)).fetchone()
        if not existing: return jsonify({'error': 'غير مصرح'}), 403
        db.execute("DELETE FROM vault_folders WHERE id=?", (folder_id,))
        db.execute("UPDATE vault_accounts SET folder_id='' WHERE folder_id=?", (folder_id,))
        db.commit()
    return jsonify({'ok': True})

@app.route('/api/vault/accounts', methods=['GET', 'OPTIONS'])
@require_auth
def get_vault_accounts():
    with get_db() as db:
        where, params = vault_filter(db, request.user_id)
        rows = db.execute(
            f"SELECT va.*, u.username as added_by FROM vault_accounts va "
            f"LEFT JOIN users u ON u.id=va.user_id "
            f"WHERE {where} ORDER BY va.created_at DESC",
            params
        ).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/vault/accounts', methods=['POST'])
@require_auth
def add_vault_account():
    d   = request.get_json(force=True)
    aid = 'va' + secrets.token_hex(8)
    with get_db() as db:
        gid, _ = get_scope(db, request.user_id)
        db.execute(
            "INSERT INTO vault_accounts(id,user_id,group_id,folder_id,title,username,password,email,pin,url,note,created_at) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
            (aid, request.user_id, gid, d.get('folderId',''), d['title'],
             d.get('username',''), d.get('password',''), d.get('email',''),
             d.get('pin',''), d.get('url',''), d.get('note',''),
             int(datetime.now().timestamp()*1000))
        )
        db.commit()
        row = db.execute(
            "SELECT va.*, u.username as added_by FROM vault_accounts va "
            "LEFT JOIN users u ON u.id=va.user_id WHERE va.id=?", (aid,)
        ).fetchone()
    return jsonify(dict(row)), 201

@app.route('/api/vault/accounts/<acc_id>', methods=['PUT', 'OPTIONS'])
@require_auth
def update_vault_account(acc_id):
    if request.method == 'OPTIONS': return make_response('', 200)
    d = request.get_json(force=True)
    with get_db() as db:
        where, params = vault_filter(db, request.user_id)
        existing = db.execute(f"SELECT id FROM vault_accounts WHERE id=? AND {where}", (acc_id, *params)).fetchone()
        if not existing: return jsonify({'error': 'غير مصرح'}), 403
        db.execute(
            "UPDATE vault_accounts SET folder_id=?,title=?,username=?,password=?,email=?,pin=?,url=?,note=? WHERE id=?",
            (d.get('folderId',''), d['title'], d.get('username',''), d.get('password',''),
             d.get('email',''), d.get('pin',''), d.get('url',''), d.get('note',''), acc_id)
        )
        db.commit()
        row = db.execute(
            "SELECT va.*, u.username as added_by FROM vault_accounts va "
            "LEFT JOIN users u ON u.id=va.user_id WHERE va.id=?", (acc_id,)
        ).fetchone()
    return jsonify(dict(row))

@app.route('/api/vault/accounts/<acc_id>', methods=['DELETE', 'OPTIONS'])
@require_auth
def delete_vault_account(acc_id):
    if request.method == 'OPTIONS': return make_response('', 200)
    with get_db() as db:
        where, params = vault_filter(db, request.user_id)
        existing = db.execute(f"SELECT id FROM vault_accounts WHERE id=? AND {where}", (acc_id, *params)).fetchone()
        if not existing: return jsonify({'error': 'غير مصرح'}), 403
        db.execute("DELETE FROM vault_accounts WHERE id=?", (acc_id,))
        db.commit()
    return jsonify({'ok': True})

# ═══════════════════════════════════════════════════════════════════════
# GROUPS — Admin only
# ═══════════════════════════════════════════════════════════════════════
@app.route('/api/admin/groups', methods=['GET', 'OPTIONS'])
@require_auth
def list_groups():
    if request.role != 'admin': return jsonify({'error': 'ممنوع'}), 403
    with get_db() as db:
        groups = db.execute("SELECT * FROM groups ORDER BY id").fetchall()
        result = []
        for g in groups:
            members = db.execute(
                "SELECT u.id, u.username, u.role FROM group_members gm "
                "JOIN users u ON u.id=gm.user_id WHERE gm.group_id=?", (g['id'],)
            ).fetchall()
            result.append({**dict(g), 'members': [dict(m) for m in members]})
    return jsonify(result)

@app.route('/api/admin/groups', methods=['POST'])
@require_auth
def create_group():
    if request.role != 'admin': return jsonify({'error': 'ممنوع'}), 403
    d    = request.get_json(force=True)
    name = d.get('name','').strip()
    if not name: return jsonify({'error': 'أدخل اسم المجموعة'}), 400
    with get_db() as db:
        cur = db.execute("INSERT INTO groups(name) VALUES(?)", (name,))
        db.commit()
        gid = cur.lastrowid
    return jsonify({'id': gid, 'name': name, 'members': []}), 201

@app.route('/api/admin/groups/<int:gid>', methods=['DELETE', 'OPTIONS'])
@require_auth
def delete_group(gid):
    if request.method == 'OPTIONS': return make_response('', 200)
    if request.role != 'admin': return jsonify({'error': 'ممنوع'}), 403
    with get_db() as db:
        db.execute("DELETE FROM group_members WHERE group_id=?", (gid,))
        db.execute("DELETE FROM groups WHERE id=?", (gid,))
        db.commit()
    return jsonify({'ok': True})

@app.route('/api/admin/groups/<int:gid>/members', methods=['POST'])
@require_auth
def add_group_member(gid):
    if request.role != 'admin': return jsonify({'error': 'ممنوع'}), 403
    d   = request.get_json(force=True)
    uid = d.get('user_id')
    if not uid: return jsonify({'error': 'أدخل user_id'}), 400
    with get_db() as db:
        group = db.execute("SELECT id FROM groups WHERE id=?", (gid,)).fetchone()
        if not group: return jsonify({'error': 'المجموعة غير موجودة'}), 404
        db.execute(
            "INSERT INTO group_members(user_id,group_id) VALUES(?,?) "
            "ON CONFLICT(user_id) DO UPDATE SET group_id=excluded.group_id",
            (uid, gid)
        )
        db.commit()
    return jsonify({'ok': True})

@app.route('/api/admin/groups/<int:gid>/members/<int:uid>', methods=['DELETE', 'OPTIONS'])
@require_auth
def remove_group_member(gid, uid):
    if request.method == 'OPTIONS': return make_response('', 200)
    if request.role != 'admin': return jsonify({'error': 'ممنوع'}), 403
    with get_db() as db:
        db.execute("DELETE FROM group_members WHERE group_id=? AND user_id=?", (gid, uid))
        db.commit()
    return jsonify({'ok': True})

# get all users (for assignment UI)
@app.route('/api/admin/users', methods=['GET', 'OPTIONS'])
@require_auth
def admin_list_users():
    if request.role != 'admin': return jsonify({'error': 'ممنوع'}), 403
    with get_db() as db:
        rows = db.execute("SELECT id, username, role, created FROM users").fetchall()
        # أضف المجموعة لكل يوزر
        result = []
        for u in rows:
            gm = db.execute(
                "SELECT g.id, g.name FROM group_members gm JOIN groups g ON g.id=gm.group_id WHERE gm.user_id=?",
                (u['id'],)
            ).fetchone()
            result.append({**dict(u), 'group': {'id': gm['id'], 'name': gm['name']} if gm else None})
    return jsonify(result)

@app.route('/api/admin/users', methods=['POST'])
@require_auth
def admin_add_user():
    if request.role != 'admin': return jsonify({'error': 'ممنوع'}), 403
    d     = request.get_json(force=True)
    uname = d.get('username','').strip()
    upass = d.get('password','')
    role  = d.get('role','user')
    if not uname or not upass:
        return jsonify({'error': 'اسم المستخدم وكلمة المرور مطلوبان'}), 400
    try:
        with get_db() as db:
            db.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                       (uname, hash_pass(upass), role))
            db.commit()
        return jsonify({'ok': True, 'username': uname}), 201
    except Exception:
        return jsonify({'error': 'اسم المستخدم موجود بالفعل'}), 409

@app.route('/api/admin/users/<int:uid>', methods=['DELETE', 'OPTIONS'])
@require_auth
def admin_delete_user(uid):
    if request.method == 'OPTIONS': return make_response('', 200)
    if request.role != 'admin': return jsonify({'error': 'ممنوع'}), 403
    if uid == request.user_id: return jsonify({'error': 'لا يمكنك حذف نفسك'}), 400
    with get_db() as db:
        db.execute("DELETE FROM users WHERE id=?", (uid,))
        db.execute("DELETE FROM sessions WHERE user_id=?", (uid,))
        db.execute("DELETE FROM group_members WHERE user_id=?", (uid,))
        db.commit()
    return jsonify({'ok': True})

# ─── تعيين يوزر لمجموعة مباشرة (shortcut) ───────────────────────────
@app.route('/api/admin/users/<int:uid>/group', methods=['PUT', 'OPTIONS'])
@require_auth
def set_user_group(uid):
    if request.method == 'OPTIONS': return make_response('', 200)
    if request.role != 'admin': return jsonify({'error': 'ممنوع'}), 403
    d   = request.get_json(force=True)
    gid = d.get('group_id')  # None = إزالة من المجموعة
    with get_db() as db:
        if gid is None:
            db.execute("DELETE FROM group_members WHERE user_id=?", (uid,))
        else:
            group = db.execute("SELECT id FROM groups WHERE id=?", (gid,)).fetchone()
            if not group: return jsonify({'error': 'المجموعة غير موجودة'}), 404
            db.execute(
                "INSERT INTO group_members(user_id,group_id) VALUES(?,?) "
                "ON CONFLICT(user_id) DO UPDATE SET group_id=excluded.group_id",
                (uid, gid)
            )
        db.commit()
    return jsonify({'ok': True})

# ═══════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ═══════════════════════════════════════════════════════════════════════
@app.route('/api/export-excel', methods=['POST', 'OPTIONS'])
@require_auth
def export_excel():
    if request.method == 'OPTIONS': return make_response('', 200)
    if not EXCEL_AVAILABLE:
        return jsonify({'error': 'مكتبة openpyxl غير مثبتة'}), 500
    d        = request.get_json(force=True)
    txs      = d.get('txs', [])
    label    = d.get('label', 'تقرير مالي')
    wb_bytes = build_workbook(txs, label)
    buf      = io.BytesIO(wb_bytes); buf.seek(0)
    filename = label.replace(' ', '_').replace('/', '-') + '.xlsx'
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ═══════════════════════════════════════════════════════════════════════
# BACKUP
# ═══════════════════════════════════════════════════════════════════════
@app.route('/api/backup', methods=['GET', 'OPTIONS'])
@require_auth
def backup_db():
    if request.method == 'OPTIONS': return make_response('', 200)
    if request.role != 'admin': return jsonify({'error': 'ممنوع'}), 403
    if not os.path.exists(DB_PATH): return jsonify({'error': 'قاعدة البيانات غير موجودة'}), 404
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(DB_PATH, as_attachment=True,
                     download_name=f'hizb_backup_{ts}.db',
                     mimetype='application/octet-stream')

# ═══════════════════════════════════════════════════════════════════════
# BOT INCOME ENDPOINT — يُستخدم من Discord Bot فقط
# ═══════════════════════════════════════════════════════════════════════
BOT_API_KEY = os.environ.get('BOT_API_KEY', 'changeme-secret-key')

@app.route('/api/bot/income', methods=['POST', 'OPTIONS'])
def bot_income():
    """
    يستقبل income من البوت بعد إغلاق التيكيت.
    Headers: X-Bot-Key: <BOT_API_KEY>
    Body JSON: { amt, dsc, cat, note, username }
        username = اسم المستخدم في الموقع اللي هيتسجل باسمه
    """
    if request.method == 'OPTIONS':
        return make_response('', 200)

    # ── التحقق من الـ API Key ──
    key = request.headers.get('X-Bot-Key', '')
    if key != BOT_API_KEY:
        return jsonify({'error': 'Unauthorized'}), 401

    d = request.get_json(force=True)
    amt      = float(d.get('amt', 0))
    dsc      = d.get('dsc', '').strip()
    cat      = d.get('cat', 'Discord Sales').strip()
    note     = d.get('note', '').strip()
    username = d.get('username', 'admin').strip()
    dt       = datetime.now().strftime('%Y-%m-%d')

    if amt <= 0 or not dsc:
        return jsonify({'error': 'amt و dsc مطلوبان'}), 400

    with get_db() as db:
        # نجيب الـ user_id من الـ username
        user = db.execute("SELECT id FROM users WHERE username=?", (username,)).fetchone()
        if not user:
            # fallback على admin
            user = db.execute("SELECT id FROM users WHERE role='admin' LIMIT 1").fetchone()
        if not user:
            return jsonify({'error': 'لا يوجد مستخدم'}), 500

        uid = user['id']
        gid, _ = get_scope(db, uid)

        db.execute(
            "INSERT INTO transactions(user_id, group_id, type, amt, dsc, cat, dt, note) "
            "VALUES(?, ?, 'income', ?, ?, ?, ?, ?)",
            (uid, gid, amt, dsc, cat, dt, note)
        )
        db.commit()

    return jsonify({'ok': True, 'recorded': {'amt': amt, 'dsc': dsc, 'cat': cat}}), 201

# ═══════════════════════════════════════════════════════════════════════
# CHANGE USERNAME / PASSWORD
# ═══════════════════════════════════════════════════════════════════════
@app.route('/api/change-username', methods=['POST', 'OPTIONS'])
@require_auth
def change_username():
    if request.method == 'OPTIONS': return make_response('', 200)
    d            = request.get_json(force=True)
    new_username = d.get('new_username','').strip()
    password     = d.get('password','')
    if not new_username or not password:
        return jsonify({'error': 'أدخل اسم المستخدم الجديد وكلمة المرور'}), 400
    if len(new_username) < 3:
        return jsonify({'error': 'اسم المستخدم يجب أن يكون 3 أحرف على الأقل'}), 400
    with get_db() as db:
        user = db.execute("SELECT * FROM users WHERE id=?", (request.user_id,)).fetchone()
        if not user or user['password'] != hash_pass(password):
            return jsonify({'error': 'كلمة المرور غير صحيحة'}), 401
        if db.execute("SELECT id FROM users WHERE username=? AND id!=?",
                      (new_username, request.user_id)).fetchone():
            return jsonify({'error': 'اسم المستخدم موجود بالفعل'}), 409
        db.execute("UPDATE users SET username=? WHERE id=?", (new_username, request.user_id))
        db.commit()
    return jsonify({'ok': True, 'username': new_username})

@app.route('/api/change-password', methods=['POST', 'OPTIONS'])
@require_auth
def change_password():
    if request.method == 'OPTIONS': return make_response('', 200)
    d        = request.get_json(force=True)
    old_pass = d.get('old_password','')
    new_pass = d.get('new_password','')
    if not old_pass or not new_pass:
        return jsonify({'error': 'أدخل كلمة المرور القديمة والجديدة'}), 400
    if len(new_pass) < 6:
        return jsonify({'error': 'كلمة المرور يجب أن تكون 6 أحرف على الأقل'}), 400
    with get_db() as db:
        user = db.execute("SELECT * FROM users WHERE id=?", (request.user_id,)).fetchone()
        if not user or user['password'] != hash_pass(old_pass):
            return jsonify({'error': 'كلمة المرور القديمة غير صحيحة'}), 401
        db.execute("UPDATE users SET password=? WHERE id=?", (hash_pass(new_pass), request.user_id))
        db.execute("DELETE FROM sessions WHERE user_id=? AND token!=?",
                   (request.user_id, request.headers.get('X-Auth-Token','')))
        db.commit()
    return jsonify({'ok': True})

# ═══════════════════════════════════════════════════════════════════════
# استدعاء init_db عند تحميل الـ module (يشتغل مع gunicorn وبدونه)
init_db()

if __name__ == '__main__':
    port  = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('DEBUG', 'true').lower() == 'true'
    print(f'\n★  الحزب الاشتراكى Finance Server v2.1')
    print(f'★  قاعدة البيانات: {DB_PATH}')
    print(f'★  افتح المتصفح على: http://localhost:{port}\n')
    app.run(debug=debug, port=port, host='0.0.0.0')
