from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file, make_response
import sqlite3, os, json, hashlib, csv, io, base64
from datetime import datetime, date, timedelta
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash

try:
    import qrcode
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

app = Flask(__name__)
app.secret_key = 'chatorichaat_secret_2025_v2'
DB_PATH = os.path.join(os.path.dirname(__file__), 'instance', 'chatori.db')

# ─── DB ────────────────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = get_db()
    c = conn.cursor()
    c.executescript('''
        CREATE TABLE IF NOT EXISTS admin (
            id INTEGER PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            mobile TEXT UNIQUE,
            email TEXT,
            password TEXT,
            is_registered INTEGER DEFAULT 0,
            total_visits INTEGER DEFAULT 0,
            total_spent REAL DEFAULT 0,
            first_visit TEXT,
            last_visit TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_number TEXT UNIQUE NOT NULL,
            customer_id INTEGER,
            customer_name TEXT,
            mobile TEXT,
            subtotal REAL DEFAULT 0,
            discount REAL DEFAULT 0,
            gst REAL DEFAULT 0,
            grand_total REAL DEFAULT 0,
            payment_method TEXT DEFAULT 'UPI',
            payment_status TEXT DEFAULT 'Pending',
            order_status TEXT DEFAULT 'Placed',
            upi_ref TEXT,
            bill_id INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(customer_id) REFERENCES customers(id),
            FOREIGN KEY(bill_id) REFERENCES bills(id)
        );
        CREATE TABLE IF NOT EXISTS order_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER NOT NULL,
            item_name TEXT,
            price REAL,
            quantity INTEGER,
            total REAL,
            FOREIGN KEY(order_id) REFERENCES orders(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_name TEXT UNIQUE NOT NULL,
            unit TEXT DEFAULT 'kg',
            available_qty REAL DEFAULT 0,
            used_qty REAL DEFAULT 0,
            purchase_cost REAL DEFAULT 0,
            selling_price REAL DEFAULT 0,
            low_stock_threshold REAL DEFAULT 5,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        );
        CREATE TABLE IF NOT EXISTS menu_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            price REAL NOT NULL,
            category TEXT DEFAULT 'Special',
            description TEXT,
            available INTEGER DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS bills (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bill_number TEXT UNIQUE NOT NULL,
            customer_id INTEGER,
            customer_name TEXT,
            mobile TEXT,
            date TEXT NOT NULL,
            subtotal REAL DEFAULT 0,
            discount REAL DEFAULT 0,
            gst REAL DEFAULT 0,
            grand_total REAL DEFAULT 0,
            payment_method TEXT DEFAULT 'Cash',
            payment_status TEXT DEFAULT 'Paid',
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(customer_id) REFERENCES customers(id)
        );
        CREATE TABLE IF NOT EXISTS bill_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bill_id INTEGER NOT NULL,
            item_name TEXT,
            price REAL,
            quantity INTEGER,
            total REAL,
            FOREIGN KEY(bill_id) REFERENCES bills(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL CHECK(type IN ('credit','debit')),
            category TEXT NOT NULL,
            description TEXT,
            amount REAL NOT NULL,
            date TEXT NOT NULL,
            reference_id TEXT,
            payment_method TEXT DEFAULT 'Cash',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            category TEXT NOT NULL,
            description TEXT,
            amount REAL NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS reviews (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            rating INTEGER NOT NULL CHECK(rating BETWEEN 1 AND 5),
            comment TEXT,
            source TEXT DEFAULT 'website',
            approved INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            message TEXT NOT NULL,
            type TEXT DEFAULT 'info',
            read_status INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS team_members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            role TEXT,
            description TEXT,
            photo TEXT,
            display_order INTEGER DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS journey_photos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            phase TEXT DEFAULT 'inauguration',
            title TEXT,
            description TEXT,
            date TEXT,
            filename TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE INDEX IF NOT EXISTS idx_bills_date ON bills(date);
        CREATE INDEX IF NOT EXISTS idx_bills_mobile ON bills(mobile);
        CREATE INDEX IF NOT EXISTS idx_transactions_date ON transactions(date);
        CREATE INDEX IF NOT EXISTS idx_expenses_date ON expenses(date);
        CREATE INDEX IF NOT EXISTS idx_orders_mobile ON orders(mobile);
        CREATE INDEX IF NOT EXISTS idx_orders_customer ON orders(customer_id);
        CREATE INDEX IF NOT EXISTS idx_customers_mobile ON customers(mobile);
    ''')

    # Migration: add columns if upgrading an older DB created before this version
    existing_cols = [row[1] for row in c.execute("PRAGMA table_info(customers)").fetchall()]
    for col, coltype in [('email','TEXT'), ('password','TEXT'), ('is_registered','INTEGER DEFAULT 0')]:
        if col not in existing_cols:
            c.execute(f"ALTER TABLE customers ADD COLUMN {col} {coltype}")

    # Default settings (UPI payment + Google Reviews integration, editable from Admin > Settings)
    default_settings = [
        ('upi_id', 'chatorichaat@upi'),
        ('upi_name', 'Chatori Chaat'),
        ('google_place_id', ''),
        ('google_api_key', ''),
        ('low_stock_default_threshold', '5'),
    ]
    for k, v in default_settings:
        c.execute("INSERT OR IGNORE INTO settings (key,value) VALUES (?,?)", (k, v))

    # Seed inventory (admin can edit quantities/costs from Admin > Inventory)
    inventory_seed = [
        ('Papdi', 'kg', 10, 2, 120, 0, 3),
        ('Puri (Pani Puri Balls)', 'packets', 20, 4, 80, 0, 5),
        ('Sev', 'kg', 8, 1.5, 160, 0, 2),
        ('Dahi (Curd)', 'litre', 15, 3, 70, 0, 4),
        ('Ragda', 'kg', 12, 2, 60, 0, 3),
        ('Cheese', 'kg', 5, 0.5, 380, 0, 1.5),
        ('Tamarind Chutney', 'litre', 6, 1, 150, 0, 2),
        ('Mint Chutney', 'litre', 6, 1, 90, 0, 2),
        ('Onion', 'kg', 10, 2, 40, 0, 3),
        ('Sprouts (Ragda Mix)', 'kg', 6, 1, 90, 0, 2),
    ]
    for name, unit, avail, used, cost, sell, threshold in inventory_seed:
        c.execute("""INSERT OR IGNORE INTO inventory
            (product_name, unit, available_qty, used_qty, purchase_cost, selling_price, low_stock_threshold)
            VALUES (?,?,?,?,?,?,?)""", (name, unit, avail, used, cost, sell, threshold))

    # Seed admin
    c.execute("INSERT OR IGNORE INTO admin (username, password) VALUES (?,?)",
              ('admin', hashlib.sha256('admin123'.encode()).hexdigest()))

    # Seed menu
    menu = [
        ('Panipuri',25,'Special'),('Masala Puri',25,'Special'),
        ('Shevpuri',45,'Special'),('Ragda Puri',40,'Special'),
        ('Dahipuri',50,'Special'),('Papdi Chaat',50,'Special'),
        ('Ragda Chaat',45,'Special'),('Dahi Khasta',40,'Special'),
        ('Bhel',45,'Special'),('Khasta Bhel',60,'Main Dishes'),
        ('Cheese Shev Puri',60,'Main Dishes'),('Cheese Ragda Chaat',60,'Main Dishes'),
    ]
    for m in menu:
        c.execute("INSERT OR IGNORE INTO menu_items (name,price,category) VALUES (?,?,?)", m)

    # Seed team
    c.execute("DELETE FROM team_members")
    team = [
        ('Piyush Manoj Patil','Co-Founder & Head Chef','The mastermind behind our secret spice blends and authentic chaat recipes.',1),
        ('Parnika Manoj Patil','Co-Founder & Operations Head','Ensures every customer leaves with a smile. Manages quality, hygiene, and warm hospitality.',2),
        ('Vaibhav Jadkar','Marketing & Social Media Handler','Responsible for brand promotion, customer engagement, and social media management.',3),
        ('Manoj Patil','Manager','Responsible for daily operations, team coordination, and customer satisfaction.',4),
        ('Druvita Patil','Support Team Member','Assists customers and supports daily business activities.',5),
        ('Rupali Patil','Support Team Member','Provides customer support and operational assistance.',6),
    ]
    for t in team:
        c.execute("INSERT OR IGNORE INTO team_members (name,role,description,display_order) VALUES (?,?,?,?)", t)

    # Seed reviews
    reviews = [
        ('Rahul Sharma',5,'Best chaat in Nashik! Panipuri is absolutely amazing. Must visit!'),
        ('Priya Patel',5,'Dahipuri yaar... mast! Hygiene is top notch. Love this place!'),
        ('Amit Deshmukh',4,'Ragda puri was delicious. Great taste, great price. Will come again!'),
        ('Sneha Joshi',5,'Cheese Ragda Chaat is a must try! Owner couple is very sweet.'),
    ]
    for r in reviews:
        c.execute("INSERT OR IGNORE INTO reviews (name,rating,comment) VALUES (?,?,?)", r)

    conn.commit()
    conn.close()

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('admin'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def customer_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('customer_id'):
            return redirect(url_for('customer_login', next=request.path))
        return f(*args, **kwargs)
    return decorated

def gen_bill_number():
    today = date.today().strftime('%Y%m%d')
    conn = get_db()
    count = conn.execute("SELECT COUNT(*) FROM bills WHERE date=?", (date.today().isoformat(),)).fetchone()[0]
    conn.close()
    return f"CC-{today}-{str(count+1).zfill(3)}"

def gen_order_number():
    today = date.today().strftime('%Y%m%d')
    conn = get_db()
    count = conn.execute("SELECT COUNT(*) FROM orders WHERE date(created_at)=?", (date.today().isoformat(),)).fetchone()[0]
    conn.close()
    return f"ORD-{today}-{str(count+1).zfill(3)}"

def add_notification(title, message, ntype='info'):
    conn = get_db()
    conn.execute("INSERT INTO notifications (title,message,type) VALUES (?,?,?)", (title,message,ntype))
    conn.commit()
    conn.close()

# ─── SETTINGS HELPERS ──────────────────────────────────────────────────────────
def get_setting(key, default=''):
    conn = get_db()
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    conn.close()
    return row['value'] if row else default

def set_setting(key, value):
    conn = get_db()
    conn.execute("INSERT INTO settings (key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=?", (key, value, value))
    conn.commit()
    conn.close()

# ─── DYNAMIC UPI QR CODE ───────────────────────────────────────────────────────
def generate_upi_qr_base64(amount, reference):
    """Generates a UPI deep-link QR code for the exact payable amount and returns a base64 PNG data URI."""
    if not QR_AVAILABLE:
        return None
    upi_id = get_setting('upi_id', 'chatorichaat@upi')
    upi_name = get_setting('upi_name', 'Chatori Chaat')
    upi_link = f"upi://pay?pa={upi_id}&pn={upi_name.replace(' ', '%20')}&am={amount:.2f}&cu=INR&tn=ChatoriChaat-{reference}"
    img = qrcode.make(upi_link, box_size=8, border=2)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    encoded = base64.b64encode(buf.getvalue()).decode('ascii')
    return f"data:image/png;base64,{encoded}", upi_link

# ─── LOW STOCK CHECK ───────────────────────────────────────────────────────────
def check_low_stock(product_name=None):
    """Scans inventory and raises a notification for any item below its low-stock threshold.
    Avoids duplicate spam by only notifying once per item per day."""
    conn = get_db()
    q = "SELECT * FROM inventory"
    params = []
    if product_name:
        q += " WHERE product_name=?"
        params.append(product_name)
    items = conn.execute(q, params).fetchall()
    today = date.today().isoformat()
    for item in items:
        remaining = item['available_qty'] - item['used_qty']
        if remaining <= item['low_stock_threshold']:
            already = conn.execute("""SELECT COUNT(*) as c FROM notifications
                WHERE type='low_stock' AND title LIKE ? AND date(created_at)=?""",
                (f"%{item['product_name']}%", today)).fetchone()['c']
            if not already:
                level = "almost finished" if remaining <= 0 else "running low"
                conn.execute("""INSERT INTO notifications (title,message,type)
                    VALUES (?,?,?)""",
                    (f"⚠️ Low Stock Alert: {item['product_name']}",
                     f"{item['product_name']} stock is {level}. Remaining: {remaining:.1f} {item['unit']}.",
                     'low_stock'))
    conn.commit()
    conn.close()

# ─── GOOGLE REVIEWS INTEGRATION ────────────────────────────────────────────────
def fetch_google_reviews():
    """Fetches live reviews from the Google Places API if a Place ID + API key are configured
    in Admin > Settings. Returns None (graceful fallback to local reviews) if not configured
    or if the request fails for any reason."""
    place_id = get_setting('google_place_id', '')
    api_key = get_setting('google_api_key', '')
    if not place_id or not api_key:
        return None
    try:
        import requests
        url = "https://maps.googleapis.com/maps/api/place/details/json"
        resp = requests.get(url, params={
            'place_id': place_id, 'fields': 'name,rating,user_ratings_total,reviews',
            'key': api_key
        }, timeout=5)
        data = resp.json()
        if data.get('status') != 'OK':
            return None
        result = data.get('result', {})
        reviews = [{
            'name': r.get('author_name', 'Google User'),
            'rating': r.get('rating', 5),
            'comment': r.get('text', ''),
            'created_at': datetime.fromtimestamp(r.get('time', 0)).isoformat() if r.get('time') else '',
            'source': 'google'
        } for r in result.get('reviews', [])]
        return {
            'reviews': reviews,
            'avg_rating': result.get('rating', 0),
            'total_reviews': result.get('user_ratings_total', 0)
        }
    except Exception:
        return None

# ─── PUBLIC ROUTES ─────────────────────────────────────────────────────────────
@app.route('/')
def index():
    conn = get_db()
    reviews = conn.execute("SELECT * FROM reviews WHERE approved=1 ORDER BY created_at DESC LIMIT 6").fetchall()
    menu = conn.execute("SELECT * FROM menu_items WHERE available=1 LIMIT 6").fetchall()
    conn.close()
    return render_template('index.html', reviews=reviews, menu=menu)

@app.route('/about')
def about():
    conn = get_db()
    team = conn.execute("SELECT * FROM team_members ORDER BY display_order").fetchall()
    conn.close()
    return render_template('about.html', team=team)

@app.route('/menu')
def menu():
    conn = get_db()
    items = conn.execute("SELECT * FROM menu_items WHERE available=1 ORDER BY category,price").fetchall()
    conn.close()
    return render_template('menu.html', items=items, customer_logged_in=bool(session.get('customer_id')))

@app.route('/billing')
@login_required
def billing():
    conn = get_db()
    items = conn.execute("SELECT * FROM menu_items WHERE available=1 ORDER BY category,name").fetchall()
    bill_number = gen_bill_number()
    conn.close()
    return render_template('billing.html', items=items, bill_number=bill_number)

@app.route('/aradhya')
def aradhya():
    # Deity / inspiration images currently available in static/images.
    # Add more files to static/images/ and list them here to expand the gallery.
    deity_images = ['god.png', 'menu_banner.png']
    return render_template('aradhya.html', deity_images=deity_images)

@app.route('/reviews')
def reviews():
    conn = get_db()
    all_reviews = conn.execute("SELECT * FROM reviews WHERE approved=1 ORDER BY created_at DESC").fetchall()
    avg = conn.execute("SELECT AVG(rating) as avg FROM reviews WHERE approved=1").fetchone()
    count = conn.execute("SELECT COUNT(*) as cnt FROM reviews WHERE approved=1").fetchone()
    conn.close()
    google_data = fetch_google_reviews()
    return render_template('reviews.html', reviews=all_reviews,
                           avg_rating=round(avg['avg'] or 0,1), count=count['cnt'],
                           google_data=google_data,
                           google_configured=bool(get_setting('google_place_id') and get_setting('google_api_key')))

@app.route('/contact')
def contact():
    return render_template('contact.html')

@app.route('/journey')
def journey():
    conn = get_db()
    photos = conn.execute("SELECT * FROM journey_photos ORDER BY phase,date ASC").fetchall()
    conn.close()
    return render_template('journey.html', photos=photos)

# ─── PUBLIC APIs ───────────────────────────────────────────────────────────────
@app.route('/api/menu')
def api_menu():
    conn = get_db()
    items = conn.execute("SELECT * FROM menu_items WHERE available=1").fetchall()
    conn.close()
    return jsonify([dict(i) for i in items])

@app.route('/api/bill', methods=['POST'])
def save_bill():
    data = request.json
    conn = get_db()
    bill_number = data.get('bill_number') or gen_bill_number()
    today = date.today().isoformat()
    bill_date = data.get('date', today)
    cust_name = data.get('customer_name') or 'Walk-in'
    mobile = data.get('mobile','')
    grand_total = data.get('grand_total',0)

    # Upsert customer
    cust_id = None
    if mobile:
        existing = conn.execute("SELECT * FROM customers WHERE mobile=?", (mobile,)).fetchone()
        if existing:
            cust_id = existing['id']
            conn.execute("""UPDATE customers SET total_visits=total_visits+1,
                total_spent=total_spent+?, last_visit=?, name=? WHERE id=?""",
                (grand_total, bill_date, cust_name, cust_id))
        else:
            conn.execute("""INSERT INTO customers (name,mobile,total_visits,total_spent,first_visit,last_visit)
                VALUES (?,?,1,?,?,?)""", (cust_name, mobile, grand_total, bill_date, bill_date))
            cust_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

    conn.execute("""INSERT INTO bills (bill_number,customer_id,customer_name,mobile,date,
        subtotal,discount,gst,grand_total,payment_method,payment_status)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
        (bill_number, cust_id, cust_name, mobile, bill_date,
         data.get('subtotal',0), data.get('discount',0), data.get('gst',0),
         grand_total, data.get('payment_method','Cash'), data.get('payment_status','Paid')))
    bill_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

    for item in data.get('items',[]):
        conn.execute("INSERT INTO bill_items (bill_id,item_name,price,quantity,total) VALUES (?,?,?,?,?)",
                     (bill_id, item['name'], item['price'], item['qty'], item['total']))

    # Record as credit transaction
    conn.execute("""INSERT INTO transactions (type,category,description,amount,date,reference_id,payment_method)
        VALUES ('credit','Sales',?,?,?,?,?)""",
        (f"Bill {bill_number} - {cust_name}", grand_total, bill_date, bill_number,
         data.get('payment_method','Cash')))

    conn.commit()
    conn.close()

    # Notification
    add_notification('💰 New Payment Received',
        f"Bill {bill_number} | {cust_name} | ₹{grand_total:.0f} | {data.get('payment_method','Cash')}",
        'payment')

    return jsonify({'success': True, 'bill_number': bill_number})

@app.route('/api/review', methods=['POST'])
def add_review():
    data = request.json
    conn = get_db()
    conn.execute("INSERT INTO reviews (name,rating,comment) VALUES (?,?,?)",
                 (data['name'], data['rating'], data.get('comment','')))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/customer-history')
def customer_history():
    mobile = request.args.get('mobile','')
    if not mobile:
        return jsonify({'error': 'mobile required'}), 400
    conn = get_db()
    customer = conn.execute("SELECT * FROM customers WHERE mobile=?", (mobile,)).fetchone()
    if not customer:
        return jsonify({'found': False})
    bills = conn.execute("""SELECT b.*, GROUP_CONCAT(bi.item_name||' x'||bi.quantity, ', ') as items_summary
        FROM bills b LEFT JOIN bill_items bi ON b.id=bi.bill_id
        WHERE b.mobile=? GROUP BY b.id ORDER BY b.created_at DESC LIMIT 20""", (mobile,)).fetchall()
    conn.close()
    return jsonify({
        'found': True,
        'customer': dict(customer),
        'bills': [dict(b) for b in bills]
    })

# ─── AUTH ──────────────────────────────────────────────────────────────────────
@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        u = request.form.get('username')
        p = hashlib.sha256(request.form.get('password','').encode()).hexdigest()
        conn = get_db()
        admin = conn.execute("SELECT * FROM admin WHERE username=? AND password=?", (u,p)).fetchone()
        conn.close()
        if admin:
            session['admin'] = True
            return redirect(url_for('admin_dashboard'))
        return render_template('login.html', error='Invalid credentials')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

# ─── CUSTOMER AUTH (role-based, separate from Admin) ──────────────────────────
@app.route('/register', methods=['GET','POST'])
def customer_register():
    if request.method == 'POST':
        name = request.form.get('name','').strip()
        mobile = request.form.get('mobile','').strip()
        email = request.form.get('email','').strip()
        password = request.form.get('password','')
        if not name or not mobile or not password:
            return render_template('customer/register.html', error='Name, mobile and password are required.')
        if len(password) < 6:
            return render_template('customer/register.html', error='Password must be at least 6 characters.')
        conn = get_db()
        existing = conn.execute("SELECT * FROM customers WHERE mobile=?", (mobile,)).fetchone()
        if existing and existing['is_registered']:
            conn.close()
            return render_template('customer/register.html', error='An account with this mobile number already exists. Please login instead.')
        hashed = generate_password_hash(password)
        today = date.today().isoformat()
        if existing:
            conn.execute("""UPDATE customers SET name=?, email=?, password=?, is_registered=1 WHERE id=?""",
                         (name, email, hashed, existing['id']))
            cust_id = existing['id']
        else:
            conn.execute("""INSERT INTO customers (name, mobile, email, password, is_registered, first_visit, last_visit)
                VALUES (?,?,?,?,1,?,?)""", (name, mobile, email, hashed, today, today))
            cust_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.commit()
        conn.close()
        session['customer_id'] = cust_id
        session['customer_name'] = name
        return redirect(request.args.get('next') or url_for('menu'))
    return render_template('customer/register.html')

@app.route('/customer-login', methods=['GET','POST'])
def customer_login():
    if request.method == 'POST':
        mobile = request.form.get('mobile','').strip()
        password = request.form.get('password','')
        conn = get_db()
        cust = conn.execute("SELECT * FROM customers WHERE mobile=? AND is_registered=1", (mobile,)).fetchone()
        conn.close()
        if cust and cust['password'] and check_password_hash(cust['password'], password):
            session['customer_id'] = cust['id']
            session['customer_name'] = cust['name']
            return redirect(request.args.get('next') or url_for('menu'))
        return render_template('customer/login.html', error='Invalid mobile number or password.', next=request.args.get('next',''))
    return render_template('customer/login.html', next=request.args.get('next',''))

@app.route('/customer-logout')
def customer_logout():
    session.pop('customer_id', None)
    session.pop('customer_name', None)
    return redirect(url_for('index'))

@app.route('/checkout', methods=['POST'])
@customer_required
def checkout():
    """Customer places an order from their cart. Creates a Pending order and
    redirects to the dynamic QR payment page for the exact payable amount."""
    cart = request.form.get('cart_json', '[]')
    try:
        items = json.loads(cart)
    except (ValueError, TypeError):
        items = []
    if not items:
        return redirect(url_for('menu'))

    conn = get_db()
    cust = conn.execute("SELECT * FROM customers WHERE id=?", (session['customer_id'],)).fetchone()
    subtotal = sum(float(i['price']) * int(i['qty']) for i in items)
    order_number = gen_order_number()
    conn.execute("""INSERT INTO orders (order_number, customer_id, customer_name, mobile, subtotal, grand_total, payment_method, payment_status, order_status)
        VALUES (?,?,?,?,?,?,?,?,?)""",
        (order_number, cust['id'], cust['name'], cust['mobile'], subtotal, subtotal, 'UPI', 'Pending', 'Placed'))
    order_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    for i in items:
        conn.execute("INSERT INTO order_items (order_id, item_name, price, quantity, total) VALUES (?,?,?,?,?)",
                     (order_id, i['name'], float(i['price']), int(i['qty']), float(i['price'])*int(i['qty'])))
    conn.commit()
    conn.close()
    return redirect(url_for('order_pay', order_id=order_id))

@app.route('/order/<int:order_id>/pay')
@customer_required
def order_pay(order_id):
    conn = get_db()
    order = conn.execute("SELECT * FROM orders WHERE id=? AND customer_id=?", (order_id, session['customer_id'])).fetchone()
    items = conn.execute("SELECT * FROM order_items WHERE order_id=?", (order_id,)).fetchall()
    conn.close()
    if not order:
        return redirect(url_for('menu'))
    if order['payment_status'] == 'Paid':
        return redirect(url_for('order_success', order_id=order_id))
    qr_data, upi_link = generate_upi_qr_base64(order['grand_total'], order['order_number']) or (None, None)
    return render_template('customer/checkout.html', order=order, items=items,
                           qr_image=qr_data, upi_link=upi_link,
                           upi_id=get_setting('upi_id'), qr_available=QR_AVAILABLE)

@app.route('/order/<int:order_id>/confirm-payment', methods=['POST'])
@customer_required
def confirm_payment(order_id):
    """Demo payment confirmation. In production, replace this with a server-side webhook
    from your UPI payment gateway (e.g. Razorpay/PhonePe/Paytm) so payment status is verified
    automatically rather than self-reported by the customer."""
    conn = get_db()
    order = conn.execute("SELECT * FROM orders WHERE id=? AND customer_id=?", (order_id, session['customer_id'])).fetchone()
    if not order:
        conn.close()
        return redirect(url_for('menu'))
    if order['payment_status'] != 'Paid':
        upi_ref = request.form.get('upi_ref', '').strip() or f"DEMO{order_id}{int(datetime.now().timestamp())}"
        bill_number = gen_bill_number()
        today = date.today().isoformat()
        conn.execute("""INSERT INTO bills (bill_number, customer_id, customer_name, mobile, date,
            subtotal, discount, gst, grand_total, payment_method, payment_status, notes)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
            (bill_number, order['customer_id'], order['customer_name'], order['mobile'], today,
             order['subtotal'], order['discount'], order['gst'], order['grand_total'], 'UPI', 'Paid',
             f"Online order {order['order_number']}"))
        bill_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        items = conn.execute("SELECT * FROM order_items WHERE order_id=?", (order_id,)).fetchall()
        for it in items:
            conn.execute("INSERT INTO bill_items (bill_id,item_name,price,quantity,total) VALUES (?,?,?,?,?)",
                         (bill_id, it['item_name'], it['price'], it['quantity'], it['total']))
        conn.execute("""UPDATE orders SET payment_status='Paid', upi_ref=?, bill_id=? WHERE id=?""",
                     (upi_ref, bill_id, order_id))
        existing_cust = conn.execute("SELECT * FROM customers WHERE id=?", (order['customer_id'],)).fetchone()
        conn.execute("""UPDATE customers SET total_visits=total_visits+1, total_spent=total_spent+?, last_visit=? WHERE id=?""",
                     (order['grand_total'], today, order['customer_id']))
        conn.execute("""INSERT INTO transactions (type,category,description,amount,date,reference_id,payment_method)
            VALUES ('credit','Sales',?,?,?,?,?)""",
            (f"Online order {order['order_number']} - {order['customer_name']}", order['grand_total'], today, bill_number, 'UPI'))
        conn.commit()
        add_notification('🆕 New Online Order', f"{order['order_number']} | {order['customer_name']} | ₹{order['grand_total']:.0f} | UPI Paid", 'order')
    conn.close()
    return redirect(url_for('order_success', order_id=order_id))

@app.route('/order/<int:order_id>/success')
@customer_required
def order_success(order_id):
    conn = get_db()
    order = conn.execute("SELECT * FROM orders WHERE id=? AND customer_id=?", (order_id, session['customer_id'])).fetchone()
    items = conn.execute("SELECT * FROM order_items WHERE order_id=?", (order_id,)).fetchall()
    conn.close()
    if not order:
        return redirect(url_for('menu'))
    return render_template('customer/order_success.html', order=order, items=items)

@app.route('/my-orders')
@customer_required
def my_orders():
    conn = get_db()
    orders = conn.execute("""SELECT o.*, GROUP_CONCAT(oi.item_name||' x'||oi.quantity, ', ') as items_summary
        FROM orders o LEFT JOIN order_items oi ON o.id=oi.order_id
        WHERE o.customer_id=? GROUP BY o.id ORDER BY o.created_at DESC""", (session['customer_id'],)).fetchall()
    conn.close()
    return render_template('customer/my_orders.html', orders=orders)

# ─── ADMIN ROUTES ──────────────────────────────────────────────────────────────
@app.route('/admin')
@login_required
def admin_dashboard():
    conn = get_db()
    today = date.today().isoformat()
    total_sales = conn.execute("SELECT COALESCE(SUM(grand_total),0) as t FROM bills").fetchone()['t']
    total_orders = conn.execute("SELECT COUNT(*) as c FROM bills").fetchone()['c']
    total_expenses = conn.execute("SELECT COALESCE(SUM(amount),0) as t FROM expenses").fetchone()['t']
    total_customers = conn.execute("SELECT COUNT(*) as c FROM customers").fetchone()['c']
    today_sales = conn.execute("SELECT COALESCE(SUM(grand_total),0) as t FROM bills WHERE date=?", (today,)).fetchone()['t']
    today_orders = conn.execute("SELECT COUNT(*) as c FROM bills WHERE date=?", (today,)).fetchone()['c']
    recent_bills = conn.execute("SELECT * FROM bills ORDER BY created_at DESC LIMIT 10").fetchall()
    unread_notifs = conn.execute("SELECT COUNT(*) as c FROM notifications WHERE read_status=0").fetchone()['c']
    pending_orders = conn.execute("SELECT COUNT(*) as c FROM orders WHERE order_status NOT IN ('Completed','Cancelled')").fetchone()['c']
    inv = conn.execute("SELECT * FROM inventory").fetchall()
    total_stock = sum(i['available_qty'] for i in inv)
    used_stock = sum(i['used_qty'] for i in inv)
    remaining_stock = total_stock - used_stock
    stock_value = sum((i['available_qty'] - i['used_qty']) * i['purchase_cost'] for i in inv)
    low_stock_count = sum(1 for i in inv if (i['available_qty'] - i['used_qty']) <= i['low_stock_threshold'])
    conn.close()
    return render_template('admin/dashboard.html',
        total_sales=total_sales, total_orders=total_orders,
        total_expenses=total_expenses, total_profit=total_sales-total_expenses,
        total_customers=total_customers, today_sales=today_sales, today_orders=today_orders,
        recent_bills=recent_bills, unread_notifs=unread_notifs, pending_orders=pending_orders,
        total_stock=total_stock, used_stock=used_stock, remaining_stock=remaining_stock,
        stock_value=stock_value, low_stock_count=low_stock_count)

@app.route('/admin/menu')
@login_required
def admin_menu():
    conn = get_db()
    items = conn.execute("SELECT * FROM menu_items ORDER BY category,name").fetchall()
    conn.close()
    return render_template('admin/menu.html', items=items)

@app.route('/admin/cashbook')
@login_required
def admin_cashbook():
    conn = get_db()
    today = date.today().isoformat()
    daily_sales = conn.execute("SELECT COALESCE(SUM(grand_total),0) as t FROM bills WHERE date=?", (today,)).fetchone()['t']
    daily_exp = conn.execute("SELECT COALESCE(SUM(amount),0) as t FROM expenses WHERE date=?", (today,)).fetchone()['t']
    expenses = conn.execute("SELECT * FROM expenses ORDER BY date DESC LIMIT 100").fetchall()
    bills = conn.execute("SELECT * FROM bills ORDER BY created_at DESC LIMIT 30").fetchall()
    conn.close()
    return render_template('admin/cashbook.html',
        daily_sales=daily_sales, daily_exp=daily_exp,
        daily_profit=daily_sales-daily_exp,
        expenses=expenses, bills=bills)

@app.route('/admin/analytics')
@login_required
def admin_analytics():
    return render_template('admin/analytics.html')

@app.route('/admin/reviews')
@login_required
def admin_reviews():
    conn = get_db()
    reviews = conn.execute("SELECT * FROM reviews ORDER BY created_at DESC").fetchall()
    conn.close()
    return render_template('admin/reviews.html', reviews=reviews)

@app.route('/admin/bills')
@login_required
def admin_bills():
    conn = get_db()
    date_filter = request.args.get('date','')
    mobile_filter = request.args.get('mobile','')
    payment_filter = request.args.get('payment','')
    q = "SELECT b.*, GROUP_CONCAT(bi.item_name||' x'||bi.quantity, ', ') as items_list FROM bills b LEFT JOIN bill_items bi ON b.id=bi.bill_id WHERE 1=1"
    params = []
    if date_filter: q += " AND b.date=?"; params.append(date_filter)
    if mobile_filter: q += " AND b.mobile LIKE ?"; params.append(f'%{mobile_filter}%')
    if payment_filter: q += " AND b.payment_method=?"; params.append(payment_filter)
    q += " GROUP BY b.id ORDER BY b.created_at DESC LIMIT 200"
    bills = conn.execute(q, params).fetchall()
    conn.close()
    return render_template('admin/bills.html', bills=bills,
                           date_filter=date_filter, mobile_filter=mobile_filter, payment_filter=payment_filter)

@app.route('/admin/customers')
@login_required
def admin_customers():
    conn = get_db()
    customers = conn.execute("SELECT * FROM customers ORDER BY total_spent DESC").fetchall()
    conn.close()
    return render_template('admin/customers.html', customers=customers)

@app.route('/admin/notifications')
@login_required
def admin_notifications():
    conn = get_db()
    conn.execute("UPDATE notifications SET read_status=1")
    notifs = conn.execute("SELECT * FROM notifications ORDER BY created_at DESC LIMIT 100").fetchall()
    conn.commit()
    conn.close()
    return render_template('admin/notifications.html', notifs=notifs)

@app.route('/admin/financial')
@login_required
def admin_financial():
    return render_template('admin/financial.html')

@app.route('/admin/orders')
@login_required
def admin_orders():
    conn = get_db()
    status_filter = request.args.get('status','')
    mobile_filter = request.args.get('mobile','')
    q = """SELECT o.*, GROUP_CONCAT(oi.item_name||' x'||oi.quantity, ', ') as items_summary
        FROM orders o LEFT JOIN order_items oi ON o.id=oi.order_id WHERE 1=1"""
    params = []
    if status_filter: q += " AND o.order_status=?"; params.append(status_filter)
    if mobile_filter: q += " AND o.mobile LIKE ?"; params.append(f'%{mobile_filter}%')
    q += " GROUP BY o.id ORDER BY o.created_at DESC LIMIT 200"
    orders = conn.execute(q, params).fetchall()
    conn.close()
    return render_template('admin/orders.html', orders=orders, status_filter=status_filter, mobile_filter=mobile_filter)

@app.route('/admin/inventory')
@login_required
def admin_inventory():
    conn = get_db()
    items = conn.execute("SELECT * FROM inventory ORDER BY product_name").fetchall()
    conn.close()
    total_stock = sum(i['available_qty'] for i in items)
    used_stock = sum(i['used_qty'] for i in items)
    remaining_stock = total_stock - used_stock
    stock_value = sum((i['available_qty'] - i['used_qty']) * i['purchase_cost'] for i in items)
    return render_template('admin/inventory.html', items=items, total_stock=total_stock,
                           used_stock=used_stock, remaining_stock=remaining_stock, stock_value=stock_value)

@app.route('/admin/payments')
@login_required
def admin_payments():
    conn = get_db()
    mobile_filter = request.args.get('mobile','')
    date_filter = request.args.get('date','')
    base_q = "SELECT * FROM bills WHERE payment_status='Paid'"
    params = []
    if mobile_filter: base_q += " AND mobile LIKE ?"; params.append(f'%{mobile_filter}%')
    if date_filter: base_q += " AND date=?"; params.append(date_filter)
    online = conn.execute(base_q + " AND payment_method IN ('UPI') ORDER BY created_at DESC LIMIT 200", params).fetchall()
    cash = conn.execute(base_q + " AND payment_method='Cash' ORDER BY created_at DESC LIMIT 200", params).fetchall()
    card = conn.execute(base_q + " AND payment_method='Card' ORDER BY created_at DESC LIMIT 200", params).fetchall()
    totals = conn.execute("""SELECT payment_method, COALESCE(SUM(grand_total),0) as total, COUNT(*) as cnt
        FROM bills WHERE payment_status='Paid' GROUP BY payment_method""").fetchall()
    conn.close()
    return render_template('admin/payments.html', online=online, cash=cash, card=card, totals=totals,
                           mobile_filter=mobile_filter, date_filter=date_filter)

@app.route('/admin/monthly')
@login_required
def admin_monthly():
    return render_template('admin/monthly.html')

@app.route('/admin/settings', methods=['GET','POST'])
@login_required
def admin_settings():
    if request.method == 'POST':
        for key in ['upi_id','upi_name','google_place_id','google_api_key','low_stock_default_threshold']:
            if key in request.form:
                set_setting(key, request.form.get(key,'').strip())
        return redirect(url_for('admin_settings', saved=1))
    settings = {
        'upi_id': get_setting('upi_id'),
        'upi_name': get_setting('upi_name'),
        'google_place_id': get_setting('google_place_id'),
        'google_api_key': get_setting('google_api_key'),
        'low_stock_default_threshold': get_setting('low_stock_default_threshold'),
    }
    return render_template('admin/settings.html', settings=settings, saved=request.args.get('saved'))

# ─── ADMIN APIs ────────────────────────────────────────────────────────────────
@app.route('/api/analytics')
@login_required
def analytics_api():
    conn = get_db()
    total_sales = conn.execute("SELECT COALESCE(SUM(grand_total),0) as t FROM bills").fetchone()['t']
    total_orders = conn.execute("SELECT COUNT(*) as c FROM bills").fetchone()['c']
    total_expenses = conn.execute("SELECT COALESCE(SUM(amount),0) as t FROM expenses").fetchone()['t']
    total_customers = conn.execute("SELECT COUNT(*) as c FROM customers").fetchone()['c']
    monthly = conn.execute("""SELECT strftime('%Y-%m',date) as month, SUM(grand_total) as sales, COUNT(*) as orders
        FROM bills GROUP BY month ORDER BY month DESC LIMIT 12""").fetchall()
    monthly_exp = conn.execute("""SELECT strftime('%Y-%m',date) as month, SUM(amount) as expenses
        FROM expenses GROUP BY month ORDER BY month DESC LIMIT 12""").fetchall()
    daily_7 = conn.execute("""SELECT date, SUM(grand_total) as sales, COUNT(*) as orders
        FROM bills WHERE date >= date('now','-7 days') GROUP BY date ORDER BY date""").fetchall()
    conn.close()
    return jsonify({
        'total_sales': total_sales, 'total_orders': total_orders,
        'total_expenses': total_expenses, 'total_profit': total_sales-total_expenses,
        'total_customers': total_customers,
        'monthly': [dict(m) for m in monthly],
        'monthly_exp': [dict(m) for m in monthly_exp],
        'daily_7': [dict(d) for d in daily_7],
    })

@app.route('/api/financial')
@login_required
def financial_api():
    period = request.args.get('period','monthly')
    conn = get_db()
    today = date.today()

    if period == 'daily':
        start = today.isoformat()
        end = today.isoformat()
    elif period == 'weekly':
        start = (today - timedelta(days=7)).isoformat()
        end = today.isoformat()
    elif period == 'monthly':
        start = today.replace(day=1).isoformat()
        end = today.isoformat()
    else:  # yearly
        start = today.replace(month=1,day=1).isoformat()
        end = today.isoformat()

    revenue = conn.execute("SELECT COALESCE(SUM(grand_total),0) as t FROM bills WHERE date BETWEEN ? AND ?",
                           (start,end)).fetchone()['t']
    expenses = conn.execute("SELECT COALESCE(SUM(amount),0) as t FROM expenses WHERE date BETWEEN ? AND ?",
                            (start,end)).fetchone()['t']
    orders = conn.execute("SELECT COUNT(*) as c FROM bills WHERE date BETWEEN ? AND ?",
                          (start,end)).fetchone()['c']

    # Monthly breakdown
    months = conn.execute("""SELECT strftime('%Y-%m',date) as month,
        SUM(grand_total) as revenue, COUNT(*) as orders
        FROM bills WHERE date >= date('now','-12 months')
        GROUP BY month ORDER BY month""").fetchall()
    exp_months = conn.execute("""SELECT strftime('%Y-%m',date) as month, SUM(amount) as expenses
        FROM expenses WHERE date >= date('now','-12 months')
        GROUP BY month ORDER BY month""").fetchall()

    # Payment method breakdown
    payment_breakdown = conn.execute("""SELECT payment_method, SUM(grand_total) as total, COUNT(*) as count
        FROM bills WHERE date BETWEEN ? AND ? GROUP BY payment_method""", (start,end)).fetchall()

    # Expense categories
    exp_categories = conn.execute("""SELECT category, SUM(amount) as total
        FROM expenses WHERE date BETWEEN ? AND ? GROUP BY category""", (start,end)).fetchall()

    conn.close()
    profit = revenue - expenses
    return jsonify({
        'period': period, 'start': start, 'end': end,
        'revenue': revenue, 'expenses': expenses,
        'profit': profit, 'orders': orders,
        'profit_pct': round((profit/revenue*100) if revenue else 0, 1),
        'monthly': [dict(m) for m in months],
        'exp_monthly': [dict(m) for m in exp_months],
        'payment_breakdown': [dict(p) for p in payment_breakdown],
        'exp_categories': [dict(e) for e in exp_categories],
    })

@app.route('/api/notifications')
@login_required
def notif_api():
    conn = get_db()
    unread = conn.execute("SELECT COUNT(*) as c FROM notifications WHERE read_status=0").fetchone()['c']
    recent = conn.execute("SELECT * FROM notifications ORDER BY created_at DESC LIMIT 5").fetchall()
    conn.close()
    return jsonify({'unread': unread, 'notifications': [dict(n) for n in recent]})

@app.route('/api/admin/menu', methods=['POST','PUT','DELETE'])
@login_required
def admin_menu_api():
    conn = get_db()
    if request.method == 'POST':
        d = request.json
        conn.execute("INSERT INTO menu_items (name,price,category,description) VALUES (?,?,?,?)",
                     (d['name'], d['price'], d.get('category','Special'), d.get('description','')))
    elif request.method == 'PUT':
        d = request.json
        conn.execute("UPDATE menu_items SET name=?,price=?,category=?,available=? WHERE id=?",
                     (d['name'],d['price'],d['category'],d['available'],d['id']))
    elif request.method == 'DELETE':
        conn.execute("DELETE FROM menu_items WHERE id=?", (request.json['id'],))
    conn.commit(); conn.close()
    return jsonify({'success': True})

@app.route('/api/expenses', methods=['GET','POST'])
@login_required
def expenses_api():
    conn = get_db()
    if request.method == 'POST':
        d = request.json
        conn.execute("INSERT INTO expenses (date,category,description,amount) VALUES (?,?,?,?)",
                     (d['date'],d['category'],d.get('description',''),d['amount']))
        conn.execute("""INSERT INTO transactions (type,category,description,amount,date)
            VALUES ('debit',?,?,?,?)""", (d['category'],d.get('description',''),d['amount'],d['date']))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    exps = conn.execute("SELECT * FROM expenses ORDER BY date DESC LIMIT 100").fetchall()
    conn.close()
    return jsonify([dict(e) for e in exps])

@app.route('/api/admin/review/<int:rid>', methods=['PUT','DELETE'])
@login_required
def admin_review_action(rid):
    conn = get_db()
    if request.method == 'PUT':
        conn.execute("UPDATE reviews SET approved=? WHERE id=?", (request.json['approved'],rid))
    else:
        conn.execute("DELETE FROM reviews WHERE id=?", (rid,))
    conn.commit(); conn.close()
    return jsonify({'success': True})

@app.route('/api/admin/order/<int:order_id>', methods=['PUT'])
@login_required
def admin_order_update(order_id):
    d = request.json
    conn = get_db()
    if 'order_status' in d:
        conn.execute("UPDATE orders SET order_status=? WHERE id=?", (d['order_status'], order_id))
    if 'payment_status' in d:
        conn.execute("UPDATE orders SET payment_status=? WHERE id=?", (d['payment_status'], order_id))
    conn.commit(); conn.close()
    return jsonify({'success': True})

@app.route('/api/inventory', methods=['GET','POST','PUT','DELETE'])
@login_required
def inventory_api():
    conn = get_db()
    if request.method == 'POST':
        d = request.json
        conn.execute("""INSERT INTO inventory (product_name, unit, available_qty, used_qty, purchase_cost, selling_price, low_stock_threshold)
            VALUES (?,?,?,?,?,?,?)""",
            (d['product_name'], d.get('unit','kg'), d.get('available_qty',0), d.get('used_qty',0),
             d.get('purchase_cost',0), d.get('selling_price',0), d.get('low_stock_threshold',5)))
        conn.commit()
        check_low_stock(d['product_name'])
        conn.close()
        return jsonify({'success': True})
    elif request.method == 'PUT':
        d = request.json
        conn.execute("""UPDATE inventory SET product_name=?, unit=?, available_qty=?, used_qty=?,
            purchase_cost=?, selling_price=?, low_stock_threshold=?, updated_at=CURRENT_TIMESTAMP WHERE id=?""",
            (d['product_name'], d.get('unit','kg'), d.get('available_qty',0), d.get('used_qty',0),
             d.get('purchase_cost',0), d.get('selling_price',0), d.get('low_stock_threshold',5), d['id']))
        conn.commit()
        check_low_stock(d['product_name'])
        conn.close()
        return jsonify({'success': True})
    elif request.method == 'DELETE':
        conn.execute("DELETE FROM inventory WHERE id=?", (request.json['id'],))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    items = conn.execute("SELECT * FROM inventory ORDER BY product_name").fetchall()
    conn.close()
    return jsonify([dict(i) for i in items])

@app.route('/api/admin/monthly')
@login_required
def admin_monthly_api():
    """Detailed Monthly Profit & Loss analytics — admin only."""
    month_param = request.args.get('month')  # format YYYY-MM, defaults to current month
    conn = get_db()
    if month_param:
        year, mon = month_param.split('-')
        start = date(int(year), int(mon), 1)
    else:
        start = date.today().replace(day=1)
    if start.month == 12:
        end = date(start.year+1, 1, 1) - timedelta(days=1)
    else:
        end = date(start.year, start.month+1, 1) - timedelta(days=1)
    s, e = start.isoformat(), end.isoformat()

    revenue = conn.execute("SELECT COALESCE(SUM(grand_total),0) as t FROM bills WHERE date BETWEEN ? AND ?", (s,e)).fetchone()['t']
    expenses = conn.execute("SELECT COALESCE(SUM(amount),0) as t FROM expenses WHERE date BETWEEN ? AND ?", (s,e)).fetchone()['t']
    orders_count = conn.execute("SELECT COUNT(*) as c FROM bills WHERE date BETWEEN ? AND ?", (s,e)).fetchone()['c']
    customers_count = conn.execute("""SELECT COUNT(DISTINCT mobile) as c FROM bills
        WHERE date BETWEEN ? AND ? AND mobile != ''""", (s,e)).fetchone()['c']

    best_selling = conn.execute("""SELECT bi.item_name, SUM(bi.total) as revenue, SUM(bi.quantity) as qty
        FROM bill_items bi JOIN bills b ON bi.bill_id=b.id
        WHERE b.date BETWEEN ? AND ? GROUP BY bi.item_name ORDER BY revenue DESC LIMIT 1""", (s,e)).fetchone()
    most_ordered = conn.execute("""SELECT bi.item_name, SUM(bi.quantity) as qty
        FROM bill_items bi JOIN bills b ON bi.bill_id=b.id
        WHERE b.date BETWEEN ? AND ? GROUP BY bi.item_name ORDER BY qty DESC LIMIT 1""", (s,e)).fetchone()
    highest_rev_day = conn.execute("""SELECT date, SUM(grand_total) as total FROM bills
        WHERE date BETWEEN ? AND ? GROUP BY date ORDER BY total DESC LIMIT 1""", (s,e)).fetchone()
    highest_exp_day = conn.execute("""SELECT date, SUM(amount) as total FROM expenses
        WHERE date BETWEEN ? AND ? GROUP BY date ORDER BY total DESC LIMIT 1""", (s,e)).fetchone()
    days_in_month = (end - start).days + 1
    daily_trend = conn.execute("""SELECT date, SUM(grand_total) as revenue FROM bills
        WHERE date BETWEEN ? AND ? GROUP BY date ORDER BY date""", (s,e)).fetchall()
    exp_categories = conn.execute("""SELECT category, SUM(amount) as total FROM expenses
        WHERE date BETWEEN ? AND ? GROUP BY category""", (s,e)).fetchall()
    conn.close()

    profit = revenue - expenses
    return jsonify({
        'month': start.strftime('%B %Y'), 'start': s, 'end': e,
        'revenue': revenue, 'expenses': expenses, 'profit': profit,
        'orders': orders_count, 'customers': customers_count,
        'is_profit': profit >= 0,
        'best_selling_product': dict(best_selling) if best_selling else None,
        'most_ordered_product': dict(most_ordered) if most_ordered else None,
        'highest_revenue_day': dict(highest_rev_day) if highest_rev_day else None,
        'highest_expense_day': dict(highest_exp_day) if highest_exp_day else None,
        'avg_daily_revenue': round(revenue/days_in_month, 2),
        'avg_daily_profit': round(profit/days_in_month, 2),
        'daily_trend': [dict(d) for d in daily_trend],
        'exp_categories': [dict(c) for c in exp_categories],
    })

# ─── EXPORT APIs ───────────────────────────────────────────────────────────────
@app.route('/api/export/bills/csv')
@login_required
def export_bills_csv():
    conn = get_db()
    bills = conn.execute("""SELECT b.bill_number, b.customer_name, b.mobile, b.date,
        b.subtotal, b.discount, b.gst, b.grand_total, b.payment_method, b.payment_status,
        GROUP_CONCAT(bi.item_name||' x'||bi.quantity, '; ') as items
        FROM bills b LEFT JOIN bill_items bi ON b.id=bi.bill_id
        GROUP BY b.id ORDER BY b.created_at DESC""").fetchall()
    conn.close()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Bill No','Customer','Mobile','Date','Subtotal','Discount','GST','Grand Total','Payment','Status','Items'])
    for b in bills:
        writer.writerow([b['bill_number'],b['customer_name'],b['mobile'],b['date'],
                         b['subtotal'],b['discount'],b['gst'],b['grand_total'],
                         b['payment_method'],b['payment_status'],b['items']])
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode()), mimetype='text/csv',
                     as_attachment=True, download_name='chatori_chaat_bills.csv')

@app.route('/api/export/expenses/csv')
@login_required
def export_expenses_csv():
    conn = get_db()
    exps = conn.execute("SELECT * FROM expenses ORDER BY date DESC").fetchall()
    conn.close()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Date','Category','Description','Amount'])
    for e in exps:
        writer.writerow([e['date'],e['category'],e['description'],e['amount']])
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode()), mimetype='text/csv',
                     as_attachment=True, download_name='chatori_chaat_expenses.csv')

@app.route('/api/export/payments/csv')
@login_required
def export_payments_csv():
    ptype = request.args.get('type','')  # 'online' (UPI), 'cash', or '' for all
    conn = get_db()
    q = "SELECT bill_number, customer_name, mobile, date, grand_total, payment_method FROM bills WHERE payment_status='Paid'"
    params = []
    if ptype == 'online':
        q += " AND payment_method='UPI'"
    elif ptype == 'cash':
        q += " AND payment_method='Cash'"
    q += " ORDER BY created_at DESC"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Bill No','Customer Name','Mobile','Date','Amount','Payment Method'])
    for r in rows:
        writer.writerow([r['bill_number'], r['customer_name'], r['mobile'], r['date'], r['grand_total'], r['payment_method']])
    output.seek(0)
    fname = f"chatori_chaat_payments_{ptype or 'all'}.csv"
    return send_file(io.BytesIO(output.getvalue().encode()), mimetype='text/csv',
                     as_attachment=True, download_name=fname)

def _get_monthly_pnl_data(month_param):
    conn = get_db()
    if month_param:
        year, mon = month_param.split('-')
        start = date(int(year), int(mon), 1)
    else:
        start = date.today().replace(day=1)
    if start.month == 12:
        end = date(start.year+1, 1, 1) - timedelta(days=1)
    else:
        end = date(start.year, start.month+1, 1) - timedelta(days=1)
    s, e = start.isoformat(), end.isoformat()
    revenue = conn.execute("SELECT COALESCE(SUM(grand_total),0) as t FROM bills WHERE date BETWEEN ? AND ?", (s,e)).fetchone()['t']
    expenses = conn.execute("SELECT COALESCE(SUM(amount),0) as t FROM expenses WHERE date BETWEEN ? AND ?", (s,e)).fetchone()['t']
    orders_count = conn.execute("SELECT COUNT(*) as c FROM bills WHERE date BETWEEN ? AND ?", (s,e)).fetchone()['c']
    bills = conn.execute("SELECT * FROM bills WHERE date BETWEEN ? AND ? ORDER BY date", (s,e)).fetchall()
    exps = conn.execute("SELECT * FROM expenses WHERE date BETWEEN ? AND ? ORDER BY date", (s,e)).fetchall()
    conn.close()
    return {
        'month': start.strftime('%B %Y'), 'start': s, 'end': e,
        'revenue': revenue, 'expenses': expenses, 'profit': revenue - expenses,
        'orders': orders_count, 'bills': bills, 'exps': exps
    }

@app.route('/api/export/monthly/csv')
@login_required
def export_monthly_csv():
    data = _get_monthly_pnl_data(request.args.get('month'))
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Chatori Chaat - Monthly P&L Report', data['month']])
    writer.writerow([])
    writer.writerow(['Revenue', data['revenue']])
    writer.writerow(['Expenses', data['expenses']])
    writer.writerow(['Profit/Loss', data['profit']])
    writer.writerow(['Total Orders', data['orders']])
    writer.writerow([])
    writer.writerow(['--- Bills ---'])
    writer.writerow(['Bill No','Customer','Mobile','Date','Amount','Payment','Status'])
    for b in data['bills']:
        writer.writerow([b['bill_number'],b['customer_name'],b['mobile'],b['date'],b['grand_total'],b['payment_method'],b['payment_status']])
    writer.writerow([])
    writer.writerow(['--- Expenses ---'])
    writer.writerow(['Date','Category','Description','Amount'])
    for ex in data['exps']:
        writer.writerow([ex['date'],ex['category'],ex['description'],ex['amount']])
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode()), mimetype='text/csv',
                     as_attachment=True, download_name=f"chatori_chaat_monthly_{data['start'][:7]}.csv")

@app.route('/api/export/monthly/excel')
@login_required
def export_monthly_excel():
    if not EXCEL_AVAILABLE:
        return jsonify({'error': 'Excel export requires the openpyxl package. Run: pip install openpyxl'}), 500
    data = _get_monthly_pnl_data(request.args.get('month'))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Monthly P&L"
    header_fill = PatternFill(start_color="FF6B00", end_color="FF6B00", fill_type="solid")
    bold = Font(bold=True)
    ws.append(["Chatori Chaat — Monthly P&L Report", data['month']])
    ws.append([])
    ws.append(["Revenue", data['revenue']])
    ws.append(["Expenses", data['expenses']])
    ws.append(["Profit/Loss", data['profit']])
    ws.append(["Total Orders", data['orders']])
    ws.append([])
    ws.append(["Bill No","Customer","Mobile","Date","Amount","Payment","Status"])
    for cell in ws[8]:
        cell.font = bold; cell.fill = header_fill
    for b in data['bills']:
        ws.append([b['bill_number'],b['customer_name'],b['mobile'],b['date'],b['grand_total'],b['payment_method'],b['payment_status']])
    ws2 = wb.create_sheet("Expenses")
    ws2.append(["Date","Category","Description","Amount"])
    for cell in ws2[1]:
        cell.font = bold; cell.fill = header_fill
    for ex in data['exps']:
        ws2.append([ex['date'],ex['category'],ex['description'],ex['amount']])
    for sheet in (ws, ws2):
        for col in sheet.columns:
            max_len = max((len(str(c.value)) for c in col if c.value is not None), default=10)
            sheet.column_dimensions[col[0].column_letter].width = min(max_len+2, 35)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f"chatori_chaat_monthly_{data['start'][:7]}.xlsx")

@app.route('/api/export/monthly/pdf')
@login_required
def export_monthly_pdf():
    if not PDF_AVAILABLE:
        return jsonify({'error': 'PDF export requires the reportlab package. Run: pip install reportlab'}), 500
    data = _get_monthly_pnl_data(request.args.get('month'))
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=20*mm, bottomMargin=20*mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('title', parent=styles['Heading1'], textColor=colors.HexColor('#FF6B00'))
    elements = [
        Paragraph("Chatori Chaat — Monthly P&L Report", title_style),
        Paragraph(data['month'], styles['Normal']),
        Spacer(1, 10),
    ]
    summary = [
        ['Revenue', f"₹{data['revenue']:.2f}"],
        ['Expenses', f"₹{data['expenses']:.2f}"],
        ['Profit / Loss', f"₹{data['profit']:.2f}"],
        ['Total Orders', str(data['orders'])],
    ]
    t = Table(summary, colWidths=[200, 200])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#FFF3E0')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    elements += [t, Spacer(1, 16), Paragraph("Bills", styles['Heading3'])]
    bill_rows = [['Bill No','Customer','Mobile','Date','Amount','Payment','Status']]
    for b in data['bills'][:200]:
        bill_rows.append([b['bill_number'], b['customer_name'] or '', b['mobile'] or '', b['date'],
                          f"₹{b['grand_total']:.0f}", b['payment_method'], b['payment_status']])
    bt = Table(bill_rows, repeatRows=1)
    bt.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#FF6B00')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.4, colors.grey),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(bt)
    doc.build(elements)
    buf.seek(0)
    return send_file(buf, mimetype='application/pdf', as_attachment=True,
                     download_name=f"chatori_chaat_monthly_{data['start'][:7]}.pdf")

if __name__ == '__main__':
    init_db()
    app.run(debug=True, port=5000)
