# 🍛 Chatori Chaat – Smart Food Stall Management System (v2)

**"Swad ka Patakha, Mohalle ka Superstar"**

A complete restaurant/food stall management website built with Flask + SQLite + Bootstrap 5 — now with role-based customer/admin login, online ordering with dynamic UPI QR payments, inventory management, low-stock alerts, and a Monthly Profit & Loss dashboard.

---

## 🚀 Quick Start (Run in 3 Steps)

### Step 1 – Install Python dependencies
```bash
pip install -r requirements.txt
```
This installs Flask plus the optional packages used for QR codes (`qrcode`), Excel export (`openpyxl`), PDF export (`reportlab`), and Google Reviews (`requests`). The app still runs even if one of these is missing — it just disables that one feature gracefully (e.g. no QR image, manual UPI ID instead).

### Step 2 – Run the app
```bash
python app.py
```

### Step 3 – Open in browser
```
http://localhost:5000
```

That's it! The SQLite database (`instance/chatori.db`) is created and seeded automatically on first run. If you're upgrading from v1, your existing database is **preserved** — the app auto-migrates it by adding the new tables/columns it needs the first time it starts.

---

## 🔑 Logins

This app now has **two completely separate login systems**:

### Admin Login (full business access)
- **URL:** http://localhost:5000/login
- **Username:** `admin`
- **Password:** `admin123`

### Customer Login (order & pay online)
- **Register:** http://localhost:5000/register
- **Login:** http://localhost:5000/customer-login
- Customers can browse the menu, add items to a cart, check out, pay via dynamic UPI QR, and view their own order history. They have **no access** to billing, inventory, revenue, or any admin pages.

> ⚠️ Change the admin password and the demo UPI ID (Admin → Settings) before going live.

---

## 📁 Project Structure

```
chatori_chaat/
├── app.py                       # Main Flask app (routes + DB + APIs)
├── requirements.txt             # Python dependencies
├── instance/
│   └── chatori.db               # SQLite database (auto-created)
├── static/
│   ├── css/style.css            # Main stylesheet
│   ├── js/main.js               # JS utilities (lightbox, notifications, reveal)
│   └── images/                  # Logo, favicon, deity/Aradhya images, gallery photos
└── templates/
    ├── base.html                # Public layout (navbar/footer, customer login menu)
    ├── index.html, about.html, menu.html, journey.html,
    │   aradhya.html, reviews.html, contact.html, billing.html (admin POS)
    ├── login.html                # Admin login
    ├── customer/                 # Customer-facing auth & ordering
    │   ├── register.html, login.html
    │   ├── checkout.html         # Dynamic UPI QR payment page
    │   ├── order_success.html, my_orders.html
    └── admin/                    # Admin panel (role-protected)
        ├── base_admin.html       # Sidebar layout + notification bell
        ├── dashboard.html        # KPIs, stock summary, low-stock banner
        ├── orders.html           # Customer online orders (status workflow)
        ├── inventory.html        # Stock CRUD, low-stock detection
        ├── payments.html         # Online (UPI) vs Cash payment history
        ├── monthly.html          # Monthly P&L dashboard + charts + exports
        ├── financial.html, analytics.html, cashbook.html
        ├── bills.html, customers.html, menu.html
        ├── reviews.html, notifications.html, settings.html
```

---

## 🌐 Pages & Features

### Public / Customer Pages
| Page | URL | Features |
|------|-----|----------|
| Home | `/` | Hero, featured menu, Aradhya teaser, reviews, CTA |
| About | `/about` | Founders + **Meet Our Team** cards (Vaibhav, Manoj, Druvita, Rupali) |
| Menu | `/menu` | Search, filter, cart → **online checkout** |
| Aradhya | `/aradhya` | आध्यात्मिक प्रेरणा gallery with lightbox |
| Journey | `/journey` | Dream & Planning → Inauguration → Growth → Future Vision |
| Reviews | `/reviews` | Star ratings + optional **live Google Reviews** |
| Contact | `/contact` | WhatsApp form, map, hours |
| Register / Login | `/register`, `/customer-login` | Customer account creation & login |
| Checkout | `/order/<id>/pay` | **Dynamic UPI QR code** for the exact order total |
| My Orders | `/my-orders` | Customer's own order & payment history |

### Admin Panel (role-protected, `/login`)
| Page | URL | Features |
|------|-----|----------|
| Dashboard | `/admin` | Revenue/expense/profit KPIs, stock summary, low-stock banner |
| Customer Orders | `/admin/orders` | All online orders, filter by status/mobile, update order status |
| Inventory & Stock | `/admin/inventory` | Add/edit/delete stock items, auto low-stock detection |
| Payment History | `/admin/payments` | Online (UPI) vs Cash vs Card, filterable, CSV export |
| Monthly P&L | `/admin/monthly` | 🟢/🔴 profit indicator, charts, best-seller analytics, PDF/Excel/CSV export |
| Financial Reports | `/admin/financial` | Daily revenue/expense/profit pie & bar charts |
| Cashbook | `/admin/cashbook` | Expense entry & tracking |
| Bill History | `/admin/bills` | Every bill ever generated |
| Customers | `/admin/customers` | Search by mobile, lifetime value, visit count |
| Menu Management | `/admin/menu` | Add/edit/delete menu items |
| Reviews | `/admin/reviews` | Approve/delete customer reviews |
| Notifications | `/admin/notifications` | Low-stock alerts, new-order alerts, unread counter |
| Settings | `/admin/settings` | UPI ID, Google Reviews API keys, low-stock defaults |

---

## 💳 Dynamic UPI QR Payment Flow

1. Customer logs in, adds items to cart on `/menu`, clicks **Checkout & Pay (QR)**.
2. An `orders` row is created with the exact total (Pending).
3. `/order/<id>/pay` renders a **freshly generated QR code** encoding a `upi://pay?...&am=<exact amount>` deep link, plus the UPI ID as a fallback for manual payment.
4. Customer scans with any UPI app (GPay/PhonePe/Paytm/BHIM) and pays the exact amount.
5. Customer clicks **"I Have Paid"** (optionally entering the UPI reference number) → a permanent `bills` record + `transactions` record is created, the order is marked Paid, and the admin gets a real-time notification.

> 🔔 **Production note:** Step 5 is a self-reported demo confirmation. For real-world deployment, replace it with a server-side webhook from a UPI payment gateway (Razorpay / PhonePe Business / Paytm for Business) so payment status is verified automatically rather than trusted from the browser.

---

## 📦 Inventory & Low Stock Alerts

- Admin → Inventory tracks **Available / Used / Remaining** quantity and **Purchase Cost / Selling Price** per ingredient.
- Whenever stock is added or edited, the app checks if remaining quantity has dropped to/below its configured threshold and — if so — raises a notification (bell icon + dashboard banner + unread counter), e.g. *"⚠️ Papdi stock is running low."*
- Notifications are de-duplicated per item per day so the bell doesn't spam.

---

## 📊 Monthly Profit & Loss Dashboard (Admin Only)

- Not linked anywhere on the public site — accessible only to a logged-in admin at `/admin/monthly`.
- Computes Monthly Revenue, Expenses, Profit/Loss, Orders, and unique Customers for any selected month.
- Shows a 🟢 *Business Running in Profit* / 🔴 *Business Running in Loss* indicator.
- Charts: Revenue vs Expense bar, Profit doughnut, Daily Sales trend line.
- Analytics: Best-Selling Product, Most-Ordered Product, Highest Revenue/Expense Day, Average Daily Revenue/Profit.
- Export to **PDF**, **Excel**, or **CSV** with one click.

---

## ⭐ Google Reviews Integration

Optional — configure under Admin → Settings:
1. Find your business's **Google Place ID** (Google's [Place ID Finder](https://developers.google.com/maps/documentation/places/web-service/place-id)).
2. Create a **Google Places API key** with billing enabled on Google Cloud Console.
3. Paste both into Admin → Settings → Google Reviews Integration.

Once configured, `/reviews` automatically fetches and displays live Google reviews (reviewer name, rating, date, text) alongside your website's own reviews. If left blank, the page gracefully falls back to showing only website reviews — nothing breaks.

---

## 🗄️ Database Tables

`admin`, `customers` (with login + registration), `menu_items`, `bills`, `bill_items`, `orders`, `order_items`, `transactions`, `expenses`, `inventory`, `notifications`, `reviews`, `team_members`, `journey_photos`, `settings` — with foreign keys, indexes on frequently-queried columns (date, mobile, customer), and server-side validation on all write endpoints.

---

## 📞 Contact Details (Stall Info)

- **Address:** Shop No. 9 Royal Complex, Sawarkar Chowk, Indira Nagar, Nashik
- **Phone:** 8421941950 / 8483852382
- **WhatsApp:** 8421941950
- **Email:** chatorichaat2305@gmail.com
- **Instagram:** [@chatori_.chaat](https://www.instagram.com/chatori_.chaat)

---

## 🔧 Tech Stack

- **Backend:** Python Flask, Werkzeug (password hashing)
- **Database:** SQLite (built-in `sqlite3`)
- **Frontend:** HTML5 + CSS3 + Bootstrap 5.3
- **Charts:** Chart.js 4
- **QR Codes:** `qrcode` + Pillow
- **Excel Export:** `openpyxl`
- **PDF Export:** `reportlab`
- **Google Reviews:** `requests` (Google Places API)
- **Icons:** Font Awesome 6
- **Fonts:** Google Fonts (Poppins + Playfair Display)

> **Note:** Uses SQLite instead of MySQL for zero-config setup. Swap to MySQL/Postgres by replacing `sqlite3` with `flask-sqlalchemy` if needed for production scale.

---

## 🌍 Deployment (PythonAnywhere – Free)

1. Sign up at [pythonanywhere.com](https://www.pythonanywhere.com)
2. Upload the project folder
3. Create a web app → Flask → Python 3.x
4. Set WSGI file to point to `app.py`
5. Run `pip install -r requirements.txt` in the console
6. In Admin → Settings, set your real UPI ID (and optionally Google Reviews keys)
7. Reload the web app

---

## 🔐 Security Notes

- Customer passwords are hashed with Werkzeug's `generate_password_hash` (never stored in plaintext).
- Admin and customer sessions are completely separate — a customer session cannot access any `/admin/*` route, and vice versa.
- All admin write APIs (`/api/...`) are protected by the `@login_required` decorator.
- Before production: change `app.secret_key` in `app.py` to a long random value, set `debug=False`, and serve behind a proper WSGI server (gunicorn/uWSGI) instead of the Flask dev server.

---

*Made with ❤️ for Piyush & Parnika Patil | Chatori Chaat, Nashik*
