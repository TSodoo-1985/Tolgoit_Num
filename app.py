import os
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)

# --- ӨГӨГДЛИЙН САНГИЙН ТОХИРГОО ---
basedir = os.path.abspath(os.path.dirname(__file__))
raw_db_url = os.environ.get('DATABASE_URL', '').strip()

if raw_db_url:
    # Илүүдэл тэмдэгт цэвэрлэх
    raw_db_url = raw_db_url.replace('"', '').replace("'", "")
    if raw_db_url.startswith("postgres://"):
        raw_db_url = raw_db_url.replace("postgres://", "postgresql://", 1)
    app.config['SQLALCHEMY_DATABASE_URI'] = raw_db_url
else:
    # DATABASE_URL байхгүй бол SQLite ашиглах
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'num_service.db')

# Neon-оос авсан хаягаа энд яг зөв тавиарай
app.config['SQLALCHEMY_DATABASE_URI'] = "postgresql://neondb_owner:npg_J8h1MnAQlbPK@ep-mute-river-a1c92rpd-pooler.ap-southeast-1.aws.neon.tech/neondb?sslmode=require"
app.config['SECRET_KEY'] = 'Sodoo123'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

login_manager = LoginManager(app)
login_manager.login_view = 'login'

# --- DB-Г ЭНД ЗАРЛАНА ---
db = SQLAlchemy(app)

# --- Модель хэсэг ---
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password = db.Column(db.String(150), nullable=False)
    role = db.Column(db.String(50), default='user')

    def check_password(self, password):
        return check_password_hash(self.password, password)

class Product(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    sku = db.Column(db.String(50), unique=True, nullable=False)
    name = db.Column(db.String(100), nullable=False)
    category = db.Column(db.String(50))
    cost_price = db.Column(db.Float, default=0.0)
    retail_price = db.Column(db.Float, default=0.0)
    wholesale_price = db.Column(db.Float, default=0.0)
    stock = db.Column(db.Float, default=0.0)
    is_active = db.Column(db.Boolean, default=True)

class Transaction(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'))
    type = db.Column(db.String(100))
    quantity = db.Column(db.Float)
    timestamp = db.Column(db.DateTime, default=datetime.now)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    product = db.relationship('Product', backref='transactions')
    user = db.relationship('User', backref='transactions')

class Expense(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(50)) # 'Зардал' эсвэл 'Ажлын хөлс'
    description = db.Column(db.String(200))
    amount = db.Column(db.Float, default=0.0)
    date = db.Column(db.DateTime, default=datetime.now)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# Өгөгдлийн сан үүсгэх
with app.app_context():
    db.create_all()
    if not User.query.filter_by(username='admin').first():
        db.session.add(User(username='admin', password=generate_password_hash('admin123', method='pbkdf2:sha256'), role='admin'))
        db.session.commit()

# --- Маршрутууд ---

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user, remember=False)
            return redirect(url_for('dashboard'))
        else:
            flash('Хэрэглэгчийн нэр эсвэл нууц үг буруу байна!')
    return render_template('login.html')

@app.route('/')
@login_required
def home():
    return redirect(url_for('dashboard'))

@app.route('/dashboard')
@login_required
def dashboard():
    # .all() гэхийн оронд .filter_by(is_active=True).all() болгож шүүлтүүр нэмэв
    products = Product.query.filter_by(is_active=True).all()

    cats = [
        "Шинэ нум", "Хуучин нум", "Амортизатор", "Стермэнь",
        "Центр боолт", "Дэр", "Зэс түлк", "Пальц",
        "Хар түлк", "Шар түлк", "Ээмэг", "Толгойн боолт",
        "Босоо пальц", "Сорочик", "Бусад"
    ]
    return render_template('dashboard.html', products=products, categories=cats)

@app.route('/logout')
@login_required
def logout():
    logout_user()
    session.clear()
    return redirect(url_for('login'))

@app.route('/add-product-page')
@login_required
def add_product_page():
    cats = [
        "Шинэ нум", "Хуучин нум", "Амортизатор", "Стермэнь",
        "Центр боолт", "Дэр", "Зэс түлк", "Пальц",
        "Хар түлк", "Шар түлк", "Ээмэг", "Толгойн боолт",
        "Босоо пальц", "Сорочик", "Бусад"
    ]
    return render_template('add_product.html', categories=cats)

@app.route('/inventory')
@login_required
def inventory():
    products = Product.query.all()
    history = Transaction.query.filter(Transaction.type.like('Тооллого%')).order_by(Transaction.timestamp.desc()).limit(10).all()
    return render_template('inventory.html', products=products, history=history)

@app.route('/do_inventory', methods=['POST'])
@login_required
def do_inventory():
    p_id = request.form.get('product_id')
    new_qty = float(request.form.get('quantity') or 0)
    product = Product.query.get(p_id)
    if product:
        diff = new_qty - product.stock
        product.stock = new_qty
        diff_text = f"Зөрүү: {'+' if diff > 0 else ''}{diff}"
        db.session.add(Transaction(product_id=p_id, type=f'Тооллого ({diff_text})', quantity=new_qty, user_id=current_user.id))
        db.session.commit()
    return redirect(url_for('inventory'))

@app.route('/add_product', methods=['POST'])
@login_required
def add_product():
    sku = request.form.get('sku')
    name = request.form.get('name')
    category = request.form.get('category')
    stock_val = request.form.get('stock')
    cost_price = request.form.get('cost_price', 0)
    retail_price = request.form.get('retail_price', 0)
    wholesale_price = request.form.get('wholesale_price', 0)

    product = Product.query.filter_by(sku=sku).first()
    if product:
        product.name = name
        product.category = category
        product.cost_price = float(cost_price) if cost_price else 0
        product.retail_price = float(retail_price) if retail_price else 0
        product.wholesale_price = float(wholesale_price) if wholesale_price else 0
        if stock_val:
             qty = float(stock_val)
             product.stock += qty
             if qty > 0:
                 db.session.add(Transaction(product_id=product.id, type='Орлого', quantity=qty, user_id=current_user.id))
        flash('Барааны мэдээлэл шинэчлэгдлээ / Орлого нэмэгдлээ.')
    else:
        initial_stock = float(stock_val) if stock_val else 0
        new_product = Product(
            sku=sku, name=name, category=category,
            stock=initial_stock,
            cost_price=float(cost_price) if cost_price else 0,
            retail_price=float(retail_price) if retail_price else 0,
            wholesale_price=float(wholesale_price) if wholesale_price else 0
        )
        db.session.add(new_product)
        db.session.flush()
        if initial_stock > 0:
            db.session.add(Transaction(product_id=new_product.id, type='Орлого', quantity=initial_stock, user_id=current_user.id))
        flash('Шинэ бараа бүртгэгдлээ.')

    db.session.commit()
    return redirect(url_for('dashboard'))

@app.route('/add_transaction', methods=['POST'])
@login_required
def add_transaction():
    p_id = request.form.get('product_id')
    t_type = request.form.get('type')  # 'Орлого', 'Жижиглэн зарлага', 'Бөөний зарлага', 'шилжүүлэг', 'буцаалт'
    qty_str = request.form.get('quantity')

    if not p_id or not qty_str:
        flash("Мэдээлэл дутуу байна!")
        return redirect(request.referrer or url_for('dashboard'))

    qty = float(qty_str)
    product = Product.query.get(p_id)

    if product:
        # Логик: Төрлөөс хамаарч үлдэгдлийг өөрчлөх
        if t_type == 'Орлого' or t_type == 'буцаалт':
            product.stock += qty
        elif 'зарлага' in t_type or t_type == 'шилжүүлэг':
            product.stock -= qty

        # Гүйлгээг хадгалах
        new_tx = Transaction(
            product_id=p_id,
            type=t_type,
            quantity=qty,
            user_id=current_user.id
        )
        db.session.add(new_tx)
        db.session.commit()
        flash(f"{product.name} - {t_type} амжилттай бүртгэгдлээ.")

    # Өмнөх байсан хуудас руу нь буцаах
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/users')
@login_required
def users_list():
    if current_user.role != 'admin': return "Эрх хүрэхгүй", 403
    return render_template('users.html', users=User.query.all())

@app.route('/add_user', methods=['POST'])
@login_required
def add_user():
    if current_user.role == 'admin':
        hashed_pw = generate_password_hash(request.form.get('password'), method='pbkdf2:sha256')
        db.session.add(User(username=request.form.get('username'), password=hashed_pw, role=request.form.get('role')))
        db.session.commit()
    return redirect(url_for('users_list'))

@app.route('/delete_user/<int:id>')
@login_required
def delete_user(id):
    if current_user.role != 'admin':
        flash('Танд хэрэглэгч устгах эрх байхгүй байна!')
        return redirect(url_for('users_list'))

    user_to_delete = User.query.get_or_404(id)

    # Одоо 'Sodoo'-г устгаж болохгүй хамгаалалттай болгоно
    if user_to_delete.username == 'Sodoo':
        flash('Үндсэн админ Sodoo-г устгах боломжгүй!')
        return redirect(url_for('users_list'))

    db.session.delete(user_to_delete)
    db.session.commit()
    flash(f'Хэрэглэгч {user_to_delete.username} амжилттай устгагдлаа.')
    return redirect(url_for('users_list'))

# --- ТАЙЛАНГУУД ---

@app.route('/export-balance')
@login_required
def export_balance():
    products = Product.query.all()
    data = []
    for p in products:
        data.append({
            'Код (SKU)': p.sku,
            'Барааны нэр': p.name,
            'Төрөл': p.category,
            'Үлдэгдэл': p.stock,
            'Өртөг': p.cost_price,
            'Жижиглэн үнэ': p.retail_price,
            'Бөөний үнэ': p.wholesale_price
        })
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Үлдэгдэл')
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"Uldegdel_{datetime.now().strftime('%Y-%m-%d')}.xlsx")

@app.route('/export-inventory')
@login_required
def export_inventory():
    # Тооллогын түүхийг бүхэлд нь татах
    transactions = Transaction.query.filter(Transaction.type.like('Тооллого%')).all()
    data = []
    for t in transactions:
        data.append({
            "Огноо": t.timestamp.strftime('%Y-%m-%d %H:%M'),
            "Бараа": t.product.name,
            "Төрөл": t.type,
            "Шинэ үлдэгдэл": t.quantity,
            "Бүртгэсэн": t.user.username if t.user else "-"
        })
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Тооллого')
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="Toollogo_Taylan.xlsx")

@app.route('/export-transactions/<type>')
@login_required
def export_transactions(type=None):
    if not type:
        return redirect(url_for('dashboard'))

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    # Шүүлтүүрийн нэр бэлдэх
    date_str = "Бүх хугацаа"
    if start_date and end_date:
        date_str = f"{start_date}-аас {end_date}"
    elif start_date:
        date_str = f"{start_date}-аас хойш"
    elif end_date:
        date_str = f"{end_date}-хүртэл"

    query = Transaction.query.filter(Transaction.type == type)
    if start_date:
        query = query.filter(Transaction.timestamp >= f"{start_date} 00:00:00")
    if end_date:
        query = query.filter(Transaction.timestamp <= f"{end_date} 23:59:59")

    transactions = query.order_by(Transaction.timestamp.desc()).all()

    data = []
    for t in transactions:
        product = Product.query.get(t.product_id)
        # Борлуулсан үнийг тодорхойлох
        sell_price = product.retail_price if type == "Жижиглэн зарлага" else product.wholesale_price
        if type == "Орлого": sell_price = product.cost_price

        cost = product.cost_price or 0
        total_sell = t.quantity * (sell_price or 0)
        total_cost = t.quantity * cost
        profit = total_sell - total_cost

        data.append({
            "Огноо": t.timestamp.strftime('%Y-%m-%d %H:%M'),
            "Барааны нэр": product.name if product else "Устгагдсан",
            "SKU": product.sku if product else "-",
            "Тоо ширхэг": t.quantity,
            "Өртөг үнэ": cost,
            "Зарах үнэ": sell_price,
            "Нийт өртөг": total_cost,
            "Нийт борлуулалт": total_sell,
            "Цэвэр ашиг": profit
        })

    if not data:
        flash(f"{type} төрөлд гүйлгээ олдсонгүй.")
        return redirect(url_for('dashboard'))

    df = pd.DataFrame(data)

    # Нийт дүнгүүдийг тооцож доор нь нэмэх
    totals = {
        "Огноо": "НИЙТ ДҮН",
        "Нийт өртөг": df["Нийт өртөг"].sum(),
        "Нийт борлуулалт": df["Нийт борлуулалт"].sum(),
        "Цэвэр ашиг": df["Цэвэр ашиг"].sum()
    }
    df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Тайлан')
        workbook = writer.book
        worksheet = writer.sheets['Тайлан']

        # Мөнгөн дүнгийн формат
        fmt = workbook.add_format({'num_format': '#,##0"₮"', 'align': 'right'})
        bold_fmt = workbook.add_format({'num_format': '#,##0"₮"', 'bold': True, 'bg_color': '#D9EAD3'})

        worksheet.set_column('A:C', 20)
        worksheet.set_column('D:D', 10)
        worksheet.set_column('E:I', 15, fmt)

    output.seek(0)
    file_name = f"{type}_Тайлан_{date_str}.xlsx"
    return send_file(output, as_attachment=True, download_name=file_name)

    # Нийт борлуулалтын дүнг тооцож нэмэх
    total_sum = df["Нийт дүн"].sum()
    sum_row = pd.DataFrame([{"Огноо": "НИЙТ", "Барааны SKU": "", "Барааны нэр": "", "Тоо ширхэг": "", "Нэгж үнэ": "", "Нийт дүн": total_sum}])
    df = pd.concat([df, sum_row], ignore_index=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Тайлан')
        workbook  = writer.book
        worksheet = writer.sheets['Тайлан']
        money_fmt = workbook.add_format({'num_format': '#,##0"₮"', 'bold': True})
        worksheet.set_column('A:C', 20)
        worksheet.set_column('D:D', 12)
        worksheet.set_column('E:F', 15, money_fmt)

    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"{type}_Report.xlsx")

@app.route('/export-expense-report/<category>')
@login_required
def export_expense_report(category=None):
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    date_str = "Бүх хугацаа"
    if start_date and end_date: date_str = f"{start_date}-аас {end_date}"

    query = Expense.query.filter(Expense.category == category)
    if start_date: query = query.filter(Expense.date >= start_date)
    if end_date: query = query.filter(Expense.date <= end_date)

    items = query.all()
    data = [{"Огноо": i.date.strftime('%Y-%m-%d'), "Тайлбар": i.description, "Дүн": i.amount} for i in items]

    if not data:
        flash("Зардлын мэдээлэл олдсонгүй.")
        return redirect(url_for('expenses'))

    df = pd.DataFrame(data)
    total_row = pd.DataFrame([{"Огноо": "НИЙТ", "Тайлбар": "", "Дүн": df["Дүн"].sum()}])
    df = pd.concat([df, total_row], ignore_index=True)

    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"{category}_Тайлан_{date_str}.xlsx")
# --- Бусад хуудсууд ---

@app.route('/returns-page')
@login_required
def returns_page():
    products = Product.query.order_by(Product.name).all()
    categories = db.session.query(Product.category).distinct().all()
    categories = [c[0] for c in categories]
    return render_template('returns.html', products=products, categories=categories)

@app.route('/transfer-page')
@login_required
def transfer_page():
    products = Product.query.order_by(Product.name).all()
    categories = db.session.query(Product.category).distinct().all()
    categories = [c[0] for c in categories]
    return render_template('special_transfer.html', products=products, categories=categories)

@app.route('/expenses', methods=['GET', 'POST']) # Энд функцийн нэр нь 'expenses'
def expenses():
    if request.method == 'POST':
        cat = request.form.get('category')
        desc = request.form.get('description')
        amount = float(request.form.get('amount'))
        db.session.add(Expense(category=cat, description=desc, amount=amount, user_id=current_user.id))
        db.session.commit()
        flash('Амжилттай бүртгэгдлээ')
        return redirect(url_for('expenses'))
    items = Expense.query.order_by(Expense.date.desc()).limit(20).all()
    return render_template('expenses.html', items=items)

@app.route('/statistics')
@login_required
def statistics():
    start_date_str = request.args.get('start_date')
    end_date_str = request.args.get('end_date')

    # 1. Огноо тохируулах
    if start_date_str:
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
    else:
        start_date = datetime.now() - timedelta(days=30)

    if end_date_str:
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
    else:
        end_date = datetime.now()

    # 2. Өгөгдөл шүүх (Зассан: Transaction.timestamp ашиглав)
    transactions = Transaction.query.filter(
        Transaction.timestamp >= start_date,
        Transaction.timestamp <= end_date + timedelta(days=1)
    ).all()

    # 3. Графикийн өдрүүдийг бэлдэх
    dates_list = []
    sales_map = {}
    profit_map = {}
    top_products = {}
    returns_count = 0

    curr = start_date
    while curr <= end_date:
        d_str = curr.strftime('%m-%d')
        dates_list.append(d_str)
        sales_map[d_str] = 0
        profit_map[d_str] = 0
        curr += timedelta(days=1)

    # 4. Гүйлгээг тооцоолох (Зассан: Product моделийн үнийг ашиглав)
    for tx in transactions:
        d_str = tx.timestamp.strftime('%m-%d')
        if d_str not in sales_map: continue

        # Үнэ тодорхойлох
        if tx.type == "Жижиглэн зарлага":
            sell_price = tx.product.retail_price or 0
        elif tx.type == "Бөөний зарлага":
            sell_price = tx.product.wholesale_price or 0
        else:
            sell_price = 0

        cost_price = tx.product.cost_price or 0

        amount = tx.quantity * sell_price
        total_cost = tx.quantity * cost_price

        if "зарлага" in tx.type.lower():
            sales_map[d_str] += amount
            profit_map[d_str] += (amount - total_cost)

            p_name = tx.product.name if tx.product else "Тодорхойгүй"
            top_products[p_name] = top_products.get(p_name, 0) + tx.quantity
        elif "буцаалт" in tx.type.lower():
            returns_count += tx.quantity

    # 5. Зардал тооцоолох (Зассан: Expense.date ашиглав)
    expenses_data = Expense.query.filter(
        Expense.date >= start_date,
        Expense.date <= end_date + timedelta(days=1)
    ).all()

    expense_map = {d: 0 for d in dates_list}
    for ex in expenses_data:
        ex_str = ex.date.strftime('%m-%d')
        if ex_str in expense_map:
            expense_map[ex_str] += ex.amount

    # 6. Жагсаалт руу хөрвүүлэх
    sales_values = [sales_map[d] for d in dates_list]
    profit_values = [profit_map[d] for d in dates_list]
    expense_values = [expense_map[d] for d in dates_list]

    sorted_top = sorted(top_products.items(), key=lambda x: x[1], reverse=True)[:5]
    top_labels = [x[0] for x in sorted_top]
    top_values = [x[1] for x in sorted_top]

    return render_template('statistics.html',
                           start_date=start_date.strftime('%Y-%m-%d'),
                           end_date=end_date.strftime('%Y-%m-%d'),
                           dates=dates_list,
                           sales=sales_values,
                           profit=profit_values,
                           expenses=expense_values,
                           returns=[returns_count],
                           top_labels=top_labels,
                           top_values=top_values,
                           products=Product.query.all())

@app.route('/change_password', methods=['GET', 'POST'])
@login_required
def change_password():
    if request.method == 'POST':
        new_password = request.form.get('new_password')
        confirm_password = request.form.get('confirm_password')

        if new_password != confirm_password:
            flash('Нууц үг зөрүүтэй байна!')
            return redirect(url_for('change_password'))

        # Одоо нэвтэрсэн байгаа хэрэглэгчийн нууц үгийг шинэчлэх
        current_user.password = new_password
        db.session.commit()
        flash('Нууц үг амжилттай солигдлоо.')
        return redirect(url_for('dashboard'))

    return render_template('change_password.html')

@app.route('/export-inventory-report')
@login_required
def export_inventory_report():
    try:
        # 1. URL-аас хугацааны шүүлтүүрийг авах
        start_date = request.args.get('start_date')
        end_date = request.args.get('end_date')

        # Суурь query (Зөвхөн тооллогын гүйлгээ)
        query = Transaction.query.filter(Transaction.type.like('%Тооллого%'))

        # 2. Хугацаагаар шүүх
        if start_date:
            query = query.filter(Transaction.timestamp >= datetime.strptime(start_date, '%Y-%m-%d'))
        if end_date:
            # Дуусах огнооны 23:59:59 хүртэлх өгөгдлийг авах
            end_dt = datetime.strptime(end_date, '%Y-%m-%d').replace(hour=23, minute=59, second=59)
            query = query.filter(Transaction.timestamp <= end_dt)

        history = query.order_by(Transaction.timestamp.desc()).all()

        if not history:
            flash('Сонгосон хугацаанд тооллогын түүх байхгүй байна!')
            return redirect(url_for('inventory'))

        data = []
        for h in history:
            data.append({
                "Огноо": h.timestamp.strftime('%Y-%m-%d %H:%M'),
                "Барааны код": h.product.sku if h.product else "Устгагдсан",
                "Барааны нэр": h.product.name if h.product else "Устгагдсан",
                "Үйлдэл/Зөрүү": h.type,
                "Тоо хэмжээ": h.quantity,
                "Тоолсон ажилтан": h.user.username if h.user else "Систем"
            })

        df = pd.DataFrame(data)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Inventory_Report')
            worksheet = writer.sheets['Inventory_Report']
            for i, col in enumerate(df.columns):
                column_len = max(df[col].astype(str).str.len().max(), len(col)) + 2
                worksheet.set_column(i, i, column_len)

        output.seek(0)
        filename = f"Inventory_Report_{start_date}_to_{end_date}.xlsx" if start_date else "Inventory_Report_Full.xlsx"

        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        flash(f"Алдаа: {str(e)}")
        return redirect(url_for('inventory'))

@app.route('/edit-product/<int:id>', methods=['GET', 'POST'])
@login_required
def edit_product(id):
    product = Product.query.get_or_404(id)
    if request.method == 'POST':
        product.sku = request.form.get('sku')
        product.name = request.form.get('name')
        product.category = request.form.get('category')
        product.cost_price = float(request.form.get('cost_price') or 0)
        product.retail_price = float(request.form.get('retail_price') or 0)
        product.wholesale_price = float(request.form.get('wholesale_price') or 0)

        # product.stock хэсгийг эндээс хаслаа. Ингэснээр үлдэгдэл өөрчлөгдөхгүй.

        db.session.commit()
        flash(f'"{product.name}" барааны мэдээлэл амжилттай засагдлаа.')
        return redirect(url_for('dashboard'))

    return render_template('edit_product.html', product=product)

@app.route('/delete-product/<int:id>', methods=['POST'])
@login_required
def delete_product(id):
    if current_user.role != 'admin':
        flash('Эрх хүрэлцэхгүй байна!')
        return redirect(url_for('dashboard'))

    product = Product.query.get_or_404(id)
    # Датабаазаас устгахын оронд идэвхгүй болгоно
    product.is_active = False
    db.session.commit()

    flash(f'"{product.name}" барааг жагсаалтаас хаслаа.')
    return redirect(url_for('dashboard'))

# --- СИСТЕМ ЭХЛЭХ ХЭСЭГ ---
if __name__ == '__main__':
    with app.app_context():
        # 1. Бүх хүснэгтүүдийг Neon дээр автоматаар үүсгэнэ
        db.create_all()
        
        # 2. Анхдагч админ хэрэглэгч байхгүй бол үүсгэх (Нэмэлт хамгаалалт)
        if not User.query.filter_by(username='admin').first():
            admin_user = User(
                username='admin', 
                password=generate_password_hash('admin123', method='pbkdf2:sha256'), 
                role='admin'
            )
            db.session.add(admin_user)
            db.session.commit()
            print("--- Өгөгдлийн сан болон Админ хэрэглэгч бэлэн боллоо ---")

    # Render дээр ажиллах порт тохиргоо
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
