import os
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)

# --- ӨГӨГДЛИЙН САНГИЙН ТОХИРГОО (NEON.TECH) ---
# Таны өгсөн Neon холболтын хаягийг энд ашиглаж байна
app.config['SQLALCHEMY_DATABASE_URI'] = "postgresql://neondb_owner:npg_J8h1MnAQlbPK@ep-mute-river-a1c92rpd-pooler.ap-southeast-1.aws.neon.tech/neondb?sslmode=require"
app.config['SECRET_KEY'] = 'Sodoo123'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {"pool_pre_ping": True, "pool_recycle": 300}

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

# --- МОДЕЛЬ ХЭСЭГ ---
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
    category = db.Column(db.String(50))
    description = db.Column(db.String(200))
    amount = db.Column(db.Float, default=0.0)
    date = db.Column(db.DateTime, default=datetime.now)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# --- ҮНДСЭН МАРШРУТУУД ---

@app.route('/')
@login_required
def home():
    return redirect(url_for('dashboard'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        user = User.query.filter_by(username=request.form.get('username')).first()
        if user and user.check_password(request.form.get('password')):
            login_user(user)
            return redirect(url_for('dashboard'))
        flash('Нэвтрэх нэр эсвэл нууц үг буруу!')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    products = Product.query.filter_by(is_active=True).all()
    cats = ["Шинэ нум", "Хуучин нум", "Амортизатор", "Стермэнь", "Центр боолт", "Дэр", "Зэс түлк", "Пальц", "Хар түлк", "Шар түлк", "Ээмэг", "Толгойн боолт", "Босоо пальц", "Сорочик", "Бусад"]
    return render_template('dashboard.html', products=products, categories=cats)

# --- БАРАА БҮРТГЭЛ, ЗАСВАР ---

@app.route('/add-product-page')
@login_required
def add_product_page():
    cats = ["Шинэ нум", "Хуучин нум", "Амортизатор", "Стермэнь", "Центр боолт", "Дэр", "Зэс түлк", "Пальц", "Хар түлк", "Шар түлк", "Ээмэг", "Толгойн боолт", "Босоо пальц", "Сорочик", "Бусад"]
    return render_template('add_product.html', categories=cats)

@app.route('/add_product', methods=['POST'])
@login_required
def add_product():
    sku = request.form.get('sku')
    product = Product.query.filter_by(sku=sku).first()
    stock_val = float(request.form.get('stock') or 0)
    
    if product:
        # Хуучин бараа бол үлдэгдэл нэмнэ
        product.stock += stock_val
        if stock_val > 0:
            db.session.add(Transaction(product_id=product.id, type='Орлого', quantity=stock_val, user_id=current_user.id))
        flash('Барааны үлдэгдэл нэмэгдлээ.')
    else:
        # Шинэ бараа үүсгэх
        new_p = Product(
            sku=sku, name=request.form.get('name'), category=request.form.get('category'), 
            stock=stock_val, 
            cost_price=float(request.form.get('cost_price') or 0),
            retail_price=float(request.form.get('retail_price') or 0), 
            wholesale_price=float(request.form.get('wholesale_price') or 0)
        )
        db.session.add(new_p)
        db.session.flush()
        if stock_val > 0:
            db.session.add(Transaction(product_id=new_p.id, type='Орлого', quantity=stock_val, user_id=current_user.id))
        flash('Шинэ бараа амжилттай бүртгэгдлээ.')
    
    db.session.commit()
    return redirect(url_for('dashboard'))

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
        db.session.commit()
        return redirect(url_for('dashboard'))
    return render_template('edit_product.html', product=product)

@app.route('/delete-product/<int:id>', methods=['POST'])
@login_required
def delete_product(id):
    if current_user.role != 'admin':
        flash('Эрх хүрэлцэхгүй!')
        return redirect(url_for('dashboard'))
    
    product = Product.query.get_or_404(id)
    product.is_active = False # Устгахын оронд идэвхгүй болгоно
    db.session.commit()
    return redirect(url_for('dashboard'))

# --- ГҮЙЛГЭЭ, ТООЛЛОГО ---

@app.route('/add_transaction', methods=['POST'])
@login_required
def add_transaction():
    p_id = request.form.get('product_id')
    t_type = request.form.get('type')
    qty = float(request.form.get('quantity') or 0)
    product = Product.query.get(p_id)
    
    if product:
        if t_type in ['Орлого', 'буцаалт']:
            product.stock += qty
        else:
            product.stock -= qty
            
        db.session.add(Transaction(product_id=p_id, type=t_type, quantity=qty, user_id=current_user.id))
        db.session.commit()
        flash(f"{product.name} - {t_type} бүртгэгдлээ.")
        
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/inventory')
@login_required
def inventory():
    products = Product.query.filter_by(is_active=True).all()
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

# --- ТАЙЛАН, СТАТИСТИК ---

@app.route('/statistics')
@login_required
def statistics():
    # Статистик тооцоолох логик (Энгийн байдлаар)
    products = Product.query.filter_by(is_active=True).all()
    # Энд датаг боловсруулаад HTML рүү дамжуулна
    return render_template('statistics.html', 
                           products=products, 
                           dates=[], sales=[], profit=[], expenses=[], returns=[0], top_labels=[], top_values=[], 
                           start_date=datetime.now().strftime('%Y-%m-%d'), end_date=datetime.now().strftime('%Y-%m-%d'))

@app.route('/expenses', methods=['GET', 'POST'])
@login_required
def expenses():
    if request.method == 'POST':
        db.session.add(Expense(
            category=request.form.get('category'), 
            description=request.form.get('description'), 
            amount=float(request.form.get('amount') or 0), 
            user_id=current_user.id
        ))
        db.session.commit()
        return redirect(url_for('expenses'))
    
    items = Expense.query.order_by(Expense.date.desc()).limit(20).all()
    return render_template('expenses.html', items=items)

@app.route('/export-balance')
@login_required
def export_balance():
    products = Product.query.filter_by(is_active=True).all()
    df = pd.DataFrame([
        {'Код': p.sku, 'Бараа': p.name, 'Үлдэгдэл': p.stock, 'Өртөг': p.cost_price, 'Жижиглэн': p.retail_price, 'Бөөний': p.wholesale_price} 
        for p in products
    ])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"Uldegdel_{datetime.now().strftime('%Y%m%d')}.xlsx")

@app.route('/export-transactions/<type>')
@login_required
def export_transactions(type):
    # Энэ функц layout.html дээр дуудагдаж байгаа тул заавал байх ёстой
    # (Хялбарчилсан хувилбар)
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    
    query = Transaction.query.filter(Transaction.type == type)
    if start_date: query = query.filter(Transaction.timestamp >= start_date)
    if end_date: query = query.filter(Transaction.timestamp <= end_date)
    
    transactions = query.all()
    
    data = []
    for t in transactions:
        data.append({
            "Огноо": t.timestamp.strftime('%Y-%m-%d %H:%M'),
            "Бараа": t.product.name if t.product else "-",
            "Тоо": t.quantity,
            "Хэн": t.user.username if t.user else "-"
        })
        
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"{type}_Report.xlsx")

@app.route('/export-expense-report/<category>')
@login_required
def export_expense_report(category):
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    
    query = Expense.query.filter(Expense.category == category)
    if start_date: query = query.filter(Expense.date >= start_date)
    if end_date: query = query.filter(Expense.date <= end_date)
    
    expenses = query.all()
    data = [{"Огноо": e.date, "Тайлбар": e.description, "Дүн": e.amount} for e in expenses]
    
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"{category}_Report.xlsx")

@app.route('/export-inventory-report')
@login_required
def export_inventory_report():
    transactions = Transaction.query.filter(Transaction.type.like('Тооллого%')).all()
    data = [{"Огноо": t.timestamp, "Бараа": t.product.name, "Зөрүү": t.type, "Тоо": t.quantity} for t in transactions]
    
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="Inventory_Report.xlsx")

# --- ХЭРЭГЛЭГЧИЙН УДИРДЛАГА ---

@app.route('/users')
@login_required
def users_list():
    if current_user.role != 'admin': return "Эрх хүрэхгүй", 403
    return render_template('users.html', users=User.query.all())

@app.route('/add_user', methods=['POST'])
@login_required
def add_user():
    if current_user.role == 'admin':
        hashed_pw = generate_password_hash(request.form.get('password'))
        db.session.add(User(username=request.form.get('username'), password=hashed_pw, role=request.form.get('role')))
        db.session.commit()
    return redirect(url_for('users_list'))

@app.route('/delete_user/<int:id>')
@login_required
def delete_user(id):
    if current_user.role != 'admin': return redirect(url_for('users_list'))
    user = User.query.get_or_404(id)
    if user.username != 'admin':
        db.session.delete(user)
        db.session.commit()
    return redirect(url_for('users_list'))

@app.route('/change_password', methods=['GET', 'POST'])
@login_required
def change_password():
    if request.method == 'POST':
        if request.form.get('new_password') == request.form.get('confirm_password'):
            current_user.password = generate_password_hash(request.form.get('new_password'))
            db.session.commit()
            flash('Нууц үг солигдлоо.')
            return redirect(url_for('dashboard'))
        flash('Нууц үг зөрүүтэй байна!')
    return render_template('change_password.html')

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        # Админ хэрэглэгч үүсгэх/шинэчлэх (admin / 123456)
        admin = User.query.filter_by(username='admin').first()
        if not admin:
            db.session.add(User(username='admin', password=generate_password_hash('123456'), role='admin'))
        else:
            admin.password = generate_password_hash('123456')
        db.session.commit()
        
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
