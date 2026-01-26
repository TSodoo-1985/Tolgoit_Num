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
app.config['SQLALCHEMY_DATABASE_URI'] = "postgresql://neondb_owner:npg_J8h1MnAQlbPK@ep-mute-river-a1c92rpd-pooler.ap-southeast-1.aws.neon.tech/neondb?sslmode=require"
app.config['SECRET_KEY'] = 'Sodoo123'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {"pool_pre_ping": True, "pool_recycle": 300}

login_manager = LoginManager(app)
login_manager.login_view = 'login'
db = SQLAlchemy(app)

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

# --- МАРШРУТУУД ---

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
    products = Product.query.filter_by(is_active=True).all()
    cats = ["Шинэ нум", "Хуучин нум", "Амортизатор", "Стермэнь", "Центр боолт", "Дэр", "Зэс түлк", "Пальц", "Хар түлк", "Шар түлк", "Ээмэг", "Толгойн боолт", "Босоо пальц", "Сорочик", "Бусад"]
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
    cats = ["Шинэ нум", "Хуучин нум", "Амортизатор", "Стермэнь", "Центр боолт", "Дэр", "Зэс түлк", "Пальц", "Хар түлк", "Шар түлк", "Ээмэг", "Толгойн боолт", "Босоо пальц", "Сорочик", "Бусад"]
    return render_template('add_product.html', categories=cats)

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
        flash('Барааны мэдээлэл шинэчлэгдлээ.')
    else:
        initial_stock = float(stock_val) if stock_val else 0
        new_product = Product(sku=sku, name=name, category=category, stock=initial_stock, cost_price=float(cost_price) if cost_price else 0, retail_price=float(retail_price) if retail_price else 0, wholesale_price=float(wholesale_price) if wholesale_price else 0)
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
    t_type = request.form.get('type')
    qty_str = request.form.get('quantity')
    if not p_id or not qty_str:
        flash("Мэдээлэл дутуу байна!")
        return redirect(request.referrer or url_for('dashboard'))
    qty = float(qty_str)
    product = Product.query.get(p_id)
    if product:
        if t_type == 'Орлого' or t_type == 'буцаалт':
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
        db.session.add(Transaction(product_id=p_id, type=f'Тооллого (Зөрүү: {diff})', quantity=new_qty, user_id=current_user.id))
        db.session.commit()
    return redirect(url_for('inventory'))

@app.route('/expenses', methods=['GET', 'POST'])
@login_required
def expenses():
    if request.method == 'POST':
        db.session.add(Expense(category=request.form.get('category'), description=request.form.get('description'), amount=float(request.form.get('amount')), user_id=current_user.id))
        db.session.commit()
        flash('Зардал бүртгэгдлээ')
        return redirect(url_for('expenses'))
    items = Expense.query.order_by(Expense.date.desc()).limit(20).all()
    return render_template('expenses.html', items=items)

@app.route('/statistics')
@login_required
def statistics():
    # Статистик тооцоолох логик (товчилсон)
    return render_template('statistics.html', start_date=(datetime.now()-timedelta(days=30)).strftime('%Y-%m-%d'), end_date=datetime.now().strftime('%Y-%m-%d'), dates=[], sales=[], profit=[], expenses=[], returns=[0], top_labels=[], top_values=[], products=Product.query.all())

@app.route('/users')
@login_required
def users_list():
    if current_user.role != 'admin': return "Эрх хүрэхгүй", 403
    return render_template('users.html', users=User.query.all())

@app.route('/export-balance')
@login_required
def export_balance():
    products = Product.query.filter_by(is_active=True).all()
    df = pd.DataFrame([{'Код': p.sku, 'Бараа': p.name, 'Үлдэгдэл': p.stock} for p in products])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="Uldegdel.xlsx")

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        if not User.query.filter_by(username='admin').first():
            db.session.add(User(username='admin', password=generate_password_hash('123456'), role='admin'))
            db.session.commit()
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
