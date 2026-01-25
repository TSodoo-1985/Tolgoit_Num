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
# Neon-ийн холболтын хаяг
app.config['SQLALCHEMY_DATABASE_URI'] = "postgresql://neondb_owner:npg_J8h1MnAQlbPK@ep-mute-river-a1c92rpd-pooler.ap-southeast-1.aws.neon.tech/neondb?sslmode=require"
app.config['SECRET_KEY'] = 'Sodoo123'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Холболт тасрахаас сэргийлэх (SSL Error-оос сэргийлнэ)
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
    "pool_pre_ping": True,
    "pool_recycle": 300,
}

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

# Өгөгдлийн сан үүсгэх ба Админ тохируулах
with app.app_context():
    db.create_all()
    # Админ хэрэглэгч байгаа эсэхийг шалгаад, байхгүй бол үүсгэх эсвэл нууц үгийг нь '123456' болгох
    admin_user = User.query.filter_by(username='admin').first()
    if not admin_user:
        new_admin = User(
            username='admin', 
            password=generate_password_hash('123456'), 
            role='admin'
        )
        db.session.add(new_admin)
        db.session.commit()
    else:
        # Хэрэв админ байгаа бол нэвтэрч болохгүй байгаа асуудлыг шийдэхээр нууц үгийг нь хүчээр шинэчлэв
        admin_user.password = generate_password_hash('123456')
        db.session.commit()

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
        elif 'зарлага' in t_type or t_type == 'шилжүүлэг':
            product.stock -= qty

        new_tx = Transaction(product_id=p_id, type=t_type, quantity=qty, user_id=current_user.id)
        db.session.add(new_tx)
        db.session.commit()
        flash(f"{product.name} - {t_type} амжилттай бүртгэгдлээ.")

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
        diff_text = f"Зөрүү: {'+' if diff > 0 else ''}{diff}"
        db.session.add(Transaction(product_id=p_id, type=f'Тооллого ({diff_text})', quantity=new_qty, user_id=current_user.id))
        db.session.commit()
    return redirect(url_for('inventory'))

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
    if current_user.role != 'admin':
        flash('Танд хэрэглэгч устгах эрх байхгүй байна!')
        return redirect(url_for('users_list'))

    user_to_delete = User.query.get_or_404(id)
    if user_to_delete.username == 'admin':
        flash('Үндсэн админ admin-г устгах боломжгүй!')
        return redirect(url_for('users_list'))

    db.session.delete(user_to_delete)
    db.session.commit()
    flash(f'Хэрэглэгч {user_to_delete.username} амжилттай устгагдлаа.')
    return redirect(url_for('users_list'))

@app.route('/expenses', methods=['GET', 'POST'])
@login_required
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

    if start_date_str:
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
    else:
        start_date = datetime.now() - timedelta(days=30)

    if end_date_str:
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
    else:
        end_date = datetime.now()

    transactions = Transaction.query.filter(
        Transaction.timestamp >= start_date,
        Transaction.timestamp <= end_date + timedelta(days=1)
    ).all()

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

    for tx in transactions:
        d_str = tx.timestamp.strftime('%m-%d')
        if d_str not in sales_map: continue

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

    expenses_data = Expense.query.filter(
        Expense.date >= start_date,
        Expense.date <= end_date + timedelta(days=1)
    ).all()

    expense_map = {d: 0 for d in dates_list}
    for ex in expenses_data:
        ex_str = ex.date.strftime('%m-%d')
        if ex_str in expense_map:
            expense_map[ex_str] += ex.amount

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
        new_pw = request.form.get('new_password')
        conf_pw = request.form.get('confirm_password')
        if new_pw != conf_pw:
            flash('Нууц үг зөрүүтэй байна!')
            return redirect(url_for('change_password'))
        current_user.password = generate_password_hash(new_pw)
        db.session.commit()
        flash('Нууц үг амжилттай солигдлоо.')
        return redirect(url_for('dashboard'))
    return render_template('change_password.html')

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
        flash(f'"{product.name}" амжилттай засагдлаа.')
        return redirect(url_for('dashboard'))
    return render_template('edit_product.html', product=product)

@app.route('/delete-product/<int:id>', methods=['POST'])
@login_required
def delete_product(id):
    if current_user.role != 'admin':
        flash('Эрх хүрэлцэхгүй байна!')
        return redirect(url_for('dashboard'))
    product = Product.query.get_or_404(id)
    product.is_active = False
    db.session.commit()
    flash(f'"{product.name}" устгагдлаа.')
    return redirect(url_for('dashboard'))

# Excel Export хэсэг
@app.route('/export-balance')
@login_required
def export_balance():
    products = Product.query.filter_by(is_active=True).all()
    data = [{
        'Код (SKU)': p.sku, 'Бараа': p.name, 'Төрөл': p.category, 
        'Үлдэгдэл': p.stock, 'Өртөг': p.cost_price, 
        'Жижиглэн': p.retail_price, 'Бөөний': p.wholesale_price
    } for p in products]
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Үлдэгдэл')
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"Stock_{datetime.now().strftime('%Y%m%d')}.xlsx")

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
  
