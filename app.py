import os
import io
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func, or_
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from urllib.parse import quote
from flask import jsonify

app = Flask(__name__)

# --- ӨГӨГДЛИЙН САНГИЙН ТОХИРГОО (NEON.TECH) ---
# Таны өгсөн Neon холболтын хаягийг энд ашиглаж байна
app.config['SQLALCHEMY_DATABASE_URI'] = "postgresql://neondb_owner:npg_J8h1MnAQlbPK@ep-mute-river-a1c92rpd-pooler.ap-southeast-1.aws.neon.tech/neondb?sslmode=require"
app.config['SECRET_KEY'] = 'Sodoo123'
app.config['SESSION_PERMANENT'] = False
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=15)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {"pool_pre_ping": True, "pool_recycle": 300}
@app.before_request
def make_session_permanent():
    session.permanent = True
    app.permanent_session_lifetime = timedelta(minutes=15)

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
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'), nullable=True)
    type = db.Column(db.String(50), nullable=False)
    quantity = db.Column(db.Float, nullable=False)
    price = db.Column(db.Float, nullable=True, default=0.0)
    description = db.Column(db.Text, nullable=True)
    timestamp = db.Column(db.DateTime, default=datetime.now)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    product = db.relationship('Product', backref='transactions')
    user = db.relationship('User', backref='transactions')

class Expense(db.Model):
    __tablename__ = 'expense'
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(50), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    description = db.Column(db.Text)
    date = db.Column(db.DateTime, default=datetime.utcnow)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    user = db.relationship('User', backref='expense_records')

class LaborFee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    description = db.Column(db.String(200), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    staff_name = db.Column(db.String(100)) # 'staff' биш 'staff_name'
    timestamp = db.Column(db.DateTime, default=datetime.utcnow) # 'date' биш 'timestamp'

# --- БАГЦЫН ЗАГВАР (sku хасагдсан) ---
class Bundle(db.Model):
    __tablename__ = 'bundle'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    set_price = db.Column(db.Float, nullable=False)
    # sku баганыг эндээс хассан
    items = db.relationship('BundleItem', backref='bundle', cascade="all, delete-orphan", lazy=True)

    @property
    def items_json(self):
        return [{
            'product_id': item.product_id,
            'name': item.product.name,
            'quantity': item.quantity
        } for item in self.items]

class BundleItem(db.Model):
    __tablename__ = 'bundle_item'
    id = db.Column(db.Integer, primary_key=True)
    bundle_id = db.Column(db.Integer, db.ForeignKey('bundle.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'), nullable=False)
    quantity = db.Column(db.Float, nullable=False)
    product = db.relationship('Product')

# --- БУСАД ХҮСНЭГТҮҮД ---
class Bow(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_name = db.Column(db.String(200), nullable=False)
    sku = db.Column(db.String(50))
    category = db.Column(db.String(50), default="Хуучин нум")
    purchase_price = db.Column(db.Float, nullable=False)
    retail_price = db.Column(db.Float)
    quantity = db.Column(db.Integer, default=1)
    date = db.Column(db.String(50))
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))

class EmployeeLoan(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    staff_name = db.Column(db.String(100), nullable=False)
    loan_amount = db.Column(db.Float, default=0.0)
    total_paid = db.Column(db.Float, default=0.0)
    description = db.Column(db.String(200))
    date = db.Column(db.DateTime, default=datetime.utcnow)

    @property
    def remaining_balance(self):
        return self.loan_amount - self.total_paid

class ProductLink(db.Model):
    __tablename__ = 'product_link'
    id = db.Column(db.Integer, primary_key=True)
    parent_sku = db.Column(db.String(50), nullable=False, index=True)
    child_sku = db.Column(db.String(50), nullable=False)
    quantity = db.Column(db.Float, default=1.0)

class OldBow(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_name = db.Column(db.String(200), nullable=False)
    sku = db.Column(db.String(100))
    purchase_price = db.Column(db.Float, nullable=False)
    retail_price = db.Column(db.Float, nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    date = db.Column(db.String(100), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    user = db.relationship('User', backref='old_bow_entries')
login_manager = LoginManager(app)
login_manager.login_view = 'login'

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
    # Бараануудыг ID-аар нь дарааллаар нь авах (Шинэ нь хамгийн сүүлд)
    products = Product.query.filter_by(is_active=True).order_by(Product.id.asc()).all()
    cats = ["Шинэ нум", "Хуучин нум", "Амортизатор", "Стермэнь", "Центр боолт", "Дэр", "Зэс түлк", "Пальц", "Хар түлк", "Шар түлк", "Ээмэг", "Толгойн боолт", "Босоо пальц", "Сорочик", "Бусад"]
    return render_template('dashboard.html', products=products, categories=cats)
# --- БАРАА БҮРТГЭЛ, ЗАСВАР ---

@app.route('/add-product-page')
@login_required
def add_product_page():
    cats = ["Шинэ нум", "Хуучин нум", "Амортизатор", "Стермэнь", "Центр боолт", "Дэр", "Зэс түлк", "Пальц", "Хар түлк", "Шар түлк", "Ээмэг", "Толгойн боолт", "Босоо пальц", "Сорочик", "Бусад"]
    return render_template('add_product.html', categories=cats)

@app.route('/add-product', methods=['POST'])
@login_required
def add_product():
    try:
        # 1. Формоос өгөгдөл авах туслах функцүүд
        def get_float(field):
            val = request.form.get(field)
            return float(val) if val and val.strip() else 0.0

        def get_int(field):
            val = request.form.get(field)
            # HTML-ээс stock гэж ирж байгааг int болгоно
            return int(float(val)) if val and val.strip() else 0

        # Текст мэдээллүүд авах
        name = (request.form.get('name') or "").strip()
        original_sku = (request.form.get('sku') or "").strip().upper()
        category = request.form.get('category')

        # --- ЗАССАН ХЭСЭГ: HTML дэх name-тэй тааруулав ---
        cost_price = get_float('cost_price')      # 'purchase_price' байсныг 'cost_price' болгов
        retail_price = get_float('retail_price')
        wholesale_price = get_float('wholesale_price')
        quantity = get_int('stock')               # 'quantity' байсныг 'stock' болгов
        # ----------------------------------------------

        if not name or not original_sku:
            flash("Барааны нэр болон кодыг заавал бөглөнө үү!")
            return redirect(url_for('add_product_page'))

        # 2. Өгөгдлийн санд ижил бараа байгаа эсэхийг шалгах
        existing_product = Product.query.filter(
            func.lower(Product.sku) == original_sku.lower(),
            func.lower(Product.name) == name.lower(),
            Product.cost_price == cost_price,
            Product.retail_price == retail_price
        ).first()

        if existing_product:
            existing_product.stock += quantity
            db.session.commit()
            flash(f"'{name}' барааны үлдэгдэл {quantity}ш-ээр нэмэгдлээ.")
        else:
            # 3. Шинэ SKU код үүсгэх
            new_sku = original_sku
            counter = 1
            while Product.query.filter(func.lower(Product.sku) == new_sku.lower()).first():
                new_sku = f"{original_sku}-{counter}"
                counter += 1

            # 4. Шинэ бараа үүсгэж хадгалах
            new_p = Product(
                name=name,
                sku=new_sku,
                cost_price=cost_price,
                retail_price=retail_price,
                wholesale_price=wholesale_price,
                stock=quantity,
                category=category
            )
            db.session.add(new_p)
            db.session.commit()
            flash(f"Шинэ бараа амжилттай бүртгэгдлээ. (Код: {new_sku})")

    except Exception as e:
        db.session.rollback()
        print(f"ADD PRODUCT ERROR: {str(e)}")
        flash(f"Алдаа гарлаа: {str(e)}")

    return redirect(url_for('add_product_page'))

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

@app.route('/delete-product/<int:id>', methods=['GET'])
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

from flask import jsonify # Файлын дээр заавал нэмээрэй

@app.route('/add_transaction', methods=['POST'])
@login_required
def add_transaction():
    if request.is_json:
        data = request.get_json()
        if data and 'items' in data:
            try:
                for item in data['items']:
                    p_id = item.get('product_id')
                    t_type = item.get('type')
                    qty = float(item.get('quantity') or 0)
                    # 1. Сагснаас ирж буй зассан үнийг авах
                    custom_price = item.get('price')

                    product = Product.query.get(p_id)
                    if product:
                        # 2. Хэрэв сагснаас үнэ ирээгүй бол барааны үндсэн үнийг сонгох
                        if custom_price is None or float(custom_price) == 0:
                            if "Бөөний" in t_type:
                                actual_price = product.wholesale_price
                            elif "Жижиглэн" in t_type:
                                actual_price = product.retail_price
                            else:
                                actual_price = 0
                        else:
                            actual_price = float(custom_price)

                        # Сток засах
                        if t_type in ['Орлого', 'буцаалт']:
                            product.stock += qty
                        else:
                            product.stock -= qty

                        # 3. TRANSACTION-Д ҮНЭГ ХАМТ ХАДГАЛАХ
                        db.session.add(Transaction(
                            product_id=p_id,
                            type=t_type,
                            quantity=qty,
                            price=actual_price,  # ЭНЭ МӨРИЙГ НЭМЛЭЭ
                            user_id=current_user.id
                        ))

                db.session.commit()
                return jsonify({"success": True, "message": "Бүх гүйлгээ амжилттай хадгалагдлаа."})
            except Exception as e:
                db.session.rollback()
                return jsonify({"success": False, "message": str(e)}), 500

    # Ганц бараа Form-оор ирэх үеийн логик (Үнийг мөн нэмэв)
    p_id = request.form.get('product_id')
    t_type = request.form.get('type')
    qty = float(request.form.get('quantity') or 0)
    product = Product.query.get(p_id)

    if product:
        # Үнэ тодорхойлох
        if "Бөөний" in t_type:
            actual_price = product.wholesale_price
        elif "Жижиглэн" in t_type:
            actual_price = product.retail_price
        else:
            actual_price = 0

        if t_type in ['Орлого', 'буцаалт']:
            product.stock += qty
        else:
            product.stock -= qty

        # Үнээр нь хадгалах
        db.session.add(Transaction(
            product_id=p_id,
            type=t_type,
            quantity=qty,
            price=actual_price, # ЭНЭ МӨРИЙГ НЭМЛЭЭ
            user_id=current_user.id
        ))
        db.session.commit()
        flash(f"{product.name} - {t_type} бүртгэгдлээ.")

    return redirect(request.referrer or url_for('dashboard'))

@app.route('/inventory')
@login_required
def inventory():
    # 1. Бүх барааг авах
    products = Product.query.all()

    # 2. Түүх харуулах хэсэг (Алдаанаас сэргийлж түр хоосон жагсаалт болгов)
    # Хэрэв таны түүх хадгалдаг модель 'Transaction' бол Transaction.query... гэж бичнэ
    history = []

    # 3. Багц тохиргоонд орсон бараануудыг шүүж авах
    try:
        package_skus = [link.parent_sku for link in ProductLink.query.with_entities(ProductLink.parent_sku).distinct().all()]
        package_products = [p for p in products if p.sku in package_skus]
    except:
        package_products = [] # Хэрэв ProductLink хүснэгт байхгүй бол алдаа заахгүй

    return render_template('inventory.html',
                           products=products,
                           package_products=package_products,
                           history=history)

@app.route('/do_inventory', methods=['POST'])
@login_required
def do_inventory():
    product_id = request.form.get('product_id')
    new_quantity = request.form.get('quantity')

    if not product_id or not new_quantity:
        flash('Бараа болон тоо хэмжээг бүрэн бөглөнө үү!')
        return redirect(url_for('inventory'))

    product = Product.query.get_or_404(product_id)
    old_stock = product.stock or 0
    new_stock = float(new_quantity)
    diff = new_stock - old_stock

    # 1. Барааны үлдэгдлийг шинэчлэх
    product.stock = new_stock

    # 2. Гүйлгээний түүхэнд бүртгэх
    # Таны Transaction моделд 'price' болон 'note' байхгүй тул тэдгээрийг хаслаа.
    # Хэрэв зөрүүг харахыг хүсвэл 'type' баганад нь тайлбар болгож хадгалж болно.
    transaction = Transaction(
        product_id=product.id,
        quantity=new_stock, # Тоолсон бодит тоо
        type=f"Тооллого (Зөрүү: {'+' if diff >= 0 else ''}{diff})",
        timestamp=datetime.now(),
        user_id=current_user.id
    )

    db.session.add(transaction)
    db.session.commit()

    flash(f'{product.name} амжилттай тоологдлоо. Шинэ үлдэгдэл: {new_stock}')
    return redirect(url_for('inventory'))

@app.route('/expenses', methods=['GET', 'POST'])
@login_required
def expenses():
    if request.method == 'POST':
        category = request.form.get('category')
        description = request.form.get('description')
        amount = float(request.form.get('amount'))

        if category == 'Ажлын хөлс':
            # Үйлчлүүлэгчээс орж ирж буй, ажилчны нэр дээр хуримтлагдах мөнгө
            new_item = LaborFee(description=description, amount=amount, staff_name=current_user.username)
        else:
            # Цалин олголт эсвэл Ерөнхий зардал (Кассаас гарч буй мөнгө)
            new_item = Expense(category=category, description=description, amount=amount)

        db.session.add(new_item)
        db.session.commit()
        return redirect(url_for('expenses'))

    # Жагсаалтыг нэгтгэж харуулах (Өмнөх Date-ийн алдааг зассан хувилбар)
    expenses_list = Expense.query.all()
    labor_list = LaborFee.query.all()
    items = []

    for e in expenses_list:
        items.append({
            'date': e.date, # Expense хүснэгтийн багана
            'category': e.category,
            'description': e.description,
            'amount': e.amount,
            'staff': 'Систем'
        })

    for l in labor_list:
        items.append({
            'date': l.timestamp, # LaborFee хүснэгтийн багана
            'category': 'Ажлын хөлс',
            'description': l.description,
            'amount': l.amount,
            'staff': l.staff_name
        })

    items.sort(key=lambda x: x['date'], reverse=True)
    return render_template('expenses.html', items=items[:20])

    # 1. Expense хүснэгт 'date' баганатай тул e.date гэж авна
    for e in expenses_list:
        items.append({
            'date': e.date,
            'category': e.category,
            'description': e.description,
            'amount': e.amount,
            'staff': 'Систем'
        })

    # 2. LaborFee хүснэгт 'timestamp' баганатай тул l.timestamp гэж авна
    for l in labor_list:
        items.append({
            'date': l.timestamp,
            'category': 'Ажлын хөлс',
            'description': l.description,
            'amount': l.amount,
            'staff': l.staff_name
        })

    # Огноогоор нь сүүлийнхээс нь эхэлж жагсаах
    items.sort(key=lambda x: x['date'], reverse=True)
    items = items[:20]

    return render_template('expenses.html', items=items)

@app.route('/cart')
@login_required
def cart_page():
    return render_template('cart.html')

@app.route('/add_transaction_bulk', methods=['POST'])
@login_required
def add_transaction_bulk():
    data = request.get_json()
    if not data or 'items' not in data:
        return jsonify({'status': 'error', 'message': 'Өгөгдөл хоосон байна'}), 400

    import random
    now = datetime.now()
    batch_id = now.strftime('%Y%m%d%H%M') + "-" + str(random.randint(100, 999))

    try:
        for item in data['items']:
            is_bundle = item.get('is_bundle', False)
            is_labor = item.get('is_labor', False)
            raw_name = item.get('name', 'Мэдэгдэхгүй')
            
            # Төлбөрийн суваг (Бэлэн, Картаар г.м)
            raw_type = item.get('type') if item.get('type') else "Жижиглэн"
            price_val = float(item.get('price', 0))
            qty_val = float(item.get('quantity', 1))

            # --- 1. АЖЛЫН ХӨЛС ---
            if is_labor:
                # Төрлийг "Жижиглэн (Төлбөрийн хэлбэр)" болгож форматлах
                payment_type = f"Жижиглэн ({raw_type})" if "(" not in raw_type else raw_type
                clean_name = raw_name.replace("[АЖЛЫН ХӨЛС] ", "").strip()
                
                new_labor = LaborFee(
                    description=f"[{batch_id}] {clean_name} ({payment_type})",
                    amount=price_val,
                    staff_name=current_user.username,
                    timestamp=now
                )
                db.session.add(new_labor)

                new_tx = Transaction(
                    product_id=None,
                    description=f"[{batch_id}] [АЖЛЫН ХӨЛС] {clean_name}",
                    quantity=1,
                    price=price_val,
                    type=payment_type,
                    timestamp=now,
                    user_id=current_user.id
                )
                db.session.add(new_tx)

            # --- 2. БАГЦ БАРАА (BUNDLE) ---
            elif is_bundle:
                # Багц барааг мөн "Жижиглэн" ангилалд оруулна
                payment_type = f"Жижиглэн ({raw_type})" if "(" not in raw_type else raw_type

                new_tx = Transaction(
                    product_id=None,
                    description=f"[{batch_id}] {raw_name}",
                    quantity=qty_val,
                    price=price_val,
                    type=payment_type,
                    timestamp=now,
                    user_id=current_user.id
                )
                db.session.add(new_tx)

                bundle_items = item.get('bundle_items', [])
                for b_item in bundle_items:
                    p_id = b_item.get('product_id')
                    if p_id:
                        p = Product.query.get(int(p_id))
                        if p:
                            p.stock -= float(b_item.get('quantity', 0)) * qty_val

            # --- 3. БҮРТГЭЛГҮЙ БАРАА БОЛОН ЭНГИЙН БАРАА ---
            else:
                p_id = item.get('product_id')
                tx_desc = f"[{batch_id}] {raw_name}"
                target_p_id = None
                payment_type = raw_type

                if p_id and str(p_id).isdigit():
                    product = Product.query.get(int(p_id))
                    if product:
                        product.stock -= qty_val
                        target_p_id = product.id
                        tx_desc = f"[{batch_id}] {product.name}"
                
                new_tx = Transaction(
                    product_id=target_p_id,
                    description=tx_desc,
                    quantity=qty_val,
                    price=price_val,
                    type=payment_type,
                    timestamp=now,
                    user_id=current_user.id
                )
                db.session.add(new_tx)

        db.session.commit()
        return jsonify({'status': 'success', 'message': 'Амжилттай бүртгэгдлээ'})

    except Exception as e:
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500
 
@app.route('/special_transfer', methods=['GET', 'POST'])
@login_required
def special_transfer():
    # Viewer эрхтэй хүн шилжүүлэг хийж болохгүй
    if current_user.role == 'viewer':
        flash('Танд шилжүүлэг хийх эрх байхгүй байна.', 'danger')
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        # Формоос ирсэн жагсаалтуудыг авах
        product_ids = request.form.getlist('product_ids[]')
        quantities = request.form.getlist('quantities[]')
        note = request.form.get('note')

        if not product_ids:
            flash('Шилжүүлэх бараа сонгоогүй байна!', 'warning')
            return redirect(url_for('special_transfer'))

        try:
            # Бүх барааг нэг дор боловсруулах
            for p_id, qty in zip(product_ids, quantities):
                product = Product.query.get(p_id)
                q = float(qty) if qty else 0

                if product and q > 0:
                    # Сток хасах
                    product.stock -= q

                    # Гүйлгээний түүхөнд 'Өртгөөр зарлага' гэж тэмдэглэх
                    new_tx = Transaction(
                        product_id=product.id,
                        user_id=current_user.id,
                        type='Өртгөөр зарлага',
                        quantity=q,
                        # Шилжүүлэгт борлуулах үнэ биш өртөг үнийг нь авна
                        price=product.cost_price if product.cost_price else 0,
                        description=f"Шилжүүлэг: {note}",
                        timestamp=datetime.now()
                    )
                    db.session.add(new_tx)

            db.session.commit()
            flash(f'Нийт {len(product_ids)} төрлийн барааг "{note}" тайлбартайгаар амжилттай шилжүүллээ.', 'success')
            return redirect(url_for('dashboard'))

        except Exception as e:
            db.session.rollback()
            flash(f'Алдаа гарлаа: {str(e)}', 'danger')
            return redirect(url_for('special_transfer'))

    # Зөвхөн идэвхтэй байгаа бараануудыг харуулах
    products = Product.query.filter(Product.is_active == True).order_by(Product.name).all()
    return render_template('special_transfer.html', products=products)

@app.route('/labor')
@login_required
def labor_page():
    # Сүүлийн 100 бичилтийг харуулна
    fees = LaborFee.query.order_by(LaborFee.timestamp.desc()).all()
    total_labor = sum(f.amount for f in fees)
    return render_template('labor.html', fees=fees, total_labor=total_labor)

@app.route('/add_labor', methods=['POST'])
@login_required
def add_labor():
    desc = request.form.get('description')
    amt = request.form.get('amount')
    staff = request.form.get('staff_name')

    if desc and amt:
        new_fee = LaborFee(
            description=desc,
            amount=float(amt),
            staff_name=staff or current_user.username
        )
        db.session.add(new_fee)
        db.session.commit()
        flash("Ажлын хөлс амжилттай бүртгэгдлээ.")
    return redirect(url_for('labor_page'))

@app.route('/delete_labor/<int:id>')
@login_required
def delete_labor(id):
    if current_user.role != 'admin':
        return "Эрх хүрэхгүй", 403
    fee = LaborFee.query.get_or_404(id)
    db.session.delete(fee)
    db.session.commit()
    return redirect(url_for('labor_page'))

# 1. Excel бэлдэц татах (Монгол толгойтой)
@app.route('/download_template')
@login_required
def download_template():
    if current_user.role != 'admin':
        return "Хандах эрхгүй", 403

    # Багануудын нэр
    columns = ['Код (SKU)', 'Барааны нэр', 'Ангилал', 'Өртөг', 'Бөөний үнэ', 'Жижиглэн үнэ', 'Үлдэгдэл']
    df = pd.DataFrame(columns=columns)

    output = BytesIO()
    # Энд engine='openpyxl' гэж зааж өгөхөд дээрх сан заавал хэрэгтэй
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    output.seek(0)

    return send_file(
        output,
        download_name="baraa_tatakh_beldetz.xlsx",
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/import_products_action', methods=['POST'])
@login_required
def import_products_action():
    if current_user.role != 'admin':
        return "Хандах эрхгүй", 403

    file = request.files.get('file')
    if not file:
        flash("Файл сонгоно уу!")
        return redirect(url_for('import_products_page'))

    try:
        # Excel файлыг унших
        df = pd.read_excel(file, engine='openpyxl')

        # --- Хоосон (NaN) утгуудыг 0 болгож, бүхэл тоо руу хөрвүүлэх бэлтгэл ---
        # pd.to_numeric ашиглаад fillna(0) хийж, дараа нь .astype(int) болгоно
        cols_to_fix = ['Өртөг', 'Бөөний үнэ', 'Жижиглэн үнэ', 'Үлдэгдэл']
        for col in cols_to_fix:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

        df['Ангилал'] = df['Ангилал'].fillna("Бусад")
        df['Барааны нэр'] = df['Барааны нэр'].fillna("Нэргүй")

        count = 0
        # iterrows() ашиглан Excel-ийн мөрийн дарааллыг ягштал баримтална
        for index, row in df.iterrows():
            sku_val = row['Код (SKU)']
            if pd.isna(sku_val):
                continue

            # SKU-г текст хэлбэрт оруулах (бутархай .0-ийг арилгах)
            sku_str = str(int(sku_val)) if isinstance(sku_val, (int, float)) else str(sku_val)

            product = Product.query.filter_by(sku=sku_str).first()

            if product:
                # Бараа байвал мэдээллийг шинэчлэх (int ашиглаж 0 болгоно)
                product.name = str(row['Барааны нэр'])
                product.category = str(row['Ангилал'])
                product.cost_price = int(row['Өртөг'])
                product.wholesale_price = int(row['Бөөний үнэ'])
                product.retail_price = int(row['Жижиглэн үнэ'])
                product.stock = int(row['Үлдэгдэл'])
            else:
                # Шинэ бараа үүсгэх
                new_p = Product(
                    sku=sku_str,
                    name=str(row['Барааны нэр']),
                    category=str(row['Ангилал']),
                    cost_price=int(row['Өртөг']),
                    wholesale_price=int(row['Бөөний үнэ']),
                    retail_price=int(row['Жижиглэн үнэ']),
                    stock=int(row['Үлдэгдэл'])
                )
                db.session.add(new_p)

            # Мэдээллийн санд ID-г дарааллаар олгохын тулд flush хийнэ
            db.session.flush()
            count += 1

        db.session.commit()
        flash(f"Амжилттай! Нийт {count} бараа Excel-ийн дарааллаар бүртгэгдлээ.")
    except Exception as e:
        db.session.rollback()
        flash(f"Алдаа гарлаа: {str(e)}")

    return redirect(url_for('import_products_page'))

# Энэ функц нь import_products.html хуудсыг нээж харуулна
@app.route('/import_products_page')
@login_required
def import_products_page():
    if current_user.role != 'admin':
        flash("Уучлаарай, зөвхөн Админ нэвтрэх боломжтой!")
        return redirect(url_for('dashboard'))
    return render_template('import_products.html')

# ... (бусад import-ууд хэвээрээ)

@app.route('/returns', methods=['GET', 'POST'])
@login_required
def returns():
    if request.method == 'POST':
        try:
            product_id = request.form.get('product_id')
            quantity = float(request.form.get('quantity') or 0)
            price = float(request.form.get('price') or 0)

            product = Product.query.get_or_404(product_id)

            # 1. Үлдэгдэл нэмэх
            product.stock += quantity

            # 2. Буцаалтын дүнг тооцох
            total_amount = quantity * price

            # 3. Гүйлгээг бүртгэх
            # ТАЙЛБАР: Хэрэв таны Transaction модель 'amount' биш 'price' талбартай бол
            # доорх price=-price хэсгийг анхаараарай.
            new_transaction = Transaction(
                product_id=product.id,
                quantity=quantity,
                type='Буцаалт',
                price=-price,  # Нэгж үнийг хасахаар хадгалснаар нийт орлогоос хасагдана
                user_id=current_user.id,
                timestamp=datetime.now(), # 'date' эсвэл 'timestamp' таны моделийн нэрээр
                description=f"Буцаалт: {product.name}"
            )

            # 4. Касснаас мөнгө хасагдахын тулд Expense (Зардал) болгож бүртгэх
            # Хэрэв таны dashboard 'Нийт Орлого - Нийт Зардал' гэж тооцдог бол энэ хэрэгтэй.
            new_expense = Expense(
                category="Бараа буцаалт",
                amount=total_amount,
                description=f"Буцаалт: {product.name} ({quantity}ш)",
                date=datetime.now(),
                user_id=current_user.id
            )

            db.session.add(new_transaction)
            db.session.add(new_expense)
            db.session.commit()

            flash(f"'{product.name}' барааны буцаалт амжилттай. Касснаас {total_amount:,.0f}₮ хасагдлаа.")
            return redirect(url_for('dashboard'))

        except Exception as e:
            db.session.rollback()
            flash(f"Алдаа гарлаа: {str(e)}")
            return redirect(url_for('returns'))

    products = Product.query.filter_by(is_active=True).all()
    return render_template('returns.html', products=products)

@app.route('/buy-old-bow', methods=['GET', 'POST'])
@login_required
def buy_old_bow():
    if request.method == 'POST':
        selected_product_id = request.form.get('product_id')
        manual_name = request.form.get('name')
        sku = request.form.get('sku')
        cost_price = float(request.form.get('cost_price', 0))
        retail_price = float(request.form.get('retail_price', 0))
        qty_to_add = float(request.form.get('stock', 1))

        # 1. SKU-г тодорхойлох
        if selected_product_id and selected_product_id != 'other':
            product_obj = db.session.get(Product, selected_product_id)
            final_sku = sku if sku else product_obj.sku
            final_name = product_obj.name
        else:
            final_sku = sku if sku else f"OLD-{uuid.uuid4().hex[:6].upper()}"
            final_name = f"[ХУУЧИН] {manual_name}"

        # 2. Энэ SKU-тай бараа агуулахад байгаа эсэхийг шалгах
        existing_product = Product.query.filter_by(sku=final_sku).first()

        if existing_product:
            # БАЙГАА БОЛ: Үлдэгдэл дээр нэмж, үнийг шинэчилнэ
            existing_product.stock += qty_to_add
            existing_product.cost_price = cost_price  # Сүүлийн авсан өртөгөөр шинэчлэх
            existing_product.retail_price = retail_price  # Зарах үнийг шинэчлэх
            target_product = existing_product
            msg = f"'{final_name}' (SKU: {final_sku}) үлдэгдэл дээр {qty_to_add} ширхэг нэмэгдлээ."
        else:
            # БАЙХГҮЙ БОЛ: Шинээр үүсгэнэ
            new_item = Product(
                name=final_name,
                sku=final_sku,
                category="Хуучин нум",
                cost_price=cost_price,
                retail_price=retail_price,
                stock=qty_to_add,
                is_active=True
            )
            db.session.add(new_item)
            target_product = new_item
            msg = f"'{final_name}' шинээр бүртгэгдэж, агуулахад орлоо."

        # 3. Хуучин нумны түүхэнд (OldBow хүснэгт) бүртгэх
        new_history = OldBow(
            product_name=final_name,
            sku=final_sku,
            purchase_price=cost_price,
            retail_price=retail_price,
            quantity=int(qty_to_add),
            date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            user_id=current_user.id
        )
        db.session.add(new_history)

        # 4. Кассаас зарлага гаргах (Transaction хүснэгт)
        new_tx = Transaction(
            type='Зарлага (Бэлэн)',
            description=f"Хуучин нум авалт: {final_name} (SKU: {final_sku})",
            price=cost_price,
            quantity=qty_to_add,
            timestamp=datetime.now(),
            user_id=current_user.id
        )
        db.session.add(new_tx)
        
        db.session.commit()
        flash(msg)
        return redirect(url_for('buy_old_bow'))

    # Зөвхөн "Хуучин нум" ангилалтай барааг шүүж авах
    old_bow_products = Product.query.filter_by(category="Хуучин нум").all()
    return render_template('buy_old_bow.html', products=old_bow_products)

@app.route('/manage_packages', methods=['GET', 'POST'])
@login_required
def manage_packages():
    # Admin болон Staff биш бол хандах эрхгүй
    if current_user.role not in ['admin', 'staff']:
        flash("Танд хандах эрх байхгүй!")
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        parent_sku = request.form.get('parent_sku').strip().upper()
        # Textarea-аас мөр мөрөөр нь салгаж авах
        raw_child_skus = request.form.get('child_skus', '')
        child_skus = [sku.strip().upper() for sku in raw_child_skus.split('\n') if sku.strip()]

        if parent_sku and child_skus:
            # 1. Хуучин заавар байвал устгах
            ProductLink.query.filter_by(parent_sku=parent_sku).delete()

            # 2. Шинэ заавруудыг нэмэх
            for sku in child_skus:
                new_link = ProductLink(parent_sku=parent_sku, child_sku=sku, quantity=1.0)
                db.session.add(new_link)

            db.session.commit()
            flash(f"{parent_sku} комын заавар амжилттай хадгалагдлаа.")
        return redirect(url_for('manage_packages'))

    # Жагсаалтыг харуулах (Бүлэглэх)
    all_links = ProductLink.query.all()
    packages = {}
    for link in all_links:
        if link.parent_sku not in packages:
            packages[link.parent_sku] = []
        packages[link.parent_sku].append(link.child_sku)

    return render_template('manage_packages.html', packages=packages)

# УСТГАХ ФУНКЦ (ШИНЭ)
@app.route('/delete_package_instruction/<parent_sku>', methods=['POST'])
@login_required
def delete_package_instruction(parent_sku):
    if current_user.role in ['admin', 'staff']:
        ProductLink.query.filter_by(parent_sku=parent_sku).delete()
        db.session.commit()
        flash(f"{parent_sku} багцын заавар устлаа.")
    else:
        flash("Танд устгах эрх байхгүй!")
    return redirect(url_for('manage_packages'))

@app.route('/disassemble_simple', methods=['POST'])
@login_required
def disassemble_simple():
    if current_user.role not in ['admin', 'staff']:
        flash("Танд энэ үйлдлийг хийх эрх байхгүй.")
        return redirect(url_for('inventory'))

    # ЭНЭ МӨР ЧУХАЛ: Формоос ирж буй product_id-г авч байна
    product_id = request.form.get('product_id')

    if not product_id:
        flash("Бараа сонгогдоогүй байна.")
        return redirect(url_for('inventory'))

    # Одоо product_id тодорхойлогдсон тул алдаа заахгүй
    main_product = Product.query.get_or_404(product_id)

    links = ProductLink.query.filter_by(parent_sku=main_product.sku).all()

    if not links:
        flash(f"'{main_product.sku}' кодонд задлах заавар байхгүй.")
        return redirect(url_for('inventory'))

    if main_product.stock < 1:
        flash(f"'{main_product.sku}' үлдэгдэл хүрэлцээгүй.")
        return redirect(url_for('inventory'))

    try:
        main_product.stock -= 1
        for link in links:
            child = Product.query.filter_by(sku=link.child_sku).first()
            if child:
                child.stock += link.quantity

        db.session.commit()
        flash(f"{main_product.sku} амжилттай задарлаа.")
    except Exception as e:
        db.session.rollback()
        flash(f"Алдаа: {str(e)}")

    return redirect(url_for('inventory'))

# --- 2. ХУУЧИН БАРААНЫ ТАЙЛАН ---
@app.route('/old-bow-report')
@login_required
def old_bow_report():
    # Хамгийн сүүлд авснаараа эрэмбэлэгдэнэ
    reports = OldBow.query.order_by(OldBow.id.desc()).all()
    return render_template('old_bow_report.html', reports=reports)

# --- CHECKOUT ХЭСЭГТ VIEWER ЭРХИЙГ ХЯЗГААРЛАХ ---
@app.route('/checkout', methods=['POST'])
@login_required
def checkout():
    if current_user.role == 'viewer':
        return jsonify({'success': False, 'message': 'Танд зарлага гаргах эрх байхгүй!'})

@app.route('/loans', methods=['GET', 'POST'])
@login_required
def manage_loans():
    if request.method == 'POST':
        staff_name = request.form.get('staff_name')
        amount = float(request.form.get('amount'))
        desc = request.form.get('description')

        # 1. Зээлийн бүртгэл үүсгэх
        new_loan = EmployeeLoan(staff_name=staff_name, loan_amount=amount, description=desc)

        # 2. Кассаас зардал болгож хасах
        new_expense = Expense(
            category="Ажилчны зээл",
            amount=amount,
            description=f"{staff_name}-д зээл олгох: {desc}",
            user_id=current_user.id
        )

        db.session.add(new_loan)
        db.session.add(new_expense)
        db.session.commit()
        flash("Зээл амжилттай бүртгэгдэж, кассаас хасагдлаа.")
        return redirect(url_for('manage_loans'))

    loans = EmployeeLoan.query.all()
    return render_template('loans.html', loans=loans)

@app.route('/pay-loan/<int:id>', methods=['POST'])
@login_required
def pay_loan(id):
    loan = EmployeeLoan.query.get_or_404(id)
    pay_amount = float(request.form.get('pay_amount'))

    if pay_amount > 0:
        loan.total_paid += pay_amount
        # Кассанд орлого болгож нэмэх (сонголтоор)
        # Энэ нь цалингаас суутгаж байгаа бол цалингийн зардлыг багасгах байдлаар бүртгэж болно.
        db.session.commit()
        flash(f"{loan.staff_name}-ийн төлөлт бүртгэгдлээ.")

    return redirect(url_for('manage_loans'))

@app.route('/complete_sale', methods=['POST'])
@login_required
def complete_sale():
    product_ids = request.form.getlist('product_ids[]')
    quantities = request.form.getlist('quantities[]')
    sale_prices = request.form.getlist('sale_prices[]') # Сагсанд зассан үнүүд

    for i in range(len(product_ids)):
        p_id = product_ids[i]
        qty = float(quantities[i])
        actual_sale_price = float(sale_prices[i]) # Энэ бол нэмж бичсэн үнэ

        product = Product.query.get(p_id)

        # 1. Үндсэн барааны үлдэгдлийг хасна (ҮНЭ-г өөрчлөхгүй!)
        product.stock -= qty

        # 2. Гүйлгээний түүхэнд "Нэмсэн үнэ"-ээр нь хадгална
        new_log = Transaction(
            product_id=p_id,
            type='борлуулалт',
            quantity=qty,
            price=actual_sale_price, # <-- Энд нэмсэн үнэ хадгалагдана
            user_id=current_user.id,
            date=datetime.now()
        )
        db.session.add(new_log)

    db.session.commit()
    return redirect(url_for('inventory'))

@app.route('/create_bundle', methods=['POST'])
@login_required
def create_bundle():
    data = request.get_json()
    if not data:
        return jsonify({'success': False, 'message': 'Мэдээлэл хоосон байна'})

    try:
        # 1. Шинэ багц үүсгэх (sku-г хасав)
        new_bundle = Bundle(
            name=data.get('name'),
            set_price=float(data.get('set_price', 0))
        )
        db.session.add(new_bundle)
        db.session.flush()

        # 2. Багцын орц бүрийг хадгалах
        for item in data.get('items', []):
            new_item = BundleItem(
                bundle_id=new_bundle.id,
                product_id=item.get('id'),
                quantity=float(item.get('quantity', 1))
            )
            db.session.add(new_item)

        db.session.commit()
        return jsonify({'success': True})

    except Exception as e:
        db.session.rollback()
        print(f"Bundle Error: {str(e)}")
        return jsonify({'success': False, 'message': "Баазтай холбогдоход алдаа гарлаа: " + str(e)})

@app.route('/bundles')
@login_required
def bundles_page():
    # Зүүн талын жагсаалтад зориулж бараануудыг авах
    products = Product.query.all()
    # Багцуудыг баазаас татаж авах (sku хайхгүй)
    bundles = Bundle.query.order_by(Bundle.id.desc()).all()
    return render_template('bundles.html', products=products, bundles=bundles)

@app.route('/delete_bundle/<int:id>', methods=['POST'])
@login_required
def delete_bundle(id):
    bundle = Bundle.query.get_or_404(id)
    try:
        db.session.delete(bundle)
        db.session.commit()
        return jsonify({'success': True})
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': str(e)})

# --- САЛБАРЫН ОРЛОГО (BATCH ENTRY) ---

@app.route('/internal-incomes')
@login_required
def internal_income_list():
    # Зөвхөн 'Орлого' төрөлтэй гүйлгээнүүдийг шүүж авна
    page = request.args.get('page', 1, type=int)
    incomes = Transaction.query.filter_by(type='Орлого').order_by(Transaction.id.desc()).paginate(page=page, per_page=20)
    return render_template('internal_income.html', incomes=incomes)

@app.route('/add-internal-income', methods=['GET', 'POST'])
@login_required
def add_internal_income():
    if request.method == 'POST':
        product_ids = request.form.getlist('p_ids[]')
        quantities = request.form.getlist('qtys[]')
        description = request.form.get('description', 'Салбарын орлого')

        try:
            for p_id, qty in zip(product_ids, quantities):
                if not qty or int(qty) <= 0:
                    continue

                product = Product.query.get(p_id)
                if product:
                    # 1. Үлдэгдэл нэмэх
                    product.stock += int(qty)

                    # 2. Гүйлгээг 'Орлого' төрлөөр бүртгэх (Кассаас хасагдахгүй)
                    new_trans = Transaction(
                        product_id=product.id,
                        user_id=current_user.id,
                        quantity=int(qty),
                        price=product.cost_price, # Өртгөөр нь бүртгэнэ
                        type='Орлого',
                        description=description,
                        timestamp=datetime.now()
                    )
                    db.session.add(new_trans)

            db.session.commit()
            flash("Салбарын орлого амжилттай бүртгэгдлээ.")
            return redirect(url_for('internal_income_list'))

        except Exception as e:
            db.session.rollback()
            flash(f"Алдаа гарлаа: {str(e)}")
            return redirect(url_for('add_internal_income'))

    # Сонгох барааны жагсаалт
    products = Product.query.filter_by(is_active=True).order_by(Product.name).all()
    return render_template('add_internal_income.html', products=products)

# --- ТАЙЛАН, СТАТИСТИК ---
@app.route('/statistics')
@login_required
def statistics():
    # 1. Огноо болон Борлуулалтын төрлийн шүүлтүүр авах
    start_date_str = request.args.get('start_date')
    end_date_str = request.args.get('end_date')
    sale_type = request.args.get('sale_type', 'Бүгд')

    # Огнооны логик
    if start_date_str and end_date_str:
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
        delta = end_date - start_date
        dates = [(start_date + timedelta(days=i)).strftime('%Y-%m-%d') for i in range(delta.days + 1)]
    else:
        today = datetime.now().date()
        dates = [(today - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(6, -1, -1)]

    sales_data = []
    profit_data = []
    expense_data = []

    # 2. Графикийн өгөгдөл боловсруулах (Өдөр бүрээр)
    for date_str in dates:
        query = Transaction.query.filter(
            Transaction.type.like('%зарлага%'),
            db.func.date(Transaction.timestamp) == date_str
        )

        if sale_type == 'Бөөний':
            query = query.filter(Transaction.type == 'Бөөний зарлага')
        elif sale_type == 'Жижиглэн':
            query = query.filter(Transaction.type == 'Жижиглэн зарлага')

        day_transactions = query.all()

        daily_sales = 0
        daily_cost = 0

        for t in day_transactions:
            # --- БОРЛУУЛАЛТЫН ҮНЭ ТОГТООХ ---
            sell_price = t.price if (t.price and t.price > 0) else 0
            if sell_price == 0:
                if t.bundle:
                    sell_price = t.bundle.price
                elif t.product:
                    sell_price = t.product.wholesale_price if "Бөөний" in t.type else t.product.retail_price

            # --- ӨРТӨГ (COST) ТООЦООЛОХ (БАГЦЫГ ОРОЛЦУУЛАХ) ---
            unit_cost = 0
            if t.bundle:
                # Багцыг бүрдүүлэгч бараа бүрийн (өртөг * тоо ширхэг)-ийг нэмж нийт өртгийг гаргана
                for item in t.bundle.items:
                    if item.product:
                        unit_cost += (item.product.cost_price or 0) * item.quantity
            elif t.product:
                # Энгийн бараа бол өөрийнх нь өртгийг авна
                unit_cost = t.product.cost_price or 0

            daily_sales += sell_price * t.quantity
            daily_cost += unit_cost * t.quantity

        # Зардал (Зээл олголтыг статистикт оруулахгүй байх шүүлтүүртэй)
        daily_expense = db.session.query(db.func.sum(Expense.amount)).\
            filter(
                Expense.category != "Ажлын хөлс",
                Expense.category != "Loan_Disbursement", # Зээл олголтыг хасах
                db.func.date(Expense.date) == date_str
            ).scalar() or 0

        sales_data.append(float(daily_sales))
        expense_data.append(float(daily_expense))
        # Цэвэр ашиг = Борлуулалт - Өртөг - Зардал
        profit_data.append(float(daily_sales - daily_cost - daily_expense))

    # 3. Нийт ТОП 5 бараа
    top_query = db.session.query(Product.name, db.func.sum(Transaction.quantity)).\
        join(Transaction).filter(Transaction.type.like('%зарлага%'))

    if sale_type == 'Бөөний':
        top_query = top_query.filter(Transaction.type == 'Бөөний зарлага')
    elif sale_type == 'Жижиглэн':
        top_query = top_query.filter(Transaction.type == 'Жижиглэн зарлага')

    top_products_all = top_query.group_by(Product.name).order_by(db.func.sum(Transaction.quantity).desc()).limit(5).all()

    top_labels = [p[0] for p in top_products_all]
    top_values = [int(p[1]) for p in top_products_all]

    # 4. Ангилал бүрийн ТОП 5
    categories = db.session.query(Product.category).distinct().all()
    stats_data = {}

    for cat in categories:
        category_name = cat[0] if cat[0] else "Ангилалгүй"
        cat_top_query = db.session.query(
            Product.name,
            db.func.sum(Transaction.quantity).label('total')
        ).join(Transaction).filter(
            Product.category == category_name,
            Transaction.type.like('%зарлага%')
        )

        if sale_type == 'Бөөний':
            cat_top_query = cat_top_query.filter(Transaction.type == 'Бөөний зарлага')
        elif sale_type == 'Жижиглэн':
            cat_top_query = cat_top_query.filter(Transaction.type == 'Жижиглэн зарлага')

        category_top = cat_top_query.group_by(Product.name).order_by(db.desc('total')).limit(5).all()
        if category_top:
            stats_data[category_name] = category_top

    return render_template('statistics.html',
                           dates=dates,
                           sales=sales_data,
                           profit=profit_data,
                           expenses=expense_data,
                           top_labels=top_labels,
                           top_values=top_values,
                           stats_data=stats_data,
                           sale_type=sale_type,
                           start_date=start_date_str or "",
                           end_date=end_date_str or "")
@app.route('/export-loans')
@login_required
def export_loans():
    # 1. Өгөгдлийн сангаас бүх зээлийн мэдээллийг авах
    loans = EmployeeLoan.query.all()

    data = []
    for l in loans:
        data.append({
            "Ажилтны нэр": l.staff_name,
            "Олгосон огноо": l.date.strftime('%Y-%m-%d') if l.date else "",
            "Олгосон дүн (₮)": l.loan_amount,
            "Төлсөн дүн (₮)": l.total_paid,
            "Үлдэгдэл (₮)": l.loan_amount - l.total_paid,
            "Тайлбар": l.description
        })

    # 2. Pandas ашиглан DataFrame үүсгэх
    df = pd.DataFrame(data)

    # 3. Санах ойд Excel файл үүсгэх
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Зээлийн тайлан')
    output.seek(0)

    filename = f"Толгойт ажилчдын зээл {datetime.now().strftime('%Y%m%d')}.xlsx"

    return send_file(output,
                     download_name=filename,
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/export-balance')
@login_required
def export_balance():
    products = Product.query.filter_by(is_active=True).all()

    data = []
    for p in products:
        total_cost = p.stock * (p.cost_price or 0)
        potential_revenue = p.stock * (p.retail_price or 0)
        total_potential_profit = potential_revenue - total_cost

        data.append({
            'Код (SKU)': p.sku,
            'Барааны нэр': p.name,
            'Ангилал': p.category,
            'Үлдэгдэл': p.stock,
            'Өртөг үнэ': p.cost_price,
            'Жижиглэн үнэ': p.retail_price,
            'Нийт өртөг дүн': total_cost,
            'Боломжит цэвэр ашиг': total_potential_profit
        })

    df = pd.DataFrame(data)

    # Нийт дүнг тооцоолох
    totals = {
        'Код (SKU)': 'НИЙТ ДҮН:',
        'Барааны нэр': '', 'Ангилал': '',
        'Үлдэгдэл': df['Үлдэгдэл'].sum(),
        'Өртөг үнэ': '', 'Жижиглэн үнэ': '',
        'Нийт өртөг дүн': df['Нийт өртөг дүн'].sum(),
        'Боломжит цэвэр ашиг': df['Боломжит цэвэр ашиг'].sum()
    }
    df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Үлдэгдэл')

        workbook = writer.book
        worksheet = writer.sheets['Үлдэгдэл']

        # Форматууд
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
        money_fmt = workbook.add_format({'num_format': '#,##0.00'})
        total_row_fmt = workbook.add_format({'bold': True, 'bg_color': '#FCE4D6', 'border': 1, 'num_format': '#,##0.00'})

        # ЭХНИЙ МӨРИЙГ ЦАРЦААХ (1-р мөр хөдөлгөөнгүй)
        worksheet.freeze_panes(1, 0)

        # Баганын өргөн болон үнийн формат
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            if i >= 4: # Үнийн баганууд
                worksheet.set_column(i, i, column_len, money_fmt)
            else:
                worksheet.set_column(i, i, column_len)

        # Хамгийн сүүлийн мөрийг (Нийт дүн) форматлах
        last_row = len(df)
        for col_num in range(len(df.columns)):
            val = df.iloc[last_row-1, col_num]
            worksheet.write(last_row, col_num, val, total_row_fmt)

    output.seek(0)

    mgl_filename = f"Толгойт Үлдэгдэл{datetime.now().strftime('%Y%m%d')}.xlsx"
    response = send_file(
        output,
        as_attachment=True,
        download_name=mgl_filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(mgl_filename)}"
    return response

@app.route('/export-transactions/<type>')
@login_required
def export_transactions(type):
    from urllib.parse import quote
    from sqlalchemy import or_
    from datetime import datetime, timedelta
    import pandas as pd
    from io import BytesIO

    # 1. URL-аас огнооны мэдээллийг авах
    start_date_str = request.args.get('start_date', '')
    end_date_str = request.args.get('end_date', '')

    # Query бэлдэх
    if type in ['Жижиглэн зарлага', 'Жижиглэн']:
        query = Transaction.query.filter(
            or_(
                Transaction.type.like('%Жижиглэн%'),
                Transaction.type.like('Зарлага%'),
                Transaction.type == 'Багц',
                Transaction.type == 'Багц зарлага'
            )
        )
    elif type in ['Бөөний зарлага', 'Бөөний']:
        query = Transaction.query.filter(
            or_(
                Transaction.type.like('%Бөөний%'),
                Transaction.type == 'Бөөний зарлага',
                Transaction.type == 'Багц'
            )
        )
    else:
        query = Transaction.query.filter(Transaction.type.like(f'%{type}%'))

    # Огноогоор шүүх
    if start_date_str:
        start_dt = datetime.strptime(start_date_str, '%Y-%m-%d')
        query = query.filter(Transaction.timestamp >= start_dt)
    if end_date_str:
        end_dt = datetime.strptime(end_date_str, '%Y-%m-%d')
        query = query.filter(Transaction.timestamp < end_dt + timedelta(days=1))

    transactions = query.order_by(Transaction.timestamp.desc()).all()
    data = []

    for t in transactions:
        desc = t.description if t.description else ""
        batch_no = "-"
        clean_name = desc

        # Гүйлгээний дугаар салгах [2026...]
        if desc.startswith("[") and "]" in desc:
            batch_no = desc[desc.find("[")+1 : desc.find("]")]
            clean_name = desc[desc.find("]")+2:]

        calculated_cost = 0
        p_sku = "N/A"
        p_cat = "Бусад"

        # Өртөг тооцоолох
        if t.product:
            calculated_cost = float(t.product.cost_price or 0)
            p_sku = t.product.sku
            p_cat = t.product.category or "Ангилалгүй"
            clean_name = t.product.name
        else:
            # Багц (Bundle) бол өртгийг нь бүрдэл хэсгээр нь бодно
            bundle_obj = Bundle.query.filter_by(name=clean_name).first()
            if bundle_obj:
                for item in bundle_obj.items:
                    if item.product:
                        calculated_cost += (item.product.cost_price or 0) * item.quantity
                p_sku = "BUNDLE"
                p_cat = "Багц"
            elif "[БҮРТГЭЛГҮЙ]" in clean_name:
                p_sku = "MISC"
                p_cat = "Бүртгэлгүй"
                clean_name = clean_name.replace("[БҮРТГЭЛГҮЙ] ", "")
            elif "[АЖЛЫН ХӨЛС]" in clean_name:
                p_sku = "FEE"
                p_cat = "Ажлын хөлс"
                clean_name = clean_name.replace("[АЖЛЫН ХӨЛС] ", "")

        actual_price = float(t.price) if t.price else 0
        raw_type = t.type or "Зарлага"

        # Төрөл болон Төлбөрийн хэлбэр салгах
        sale_kind = "Бөөний" if "Бөөний" in raw_type else "Жижиглэн"
        
        if "(" in raw_type and ")" in raw_type:
            display_payment = raw_type.split("(")[1].split(")")[0]
        else:
            display_payment = raw_type.replace("Зарлага", "").replace("Бөөний", "").replace("Жижиглэн", "").strip() or "Бэлэн"

        data.append({
            'Огноо': t.timestamp.strftime('%Y-%m-%d'),
            'Гүйлгээ №': batch_no,
            'Борлуулалтын төрөл': sale_kind,
            'Ангилал': p_cat,
            'Барааны код': p_sku,
            'Барааны нэр': clean_name,
            'Төлбөрийн төрөл': display_payment,
            'Тоо ширхэг': t.quantity,
            'Нэгж өртөг': calculated_cost,
            'Зарсан үнэ': actual_price,
            'Нийт дүн': actual_price * t.quantity,
            'Нийт ашиг': (actual_price - calculated_cost) * t.quantity,
            'Ажилтан': t.user.username if t.user else "-"
        })

    if not data:
        flash(f"{type} төрөлд мэдээлэл олдсонгүй!")
        return redirect(url_for('dashboard'))

    df = pd.DataFrame(data)
    cols_order = ['Огноо', 'Гүйлгээ №', 'Борлуулалтын төрөл', 'Ангилал', 'Барааны код', 'Барааны нэр', 
                  'Төлбөрийн төрөл', 'Тоо ширхэг', 'Нэгж өртөг', 'Зарсан үнэ', 'Нийт дүн', 'Нийт ашиг', 'Ажилтан']
    df = df[cols_order]

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        sheet_name = "Гүйлгээний тайлан"
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Форматууд
        money_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        normal_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        total_fmt = workbook.add_format({'bold': True, 'bg_color': '#E7E6E6', 'border': 1, 'num_format': '#,##0', 'align': 'center'})

        last_row = len(df)
        
        # Өгөгдөл бичих ба форматлах
        for row_num in range(last_row):
            for col_num, col_name in enumerate(df.columns):
                val = df.iloc[row_num, col_num]
                fmt = money_fmt if any(x in col_name for x in ['өртөг', 'үнэ', 'дүн', 'ашиг']) else normal_fmt
                worksheet.write(row_num + 1, col_num, val, fmt)

        # Нийт дүнгийн мөр
        worksheet.write(last_row + 1, 0, "НИЙТ ДҮН:", total_fmt)
        # Хоосон нүднүүдэд хүрээ татах
        for c in range(1, 7):
            worksheet.write(last_row + 1, c, "", total_fmt)
            
        worksheet.write(last_row + 1, 7, df['Тоо ширхэг'].sum(), total_fmt)
        worksheet.write(last_row + 1, 8, "", total_fmt)
        worksheet.write(last_row + 1, 9, "", total_fmt)
        worksheet.write(last_row + 1, 10, df['Нийт дүн'].sum(), total_fmt)
        worksheet.write(last_row + 1, 11, df['Нийт ашиг'].sum(), total_fmt)
        worksheet.write(last_row + 1, 12, "", total_fmt)

        # Баганын өргөн
        for i, col in enumerate(df.columns):
            worksheet.set_column(i, i, max(len(col), 12) + 2)

        # Гүйлгээ № нэгтгэх логик
        batch_col = df.columns.get_loc('Гүйлгээ №')
        r = 1
        while r <= last_row:
            curr_val = str(df.iloc[r-1, batch_col])
            if curr_val != "-":
                match_count = 0
                for next_r in range(r, last_row):
                    if str(df.iloc[next_r, batch_col]) == curr_val:
                        match_count += 1
                    else:
                        break
                if match_count > 0:
                    worksheet.merge_range(r, batch_col, r + match_count, batch_col, curr_val, normal_fmt)
                    r += match_count
            r += 1

        worksheet.freeze_panes(1, 0)

    output.seek(0)
    safe_filename = quote(f"{type} Толгойт {start_date_str}.xlsx")
    return send_file(output, as_attachment=True, download_name=safe_filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
@app.route('/export-expense-report/<category>')
@login_required
def export_expense_report(category):
    start_date_str = request.args.get('start_date')
    end_date_str = request.args.get('end_date')

    # 1. Өгөгдөл шүүх хэсэг
    if category == "Бүгд":
        query = Expense.query
    else:
        query = Expense.query.filter(Expense.category == category)

    date_range_label = ""
    if start_date_str:
        try:
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            query = query.filter(Expense.date >= start_date)
            date_range_label += start_date_str
        except: pass

    if end_date_str:
        try:
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
            query = query.filter(Expense.date < end_date + timedelta(days=1))
            date_range_label += f"_to_{end_date_str}"
        except: pass

    if not date_range_label:
        date_range_label = datetime.now().strftime('%Y-%m-%d')

    expenses = query.order_by(Expense.date.desc()).all()

    # Баганын нэрийг тодорхойлох
    amount_col = 'Олгосон дүн' if category == "Ажлын хөлс" else 'Зардлын дүн'

    # 2. Тайлангийн өгөгдлийг жагсаалт болгох
    report_data = []
    for e in expenses:
        # Ажилтан бүртгэсэн эсэхийг хамгаалалттай шалгах
        staff = "-"
        if hasattr(e, 'user') and e.user:
            staff = e.user.username

        report_data.append({
            'Огноо': e.date.strftime('%Y-%m-%d') if e.date else "-",
            'Төрөл': e.category,
            'Тайлбар': e.description or "-",
            amount_col: e.amount,
            'Бүртгэсэн': staff
        })

    df = pd.DataFrame(report_data)

    # 3. Нийт дүнгийн мөрийг нэмэх
    if not df.empty:
        total_sum = df[amount_col].sum()
        totals = {
            'Огноо': 'НИЙТ:',
            'Төрөл': '',
            'Тайлбар': '',
            amount_col: total_sum,
            'Бүртгэсэн': ''
        }
        df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)

    # 4. Excel үүсгэх болон Borders (Хүрээ) тохируулах
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        sheet_name = category[:31]
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # --- ФОРМАТУУД ТОХИРУУЛАХ ---
        # Энгийн нүдний хүрээ
        border_fmt = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})

        # Мөнгөн дүнгийн формат (Таслалтай + Хүрээтэй)
        money_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1, 'align': 'right'})

        # Толгой хэсгийн формат (Ногоон өнгө + Хүрээ)
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#D7E4BC', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })

        # Нийт дүнгийн формат (Улбар шар өнгө + Хүрээ)
        total_row_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#FCE4D6', 'border': 1, 'num_format': '#,##0'
        })

        # Толгой мөрийг форматлах
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        # Бүх өгөгдлийг хүрээтэй болгох (Нийт дүнгээс бусад мөрүүд)
        for row_num in range(1, len(df)):
            for col_num in range(len(df.columns)):
                val = df.iloc[row_num-1, col_num]
                # Сүүлийн мөр биш бол энгийн хүрээ ашиглана
                if row_num < len(df):
                    if df.columns[col_num] == amount_col:
                        worksheet.write(row_num, col_num, val, money_fmt)
                    else:
                        worksheet.write(row_num, col_num, val, border_fmt)

        # Сүүлийн мөрийг (Нийт дүн) форматлах
        last_row_idx = len(df)
        for col_num in range(len(df.columns)):
            worksheet.write(last_row_idx, col_num, df.iloc[-1, col_num], total_row_fmt)

        # Баганын өргөнийг автоматаар тохируулах
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(col)) + 5
            worksheet.set_column(i, i, column_len)

        worksheet.freeze_panes(1, 0) # Дээд мөрийг царцаах

    output.seek(0)

    # Файлын нэрийг бэлдэх
    filename = f"{category} Толгойт {date_range_label}.xlsx"
    response = send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    # Монгол нэрний кодыг дэмжүүлэх
    from urllib.parse import quote
    response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(filename)}"

    return response

@app.route('/export-labor-report')
@login_required
def export_labor_report():
    if current_user.role != 'admin':
        return "Хандах эрхгүй", 403

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    # Өгөгдлийн сангаас шүүх
    query = LaborFee.query
    if start_date and end_date:
        query = query.filter(
            LaborFee.timestamp >= start_date,
            LaborFee.timestamp <= end_date + " 23:59:59"
        )

    fees = query.order_by(LaborFee.timestamp.desc()).all()

    data = []
    total_sum = 0

    for f in fees:
        # ШҮҮЛТҮҮР: Хэрэв тайлбар нь 'Бүртгэлгүй бараа' бол алгасах
        if f.description == "Бүртгэлгүй бараа":
            continue

        data.append({
            'Огноо': f.timestamp.strftime('%Y-%m-%d'),
            'Ажлын тайлбар': f.description,
            'Ажилтан': f.staff_name,
            'Дүн': f.amount
        })
        total_sum += f.amount

    # Pandas DataFrame үүсгэх
    df = pd.DataFrame(data)

    # Хамгийн доор нь НИЙТ ДҮН-г нэмэх (Бүртгэлгүй барааны дүн ороогүй нийлбэр)
    if not df.empty:
        summary = pd.DataFrame([{'Огноо': 'НИЙТ ДҮН:', 'Ажлын тайлбар': '', 'Ажилтан': '', 'Дүн': total_sum}])
        df = pd.concat([df, summary], ignore_index=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Salary')

    output.seek(0)
    filename = f"тТолгойт ажлын хөлсний {start_date if start_date else 'All'}.xlsx"

    return send_file(output, download_name=filename, as_attachment=True)

@app.route('/export-salary-report')
@login_required
def export_salary_report():
    start_date_str = request.args.get('start_date')
    end_date_str = request.args.get('end_date')

    # 1. Цалин төрөлтэй зардлуудыг шүүх
    query = Expense.query.filter_by(category='Цалин')

    # 2. Огнооны шүүлтүүрийг засах (Text-ээс Date формат руу)
    if start_date_str and end_date_str:
        try:
            # Өдрийн эхлэл 00:00:00-аас өдрийн төгсгөл 23:59:59 хүртэл шүүх
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d').replace(hour=23, minute=59, second=59)
            query = query.filter(Expense.date >= start_date, Expense.date <= end_date)
        except ValueError:
            pass

    salaries = query.all()

    # 3. Хэрэв өгөгдөл байхгүй бол хоосон файл гаргахгүйн тулд шалгах
    if not salaries:
        data = [{"Мэдээлэл": "Сонгосон хугацаанд цалингийн гүйлгээ олдсонгүй"}]
    else:
        data = []
        for s in salaries:
            data.append({
                "Огноо": s.date.strftime('%Y-%m-%d %H:%M'),
                "Төрөл": s.category,
                "Тайлбар": s.description,
                "Олгосон дүн": s.amount
            })

    # Excel үүсгэх
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Цалингийн_Тайлан')
    output.seek(0)

    # 4. Файлын нэрийг Монгол болгох
    filename = f"Толгойт ажилчдын цалин {datetime.now().strftime('%Y%m%d')}.xlsx"

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

@app.route('/export-return-report')
@login_required
def export_return_report():
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    # Зөвхөн "буцаалт" төрлийг шүүнэ
    query = Transaction.query.filter(Transaction.type.ilike('буцаалт'))

    if start_date and end_date:
        query = query.filter(Transaction.date.between(start_date, end_date))

    transactions = query.all()

    # 1. Өгөгдлөө бэлдэхдээ багануудыг яг салгаж бичнэ
    data = []
    total_refund = 0

    for t in transactions:
        refund_val = abs(t.amount) if t.amount else (t.quantity * t.price)
        total_refund += refund_val

        data.append({
            "Огноо": t.date.strftime('%Y-%m-%d'),
            "Ангилал": t.product.category if t.product else "-", # НЭМЭГДЭВ
            "SKU / Код": t.product.sku if t.product else "",
            "Барааны нэр": t.product.name if t.product else "",
            "Төрөл": t.type,
            "Тоо ширхэг": t.quantity,
            "Буцаасан дүн": refund_val,
            "Тайлбар": t.note if hasattr(t, 'note') and t.note else "",
            "Ажилтан": t.user.username if t.user else "Unknown"
        })

    df = pd.DataFrame(data)
    if not df.empty:
        # Дарааллыг баталгаажуулах: Огноо -> Ангилал -> SKU
        column_order = ["Огноо", "Ангилал", "SKU / Код", "Барааны нэр", "Төрөл", "Тоо ширхэг", "Буцаасан дүн", "Тайлбар", "Ажилтан"]
        df = df[column_order]

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Буцаалтын тайлан')

        workbook  = writer.book
        worksheet = writer.sheets['Буцаалтын тайлан']

        # Формат тохиргоо
        money_fmt = workbook.add_format({'num_format': '#,##0₮', 'bold': True, 'align': 'right'})
        total_label_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1})

        # Нийт дүнг хамгийн доор нь нэмэх
        last_row = len(df) + 1
        worksheet.write(last_row, 4, "НИЙТ БУЦААЛТ:", total_label_fmt)
        worksheet.write(last_row, 5, total_refund, money_fmt)

        # Баганын өргөнийг автоматаар тааруулах
        for i, col in enumerate(df.columns):
            worksheet.set_column(i, i, len(col) + 10)

    output.seek(0)
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"Толгойт бараа буцаалт {start_date if start_date else 'all'}.xlsx"
    )

# --- 1. ХУУЧИН НУМ ТАЙЛАН (НИЙТ ДҮНТЭЙ) ---
@app.route('/export-old-bow')
@login_required
def export_old_bow():
    try:
        start_date = request.args.get('start_date')
        end_date = request.args.get('end_date')

        query = OldBow.query
        if start_date and end_date:
            query = query.filter(OldBow.date >= start_date, OldBow.date <= end_date + " 23:59:59")

        reports = query.order_by(OldBow.id.desc()).all()

        if not reports:
            flash("Тухайн хугацаанд мэдээлэл олдсонгүй.")
            return redirect(url_for('dashboard'))

        data = []
        t_qty = 0
        t_amount = 0

        for r in reports:
            # None утга ирэхээс сэргийлж 0 болгоно
            qty = int(r.quantity) if r.quantity else 0
            price = float(r.purchase_price) if r.purchase_price else 0
            subtotal = qty * price

            # Нийт дүнг нэмж бодох
            t_qty += qty
            t_amount += subtotal

            data.append({
                "Огноо": r.date,
                "Барааны нэр": r.product_name,
                "Код/SKU": r.sku,
                "Авсан үнэ": price,
                "Зарах үнэ": float(r.retail_price) if r.retail_price else 0,
                "Тоо ширхэг": qty,
                "Нийт дүн": subtotal,
                "Бүртгэсэн": r.user.username if r.user else "Систем"
            })

        # НИЙТ ДҮН-г жагсаалтын төгсгөлд Dictionary хэлбэрээр нэмэх
        total_row = {
            "Огноо": "НИЙТ ДҮН",
            "Барааны нэр": "",
            "Код/SKU": "",
            "Авсан үнэ": "",
            "Зарах үнэ": "",
            "Тоо ширхэг": t_qty,
            "Нийт дүн": t_amount,
            "Бүртгэсэн": ""
        }
        data.append(total_row)

        # DataFrame үүсгэх
        df = pd.DataFrame(data)

        output = io.BytesIO()
        # xlsxwriter ашиглан файлыг үүсгэх
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Хуучин нум')

            # Excel-ийн форматыг жаахан сайжруулах (заавал биш)
            workbook  = writer.book
            worksheet = writer.sheets['Хуучин нум']
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'})
            # Сүүлийн мөрийг тодруулах формат
            total_format = workbook.add_format({'bold': True, 'bg_color': '#FCE4D6'})

            # Хамгийн сүүлийн мөрөнд формат өгөх
            last_row = len(df)
            worksheet.set_row(last_row, None, total_format)

        output.seek(0)
        return send_file(
            output,
            download_name=f"Жажин Хуучин нум {datetime.now().strftime('%Y%m%d')}.xlsx",
            as_attachment=True
        )

    except Exception as e:
        print(f"Excel Error: {str(e)}") # Консол дээр алдааг хэвлэх
        return f"Алдаа гарлаа: {str(e)}"

# --- 2. САЛБАРЫН ОРЛОГО ТАЙЛАН (НИЙТ ДҮНТЭЙ) ---

@app.route('/export-internal-income')
@login_required
def export_internal_income():
    try:
        start_date = request.args.get('start_date')
        end_date = request.args.get('end_date')

        query = Transaction.query.filter_by(type='Орлого')
        if start_date and end_date:
            query = query.filter(Transaction.timestamp >= start_date,
                                 Transaction.timestamp <= end_date + " 23:59:59")

        items = query.order_by(Transaction.timestamp.desc()).all()

        data = []
        t_qty = 0
        t_amount = 0

        for i in items:
            current_price = 0
            if i.price and float(i.price) > 0:
                current_price = float(i.price)
            elif i.product and i.product.cost_price:
                current_price = float(i.product.cost_price)

            qty = int(i.quantity) if i.quantity else 0
            subtotal = qty * current_price

            t_qty += qty
            t_amount += subtotal

            data.append([
                i.timestamp.strftime('%Y-%m-%d'),
                i.product.sku if i.product else "Устсан",
                i.product.name if i.product else i.description,
                qty,
                current_price,
                subtotal,
                i.description or "",
                i.user.username if i.user else "-"
            ])

        if not data:
            flash("Мэдээлэл олдсонгүй.")
            return redirect(url_for('internal_income_list'))

        columns = ["Огноо", "Код (SKU)", "Барааны нэр", "Тоо", "Өртөг", "Нийт өртөг", "Тайлбар", "Ажилтан"]

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Салбарын орлого')

            # Форматууд
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
            money_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
            normal_fmt = workbook.add_format({'border': 1})
            total_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F2F2F2'})
            total_money_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F2F2F2', 'num_format': '#,##0'})

            # 1. Толгой хэсгийг бичих
            for col_num, col_name in enumerate(columns):
                worksheet.write(0, col_num, col_name, header_fmt)

            # 2. Өгөгдлийг бичих (Зөвхөн юмтай нүднүүдийг хүрээлнэ)
            row_idx = 1
            for row_data in data:
                for col_idx, value in enumerate(row_data):
                    # Өртөг болон Нийт өртөг багана бол мөнгөн формат ашиглана
                    if col_idx in [4, 5]:
                        worksheet.write(row_idx, col_idx, value, money_fmt)
                    else:
                        worksheet.write(row_idx, col_idx, value, normal_fmt)
                row_idx += 1

            # 3. НИЙТ ДҮН мөрийг бичих
            worksheet.write(row_idx, 0, "НИЙТ ДҮН", total_fmt)
            for c in range(1, 8): # Хоосон нүднүүдэд хүрээ өгөх
                worksheet.write(row_idx, c, "", total_fmt)
            worksheet.write(row_idx, 3, t_qty, total_fmt)
            worksheet.write(row_idx, 5, t_amount, total_money_fmt)

            # 4. Баганын өргөнийг автоматаар тааруулах
            worksheet.set_column('A:A', 12)
            worksheet.set_column('B:B', 15)
            worksheet.set_column('C:C', 30)
            worksheet.set_column('D:D', 8)
            worksheet.set_column('E:F', 15)
            worksheet.set_column('G:G', 25)
            worksheet.set_column('H:H', 15)

        output.seek(0)
        return send_file(output,
                         download_name=f"ТОЛГОЙТ САЛБАРЫН ОРЛОГО_{datetime.now().strftime('%Y%m%d')}.xlsx",
                         as_attachment=True,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        print(f"EXPORT ERROR: {str(e)}")
        return f"Алдаа: {str(e)}"

# --- ХЭРЭГЛЭГЧИЙН УДИРДЛАГА ---

@app.route('/users')
@login_required
def users_list():
    if current_user.role != 'admin': return "Эрх хүрэхгүй", 403
    return render_template('users.html', users=User.query.all())

@app.route('/add_user', methods=['POST'])
@login_required
def add_user():
    if current_user.role != 'admin':
        flash("Зөвхөн админ хэрэглэгч нэмэх эрхтэй!")
        return redirect(url_for('users_list'))

    username = request.form.get('username')
    password = request.form.get('password')
    role = request.form.get('role')

    # Хэрэглэгч байгаа эсэхийг шалгах
    existing_user = User.query.filter_by(username=username).first()
    if existing_user:
        flash("Энэ хэрэглэгчийн нэр бүртгэлтэй байна!")
        return redirect(url_for('users_list'))

    # ЧУХАЛ: Нууц үгийг шифрлэж байна
    hashed_password = generate_password_hash(password)

    # password=password биш password=hashed_password болгож засав
    new_user = User(username=username, password=hashed_password, role=role)

    db.session.add(new_user)
    db.session.commit()
    flash(f"'{username}' хэрэглэгч амжилттай нэмэгдлээ.")
    return redirect(url_for('users_list'))

@app.route('/delete_user/<int:id>')  # HTML дээр id гэж дамжуулсан тул энд id байна
@login_required
def delete_user(id):
    if current_user.role != 'admin':
        flash("Танд устгах эрх байхгүй!")
        return redirect(url_for('dashboard'))

    user_to_delete = User.query.get_or_404(id)

    # Үндсэн админ 'Sodoo'-г устгахаас хамгаалах
    if user_to_delete.username == 'Sodoo':
        flash("Үндсэн админыг устгаж болохгүй!")
        return redirect(url_for('users_list'))

    try:
        # Хэрэглэгч устах үед түүний хийсэн гүйлгээг 'Unknown' болгох
        Transaction.query.filter_by(user_id=user_to_delete.id).update({"user_id": None})

        db.session.delete(user_to_delete)
        db.session.commit()
        flash(f"'{user_to_delete.username}' амжилттай устлаа.")
    except Exception as e:
        db.session.rollback()
        flash(f"Алдаа гарлаа: {str(e)}")

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

@app.route('/reports')
@login_required
def reports():
    # 1. Огноог авах
    selected_date_str = request.args.get('date', date.today().isoformat())
    try:
        selected_date = date.fromisoformat(selected_date_str)
    except:
        selected_date = date.today()
        selected_date_str = selected_date.isoformat()
    
    # 2. Тухайн өдрийн гүйлгээг сангаас татах
    transactions = Transaction.query.filter(func.date(Transaction.timestamp) == selected_date).all()
    
    # 3. Төлбөрийн төрлөөр багцлах бэлдэц
    payment_groups = {
        "Бэлнээр": [],
        "Дансаар": [],
        "Картаар": [],
        "Компаниар": [],
        "Бусад": []
    }
    
    total_sales = 0
    payment_summary = {key: 0 for key in payment_groups.keys()}

    # 4. Гүйлгээнүүдийг төрлөөр нь ялгах
    for t in transactions:
        amount = t.price * t.quantity
        total_sales += amount
        
        t_type = t.type.lower()
        if "бэлэн" in t_type:
            key = "Бэлнээр"
        elif "данс" in t_type:
            key = "Дансаар"
        elif "карт" in t_type:
            key = "Картаар"
        elif "компани" in t_type or "байгууллага" in t_type:
            key = "Компаниар"
        else:
            key = "Бусад"
            
        payment_groups[key].append(t)
        payment_summary[key] += amount

    # 5. HTML рүү бүх өгөгдлийг дамжуулах (ЭНД payment_groups-ийг заавал бичнэ)
    return render_template('reports.html', 
                           payment_groups=payment_groups, 
                           payment_summary=payment_summary,
                           total_sales=total_sales,
                           selected_date=selected_date_str)

if __name__ == "__main__":
    with app.app_context():
        db.create_all()  # Энэ мөр нь байхгүй байгаа бүх хүснэгтийг (old_bow) үүсгэнэ
    app.run(debug=True)
