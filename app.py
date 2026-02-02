import os
import io
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session
from flask_sqlalchemy import SQLAlchemy
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
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'))
    type = db.Column(db.String(100))
    quantity = db.Column(db.Float)
    timestamp = db.Column(db.DateTime, default=datetime.now)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    product = db.relationship('Product', backref='transactions')
    user = db.relationship('User', backref='transactions')

class Expense(db.Model):
    __tablename__ = 'expense'
    __table_args__ = {'extend_existing': True}  # Энэ мөр давхардал алдаанаас хамгаална
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(50), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    description = db.Column(db.Text)
    date = db.Column(db.DateTime, default=datetime.utcnow)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    user = db.relationship('User', backref='expense_records')

class LaborFee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    description = db.Column(db.String(200), nullable=False) # Хийсэн ажил
    amount = db.Column(db.Float, nullable=False)           # Хөлс
    staff_name = db.Column(db.String(100))                 # Ажилтан
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)

# --- 1. MODEL ХЭСЭГТ НЭМЭХ ---
class OldBow(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_name = db.Column(db.String(200), nullable=False)
    sku = db.Column(db.String(50))
    category = db.Column(db.String(50), default="Хуучин нум")
    purchase_price = db.Column(db.Float, nullable=False)
    retail_price = db.Column(db.Float)
    quantity = db.Column(db.Integer, default=1)
    date = db.Column(db.String(50))

class EmployeeLoan(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    staff_name = db.Column(db.String(100), nullable=False)
    loan_amount = db.Column(db.Float, default=0.0)      # Анх авсан зээл
    total_paid = db.Column(db.Float, default=0.0)       # Буцааж төлсөн нийт дүн
    description = db.Column(db.String(200))
    date = db.Column(db.DateTime, default=datetime.utcnow)

    @property
    def remaining_balance(self):
        return self.loan_amount - self.total_paid

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

from flask import jsonify # Файлын дээр заавал нэмээрэй

@app.route('/add_transaction', methods=['POST'])
@login_required
def add_transaction():
    # JSON өгөгдөл ирсэн эсэхийг шалгах (Олон бараа нэг дор сагснаас ирэх үед)
    if request.is_json:
        data = request.get_json()
        if data and 'items' in data:
            try:
                for item in data['items']:
                    p_id = item.get('product_id')
                    t_type = item.get('type')
                    qty = float(item.get('quantity') or 0)
                    product = Product.query.get(p_id)
                    
                    if product:
                        if t_type in ['Орлого', 'буцаалт']:
                            product.stock += qty
                        else:
                            product.stock -= qty
                        
                        db.session.add(Transaction(
                            product_id=p_id, 
                            type=t_type, 
                            quantity=qty, 
                            user_id=current_user.id
                        ))
                
                db.session.commit()
                return jsonify({"success": True, "message": "Бүх гүйлгээ амжилттай хадгалагдлаа."})
            except Exception as e:
                db.session.rollback() # Алдаа гарвал бүх үйлдлийг цуцална
                return jsonify({"success": False, "message": str(e)}), 500

    # Хуучин Form-оор ганц бараа ирэх үеийн логикийг хэвээр үлдээв
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
    # Бүх барааг авах
    products = Product.query.order_by(Product.name).all()
    # Сүүлийн 10 тооллогын түүхийг авах
    history = Transaction.query.filter(Transaction.type.like('%Тооллого%'))\
              .order_by(Transaction.timestamp.desc()).limit(10).all()
    return render_template('inventory.html', products=products, history=history)

@app.route('/do_inventory', methods=['POST'])
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
    data = request.json
    items = data.get('items', [])
    if not items: return jsonify({"status": "error"}), 400
    
    for item in items:
        product = Product.query.get(item['product_id'])
        if product:
            qty = float(item['quantity'])
            
            # Үлдэгдэл тооцох
            if item['type'] == 'Орлого':
                product.stock += qty
            else: # Жижиглэн, Бөөний, Өртгөөр бүгд хасагдана
                product.stock -= qty
            
            # Гүйлгээ хадгалах (timestamp ашиглана)
            db.session.add(Transaction(
                product_id=product.id,
                type=item['type'],
                quantity=qty,
                timestamp=datetime.now(), 
                user_id=current_user.id
            ))
    db.session.commit()
    flash(f"{len(items)} гүйлгээ амжилттай бүртгэгдлээ.")
    return jsonify({"status": "success"}), 200

@app.route('/special_transfer', methods=['GET', 'POST'])
@login_required
def special_transfer():
    if request.method == 'POST':
        product_id = request.form.get('product_id')
        quantity = float(request.form.get('quantity') or 0)
        note = request.form.get('note') # Хаашаа шилжүүлж буй тайлбар

        product = Product.query.get(product_id)
        if product and quantity > 0:
            product.stock -= quantity # Үлдэгдэл хасах
            
            # Гүйлгээг 'Өртгөөр зарлага' төрлөөр хадгалах
            new_tx = Transaction(
                product_id=product.id,
                type='Өртгөөр зарлага',
                quantity=quantity,
                timestamp=datetime.now(),
                user_id=current_user.id
            )
            db.session.add(new_tx)
            db.session.commit()
            flash(f'{product.name} - {quantity} ш өртгөөр амжилттай шилжлээ.')
            return redirect(url_for('dashboard'))

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
        product_id = request.form.get('product_id')
        quantity = float(request.form.get('quantity'))
        price = float(request.form.get('price'))
        
        product = Product.query.get_or_404(product_id)
        
        # 1. Үлдэгдэл нэмэх
        product.stock += quantity
        
        # 2. Гүйлгээг хасах дүнгээр бүртгэх (Орлогоос хасагдана)
        amount = quantity * price
        new_transaction = Transaction(
            product_id=product.id,
            quantity=quantity,
            type='Буцаалт',
            amount=-amount, # Хасах утга
            user_id=current_user.id,
            date=datetime.now()
        )
        
        db.session.add(new_transaction)
        db.session.commit()
        flash(f"'{product.name}' барааны буцаалт амжилттай. Орлогоос {amount:,.0f}₮ хасагдлаа.")
        return redirect(url_for('dashboard'))

    products = Product.query.all()
    return render_template('returns.html', products=products)

@app.route('/buy-old-bow', methods=['GET', 'POST'])
@login_required
def buy_old_bow():
    if request.method == 'POST':
        name = request.form.get('name')
        sku = request.form.get('sku')
        cost = float(request.form.get('cost_price'))
        retail = float(request.form.get('retail_price'))
        qty = float(request.form.get('stock')) # Product дээр stock нь Float байна

        try:
            # 1. Хуучин нумны түүхэнд хадгалах (OldBow модель)
            new_old_bow = OldBow(
                product_name=name,
                sku=sku,
                category="Хуучин нум",
                purchase_price=cost,
                retail_price=retail,
                quantity=int(qty),
                date=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            )
            
            # 2. Үндсэн барааны жагсаалт руу нэмэх (Ингэснээр Dashboard дээр харагдана)
            new_product = Product(
                name=f"[Хуучин] {name}",
                sku=sku if sku else f"OLD-{datetime.now().strftime('%m%d%H%M')}",
                category="Хуучин нум",
                cost_price=cost,      # purchase_price биш cost_price гэж зассан
                retail_price=retail,
                wholesale_price=retail, 
                stock=qty,
                is_active=True
            )

            # 3. Кассаас зардал хасах
            new_expense = Expense(
                category="Хуучин нум авалт",
                amount=cost * qty,
                description=f"{name} худалдаж авсан",
                date=datetime.utcnow(),
                user_id=current_user.id
            )

            db.session.add(new_old_bow)
            db.session.add(new_product)
            db.session.add(new_expense)
            db.session.commit()
            
            flash(f"'{name}' амжилттай бүртгэгдэж, Dashboard дээр нэмэгдлээ.")
            return redirect(url_for('dashboard'))
            
        except Exception as e:
            db.session.rollback()
            flash(f"Алдаа гарлаа: {str(e)}")
            return redirect(url_for('buy_old_bow'))

    return render_template('buy_old_bow.html')
    
# --- 2. ХУУЧИН БАРААНЫ ТАЙЛАН ---
@app.route('/old-bow-report')
@login_required
def old_bow_report():
    # Зөвхөн "Хуучин нум авалт" төрөлтэй зардлуудыг харах
    reports = Expense.query.filter_by(category="Хуучин нум авалт").order_by(Expense.date.desc()).all()
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

# --- ТАЙЛАН, СТАТИСТИК ---

@app.route('/statistics')
@login_required
def statistics():
    # 1. Огнооны шүүлтүүр авах (Хуучин хэвээрээ)
    start_date_str = request.args.get('start_date')
    end_date_str = request.args.get('end_date')
    
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
    
    # 2. Графикийн өгөгдөл боловсруулах (Хуучин хэвээрээ)
    for date_str in dates:
        # Борлуулалт
        sales = db.session.query(db.func.sum(Transaction.quantity * db.case(
                (Transaction.type == 'Бөөний зарлага', Product.wholesale_price),
                (Transaction.type == 'Жижиглэн зарлага', Product.retail_price),
                else_=0
            ))).join(Product).filter(db.func.date(Transaction.timestamp) == date_str).scalar() or 0
        
        # Өртөг
        cost_sum = db.session.query(db.func.sum(Transaction.quantity * Product.cost_price)).\
            join(Product).filter(Transaction.type.like('%зарлага%'), 
            db.func.date(Transaction.timestamp) == date_str).scalar() or 0
        
        # Зардал
        daily_expense = db.session.query(db.func.sum(Expense.amount)).\
            filter(Expense.category != "Ажлын хөлс", db.func.date(Expense.date) == date_str).scalar() or 0

        sales_data.append(float(sales))
        expense_data.append(float(daily_expense))
        profit_data.append(float(sales - cost_sum - daily_expense))

    # 3. Нийт ТОП 5 бараа (Ширхэгээр - Хуучин хэвээрээ)
    top_products_all = db.session.query(Product.name, db.func.sum(Transaction.quantity)).\
        join(Transaction).filter(Transaction.type.like('%зарлага%')).\
        group_by(Product.name).order_by(db.func.sum(Transaction.quantity).desc()).limit(5).all()
    
    top_labels = [p[0] for p in top_products_all]
    top_values = [int(p[1]) for p in top_products_all]

    # ================= ШИНЭ ХЭСЭГ: АНГИЛАЛ БҮРИЙН ТОП 5 =================
    # Бүх өвөрмөц ангиллуудыг авах
    categories = db.session.query(Product.category).distinct().all()
    stats_data = {}

    for cat in categories:
        category_name = cat[0] if cat[0] else "Ангилалгүй"
        
        # Тухайн ангиллын хамгийн их борлуулагдсан 5 бараа
        category_top = db.session.query(
            Product.name, 
            db.func.sum(Transaction.quantity).label('total')
        ).join(Transaction).filter(
            Product.category == category_name,
            Transaction.type.like('%зарлага%')
        ).group_by(Product.name).order_by(db.desc('total')).limit(5).all()
        
        if category_top:
            stats_data[category_name] = category_top
    # =================================================================

    return render_template('statistics.html', 
                           dates=dates, 
                           sales=sales_data, 
                           profit=profit_data, 
                           expenses=expense_data,
                           top_labels=top_labels,
                           top_values=top_values,
                           stats_data=stats_data,  # Шинэ дата
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

    filename = f"Loan_Report_{datetime.now().strftime('%Y%m%d')}.xlsx"
    
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
    
    mgl_filename = f"Үлдэгдэл_Тайлан_{datetime.now().strftime('%Y%m%d')}.xlsx"
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
    start_date_str = request.args.get('start_date')
    end_date_str = request.args.get('end_date')
    
    query = Transaction.query.filter(Transaction.type == type)
    
    date_range_label = ""
    if start_date_str:
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
        query = query.filter(Transaction.timestamp >= start_date)
        date_range_label += start_date_str
    
    if end_date_str:
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
        query = query.filter(Transaction.timestamp < end_date + timedelta(days=1))
        date_range_label += f"_to_{end_date_str}"
    
    if not date_range_label:
        date_range_label = datetime.now().strftime('%Y-%m-%d')

    transactions = query.all()
    
    data = []
    for t in transactions:
        cost_price = t.product.cost_price if t.product else 0
        sell_price = 0
        if "Бөөний" in t.type:
            sell_price = t.product.wholesale_price if t.product else 0
        elif "Жижиглэн" in t.type:
            sell_price = t.product.retail_price if t.product else 0
        
        unit_profit = sell_price - cost_price
        total_profit = unit_profit * t.quantity
        
        data.append({
            'Огноо': t.timestamp.strftime('%Y-%m-%d'),
            'Ангилал': t.product.category if t.product else "-", # НЭМЭГДЭВ
            'Барааны код': t.product.sku if t.product else "-",
            'Барааны нэр': t.product.name if t.product else "Устгагдсан",
            'Гүйлгээний төрөл': t.type,
            'Тоо ширхэг': t.quantity,
            'Нэгж өртөг': cost_price,
            'Зарах үнэ': sell_price,
            'Нийт ашиг': total_profit if sell_price > 0 else 0,
            'Ажилтан': t.user.username if t.user else "-"
        })
        
    df = pd.DataFrame(data)
    # Нийт дүн нэмэх хэсэгт баганы нэрийг зөрүүлэхгүй тулд:
    if not df.empty and type != 'Орлого':
        totals = {
            'Огноо': 'НИЙТ ДҮН:', 'Ангилал': '', 'Барааны код': '', 'Барааны нэр': '', 'Гүйлгээний төрөл': '',
            'Тоо ширхэг': df['Тоо ширхэг'].sum(),
            'Нэгж өртөг': '', 'Зарах үнэ': '', 
            'Нийт ашиг': df['Нийт ашиг'].sum(),
            'Ажилтан': ''
        }
        df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)
        
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        sheet_name = f"{type} Тайлан"
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name[:31]]
        
        # --- ФОРМАТУУД ТОХИРУУЛАХ ---
        border_fmt = workbook.add_format({'border': 1, 'align': 'left'})
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
        money_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1, 'align': 'right'})
        total_row_fmt = workbook.add_format({'bold': True, 'bg_color': '#FCE4D6', 'border': 1, 'num_format': '#,##0'})

        # Толгой мөрийг форматлах
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        # Бүх нүдэнд хүрээ болон мөнгөн дүнгийн форматлах
        for row_num in range(1, len(df)):
            for col_num in range(len(df.columns)):
                val = df.iloc[row_num-1, col_num]
                col_name = df.columns[col_num]
                
                # Хэрэв мөнгөн дүнтэй багана бол
                if any(x in col_name for x in ['өртөг', 'үнэ', 'ашиг']):
                    worksheet.write(row_num, col_num, val, money_fmt)
                else:
                    worksheet.write(row_num, col_num, val, border_fmt)

        # Хамгийн сүүлийн мөрийг (Total) форматлах
        if not df.empty:
            last_row_idx = len(df)
            for col_num in range(len(df.columns)):
                worksheet.write(last_row_idx, col_num, df.iloc[-1, col_num], total_row_fmt)

        # Баганын өргөнийг автоматаар тохируулах
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(col)) + 4
            worksheet.set_column(i, i, column_len)

        worksheet.freeze_panes(1, 0)

    output.seek(0)
    
    mgl_filename = f"{type}_Тайлан_{date_range_label}.xlsx"
    response = send_file(output, as_attachment=True, download_name=mgl_filename)
    response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(mgl_filename)}"
    return response
    
@app.route('/export-inventory-report')
@login_required
def export_inventory_report():
    transactions = Transaction.query.filter(Transaction.type.like('Тооллого%')).all()
    data = [{
        "Огноо": t.timestamp.strftime('%Y-%m-%d %H:%M'), 
        "Ангилал": t.product.category if t.product else "-", # НЭМЭГДЭВ
        "Барааны код": t.product.sku if t.product else "-", # НЭМЭГДЭВ
        "Бараа": t.product.name if t.product else "Устгагдсан", 
        "Зөрүү": t.type, 
        "Тоо": t.quantity
    } for t in transactions]
    
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="Inventory_Report.xlsx")

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
    filename = f"{category}_Report_{date_range_label}.xlsx"
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

    # Огноо сонгоогүй бол бүх өгөгдлийг шүүнэ
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
        data.append({
            'Огноо': f.timestamp.strftime('%Y-%m-%d %H:%M'),
            'Ажлын тайлбар': f.description,
            'Ажилтан': f.staff_name,
            'Дүн': f.amount
        })
        total_sum += f.amount

    df = pd.DataFrame(data)
    
    # Хамгийн доор нь нийт дүнг нэмэх
    if not df.empty:
        summary = pd.DataFrame([{'Огноо': 'НИЙТ ДҮН:', 'Дүн': total_sum}])
        df = pd.concat([df, summary], ignore_index=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Salary')
    
    output.seek(0)
    filename = f"Labor_Report_{start_date if start_date else 'All'}.xlsx"
    return send_file(output, download_name=filename, as_attachment=True)

@app.route('/export-salary-report')
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
    filename = f"Tsalingiin_Tailan_{datetime.now().strftime('%Y%m%d')}.xlsx"

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
            "Огноо": t.date.strftime('%Y-%m-%d %H:%M'),
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
        download_name=f"Return_Report_{start_date if start_date else 'all'}.xlsx"
    )

@app.route('/export-old-bow')
@login_required
def export_old_bow():
    try:
        start_date = request.args.get('start_date')
        end_date = request.args.get('end_date')

        # Хэрэв таны хүснэгтийн нэр OldBowPurchase биш бол OldBow гэж шалгана уу
        # Таны app.py-ийн дээр class OldBow(db.Model) гэж байгаа бол энд OldBow байна
        purchases = OldBow.query.filter(
            OldBow.date >= start_date,
            OldBow.date <= end_date
        ).all()

        data = []
        for p in purchases:
            data.append({
                "Огноо": p.date,
                "Барааны нэр": p.product_name,
                "Тоо ширхэг": p.quantity,
                "Авсан үнэ": p.purchase_price,
                "Нийт": p.quantity * p.purchase_price,
                "Тайлбар": p.description
            })

        if not data:
            flash("Тухайн хугацаанд өгөгдөл олдсонгүй.")
            return redirect(url_for('dashboard'))

        import pandas as pd
        from io import BytesIO

        df = pd.DataFrame(data)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Хуучин нум')
        
        output.seek(0)
        
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"Old_Bow_Report_{start_date}_{end_date}.xlsx"
        )
    except Exception as e:
        # Хэрэв OldBow бас биш бол алдааг дэлгэцэнд харуулна
        return f"Алдаа гарлаа: {str(e)}. Хүснэгтийн нэрээ шалгана уу."

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

@app.route('/reset-database-secure-99')
@login_required
def reset_db():
    if current_user.role == 'admin':
        try:
            # Дарааллаар нь устгах (Хамааралтай өгөгдлүүдээс эхэлнэ)
            Transaction.query.delete()
            Expense.query.delete()
            LaborFee.query.delete()
            Product.query.delete() # Хамгийн сүүлд бараагаа устгана
            
            db.session.commit()
            return "Мэдээллийн сан амжилттай цэвэрлэгдлээ. Одоо шинээр бараагаа оруулна уу."
        except Exception as e:
            db.session.rollback()
            return f"Алдаа гарлаа: {str(e)}"
    return "Эрх хүрэхгүй байна", 403
    
if __name__ == "__main__":
    with app.app_context():
        db.create_all()  # Энэ мөр нь байхгүй байгаа бүх хүснэгтийг (old_bow) үүсгэнэ
    app.run(debug=True)
