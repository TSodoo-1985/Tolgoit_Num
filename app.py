import os
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
    __tablename__ = 'expense'
    __table_args__ = {'extend_existing': True}  # Энэ мөр давхардал алдаанаас хамгаална
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(50), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    description = db.Column(db.Text)
    date = db.Column(db.DateTime, default=datetime.utcnow)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    user = db.relationship('User', backref='expense_records')

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

@app.route('/inventory', methods=['GET', 'POST'])
@login_required
def inventory():
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

@app.route('/expenses', methods=['GET', 'POST'])
@login_required
def expenses():
    if request.method == 'POST':
        category = request.form.get('category')
        amount = float(request.form.get('amount') or 0)
        description = request.form.get('description')
        
        new_expense = Expense(
            category=category,
            amount=amount,
            description=description,
            date=datetime.now(),
            user_id=current_user.id
        )
        db.session.add(new_expense)
        db.session.commit()
        flash('Зардал амжилттай бүртгэгдлээ!')
        return redirect(url_for('expenses'))

    # Зардлын жагсаалтыг харуулах
    expense_list = Expense.query.order_by(Expense.date.desc()).all()
    return render_template('expenses.html', expenses=expense_list)
    
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
    
# --- ТАЙЛАН, СТАТИСТИК ---

@app.route('/statistics')
@login_required
def statistics():
    # Огнооны шүүлтүүр авах
    start_date_str = request.args.get('start_date')
    end_date_str = request.args.get('end_date')
    
    if start_date_str and end_date_str:
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
        # Сонгосон хоногуудын жагсаалт
        delta = end_date - start_date
        dates = [(start_date + timedelta(days=i)).strftime('%Y-%m-%d') for i in range(delta.days + 1)]
    else:
        # Анхдагчаар сүүлийн 7 хоног
        today = datetime.now().date()
        dates = [(today - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(6, -1, -1)]

    sales_data = []
    profit_data = []
    expense_data = []
    
    for date_str in dates:
        # 1. Борлуулалт (Бөөний болон Жижиглэн зарлага)
        sales = db.session.query(db.func.sum(Transaction.quantity * db.case(
                (Transaction.type == 'Бөөний зарлага', Product.wholesale_price),
                (Transaction.type == 'Жижиглэн зарлага', Product.retail_price),
                else_=0
            ))).join(Product).filter(db.func.date(Transaction.timestamp) == date_str).scalar() or 0
        
        # 2. Өртөг (Зарсан барааны өртөг - Ашиг бодоход хэрэгтэй)
        cost_sum = db.session.query(db.func.sum(Transaction.quantity * Product.cost_price)).\
            join(Product).filter(Transaction.type.like('%зарлага%'), 
            db.func.date(Transaction.timestamp) == date_str).scalar() or 0
        
        # 3. Зардал (Ажлын хөлс ороогүй цэвэр зардал)
        daily_expense = db.session.query(db.func.sum(Expense.amount)).\
            filter(Expense.category != "Ажлын хөлс", db.func.date(Expense.date) == date_str).scalar() or 0

        sales_data.append(float(sales))
        expense_data.append(float(daily_expense))
        # Цэвэр ашиг = (Борлуулалт - Өртөг) - Зардал
        profit_data.append(float(sales - cost_sum - daily_expense))

    # 4. Топ 5 бараа (Ширхэгээр)
    top_products = db.session.query(Product.name, db.func.sum(Transaction.quantity)).\
        join(Transaction).filter(Transaction.type.like('%зарлага%')).\
        group_by(Product.name).order_by(db.func.sum(Transaction.quantity).desc()).limit(5).all()
    
    top_labels = [p[0] for p in top_products]
    top_values = [int(p[1]) for p in top_products]

    return render_template('statistics.html', 
                         dates=dates, 
                         sales=sales_data, 
                         profit=profit_data, 
                         expenses=expense_data,
                         top_labels=top_labels,
                         top_values=top_values,
                         start_date=start_date_str or "",
                         end_date=end_date_str or "")

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
    
    # Нийт дүн нэмэх (Орлогоос бусад тохиолдолд)
    if not df.empty and type != 'Орлого':
        totals = {
            'Огноо': 'НИЙТ ДҮН:', 'Барааны код': '', 'Барааны нэр': '', 'Гүйлгээний төрөл': '',
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
    data = [{"Огноо": t.timestamp, "Бараа": t.product.name, "Зөрүү": t.type, "Тоо": t.quantity} for t in transactions]
    
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
