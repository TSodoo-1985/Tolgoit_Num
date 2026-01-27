from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin

db = SQLAlchemy()

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password = db.Column(db.String(100), nullable=False)
    role = db.Column(db.String(20), default='staff') # admin, manager, staff

class Product(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    sku = db.Column(db.String(50), unique=True)
    stock = db.Column(db.Float, default=0.0)
    price = db.Column(db.Float)

class Transaction(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'))
    type = db.Column(db.String(20)) # 'Income', 'Wholesale', 'Retail', 'Audit'
    quantity = db.Column(db.Float)
    date = db.Column(db.DateTime, default=db.func.current_timestamp())
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))

class LaborFee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    description = db.Column(db.String(200), nullable=False) # Хийсэн ажил
    amount = db.Column(db.Float, nullable=False)           # Хөлс
    staff_name = db.Column(db.String(100))                 # Ажилтан
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)
