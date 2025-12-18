from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

# Initialize SQLAlchemy
db = SQLAlchemy()

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    city = db.Column(db.String(100), nullable=False)
    is_admin = db.Column(db.Boolean, default=False)
    # Relationship to Networks (One admin can own many networks)
    networks = db.relationship('Network', backref='admin', lazy=True)

class Network(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    city = db.Column(db.String(100), nullable=False)
    admin_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)

class FileMetadata(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    filepath = db.Column(db.String(500), nullable=False) # Local path relative to upload folder
    type = db.Column(db.String(50), nullable=False) # 'master_spreadsheet' or 'audit_report'
    
    upload_date = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Ownership (Flexible: can belong to User OR Network OR Both)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    network_id = db.Column(db.Integer, db.ForeignKey('network.id'), nullable=True)

    def to_dict(self):
        return {
            'id': self.id,
            'filename': self.filename,
            'size': 0, # Placeholder, can use os.stat if needed
            'created_at': self.upload_date.isoformat(),
            'network_name': self.network_id # Will need join query to get name
        }
