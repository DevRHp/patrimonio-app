import os
import gridfs
from pymongo import MongoClient
from dotenv import load_dotenv

load_dotenv()

# Get URI from env or default to local
MONGO_URI = os.getenv('MONGO_URI', 'mongodb://localhost:27017/patrimonio')

client = None
db = None
fs = None

def init_db_connection():
    global client, db, fs
    if not client:
        try:
            client = MongoClient(MONGO_URI)
            # Default database name 'patrimonio' or from URI
            db = client.get_database() 
            fs = gridfs.GridFS(db)
            print(f" * Connected to MongoDB at {MONGO_URI}")
        except Exception as e:
            print(f" * Failed to connect to MongoDB: {e}")
            raise e

def get_db():
    if not db:
        init_db_connection()
    return db

def get_fs():
    if not fs:
        init_db_connection()
    return fs
