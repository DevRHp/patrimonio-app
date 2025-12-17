from pymongo import MongoClient
import sys

try:
    client = MongoClient('mongodb://localhost:27017/', serverSelectionTimeoutMS=2000)
    client.admin.command('ping')
    print("SUCCESS: MongoDB is running and accessible.")
    
    # List dbs
    print("Databases:", client.list_database_names())
    
except Exception as e:
    print(f"FAILURE: Could not connect to MongoDB. Error: {e}")
    print("Please ensure MongoDB Community Server is installed and the service is running.")
