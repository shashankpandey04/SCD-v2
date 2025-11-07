"""
Flask app: export `users` collection from MongoDB to an Excel spreadsheet and return it at /spreadsheet

Requirements:
  pip install flask pymongo python-dotenv pandas openpyxl

Example .env file (place next to this script):

MONGO_URI="mongodb+srv://username:password@cluster0.xxxxx.mongodb.net/?retryWrites=true&w=majority"
DB_NAME="mydatabase"
COLLECTION_NAME="users"

Run:
  export FLASK_APP=flask_mongo_spreadsheet.py
  flask run --host=0.0.0.0 --port=5000

Then open: http://localhost:5000/spreadsheet

Notes:
- This code projects only the requested keys: email, fullname, whatsapp, registration
- registration is converted to ISO-formatted string if it's a datetime; if missing, it becomes empty
- Be careful not to commit your .env to source control
"""

from flask import Flask, send_file, jsonify
from pymongo import MongoClient
from dotenv import load_dotenv
import os
import pandas as pd
from io import BytesIO
from datetime import datetime

load_dotenv()

MONGO_URI = os.getenv("MONGO_URI")
DB_NAME = os.getenv("DB_NAME", "test")
COLLECTION_NAME = os.getenv("COLLECTION_NAME", "users")

if not MONGO_URI:
    raise RuntimeError("MONGO_URI not set in environment (.env)")

# Connect to MongoDB (AWS-hosted MongoDB Atlas or self-managed URI)
client = MongoClient(MONGO_URI)
db = client[DB_NAME]
collection = db[COLLECTION_NAME]

app = Flask(__name__)

@app.route("/", methods=["GET"])
def index():
    return jsonify({"message": "Flask MongoDB -> Spreadsheet service. GET /spreadsheet to download."})

@app.route("/spreadsheet", methods=["GET"])
def spreadsheet():
    """Query the users collection, build a DataFrame with the requested keys, and return an Excel file."""
    # projection to only fetch necessary fields (and exclude _id)
    projection = {"email": 1, "fullname": 1, "whatsapp": 1, "registration": 1, "_id": 0}

    cursor = collection.find(filter={}, projection=projection)

    rows = []
    for doc in cursor:
        # Normalize each field; convert registration to ISO string if needed
        email = doc.get("email", "")
        fullname = doc.get("fullname", "")
        whatsapp = doc.get("whatsapp", "")
        reg = doc.get("registration", "")

        # If registration is a datetime-like object, convert to ISO string
        if isinstance(reg, datetime):
            registration = reg.isoformat()
        else:
            registration = str(reg) if reg not in (None, "") else ""

        rows.append({
            "email": email,
            "fullname": fullname,
            "whatsapp": whatsapp,
            "registration": registration,
        })

    # Create DataFrame and ensure column order
    df = pd.DataFrame(rows, columns=["email", "fullname", "whatsapp", "registration"])

    # Convert DataFrame to Excel in-memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="users")
    output.seek(0)

    # Send as a downloadable attachment
    return send_file(
        output,
        as_attachment=True,
        download_name="users.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if __name__ == "__main__":
    # For local testing only. In production, run with a WSGI server.
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)))
