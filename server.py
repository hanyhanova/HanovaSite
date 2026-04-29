import os
import re
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import psycopg2
from psycopg2.extras import RealDictCursor
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__, static_folder="site", static_url_path="")
CORS(app, origins=os.environ.get("ALLOWED_ORIGINS", "*"))

DATABASE_URL = os.environ.get("DATABASE_URL")


def get_db():
    conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
    return conn


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS consultations (
            id          SERIAL PRIMARY KEY,
            name        TEXT NOT NULL,
            email       TEXT NOT NULL,
            title       TEXT,
            phone       TEXT,
            company     TEXT NOT NULL,
            sector      TEXT,
            interest    TEXT,
            message     TEXT NOT NULL,
            submitted_at TIMESTAMPTZ DEFAULT NOW()
        )
    """)
    conn.commit()
    cur.close()
    conn.close()


init_db()

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


@app.route("/api/contact", methods=["POST"])
def contact():
    data = request.get_json(silent=True) or {}

    name    = str(data.get("name", "")).strip()
    email   = str(data.get("email", "")).strip()
    company = str(data.get("company", "")).strip()
    message = str(data.get("message", "")).strip()

    # Validate required fields
    if not name or not email or not company or not message:
        return jsonify({"error": "name, email, company and message are required."}), 400

    if not EMAIL_RE.match(email):
        return jsonify({"error": "Invalid email address."}), 400

    # Sanitise optional fields (plain text only, max lengths)
    title    = str(data.get("title",    "")).strip()[:120]
    phone    = str(data.get("phone",    "")).strip()[:40]
    sector   = str(data.get("sector",   "")).strip()[:120]
    interest = str(data.get("interest", "")).strip()[:120]

    try:
        conn = get_db()
        cur  = conn.cursor()
        cur.execute(
            """
            INSERT INTO consultations
                (name, email, title, phone, company, sector, interest, message)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id, submitted_at
            """,
            (name, email, title, phone, company, sector, interest, message),
        )
        row = cur.fetchone()
        conn.commit()
        cur.close()
        conn.close()
    except Exception as exc:
        app.logger.error("DB error: %s", exc)
        return jsonify({"error": "Could not save submission. Please try again later."}), 500

    return jsonify({"ok": True, "id": row["id"]}), 201


# Serve static site pages
@app.route("/", defaults={"path": "index.html"})
@app.route("/<path:path>")
def static_files(path):
    return send_from_directory(app.static_folder, path)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
