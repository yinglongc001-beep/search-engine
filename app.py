from __future__ import annotations

import io
import json
import math
import os
import re
import hashlib
import hmac
import sqlite3
import secrets
import threading
import zipfile
from contextlib import contextmanager
from functools import wraps
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from flask import Flask, Response, jsonify, redirect, render_template, request, send_file, session
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

BASE_DIR = Path(__file__).resolve().parent
LOCAL_ENV = BASE_DIR / ".env"
if LOCAL_ENV.exists():
    for line in LOCAL_ENV.read_text(encoding="utf-8").splitlines():
        if "=" in line and not line.strip().startswith("#"):
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip())

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("PUBLIC_DEPLOYMENT") == "1"
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=2)
limiter = Limiter(get_remote_address, app=app, default_limits=[], storage_uri="memory://")
DATABASE_PATH = Path(os.environ.get("DATABASE_PATH", BASE_DIR / "instance" / "users.sqlite3"))
DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
SHARED_DATA_PATH = Path(os.environ.get("SHARED_DATA_PATH", DATABASE_PATH.parent / "shared_reports.pkl"))
SHARED_DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
SHARED_DATA = {"digest": None, "total_port": None, "df_delivery": None}
DATA_LOCK = threading.Lock()

REQUIRED_FILES = [
    "Delfin Summary.xlsx", "POT MG - Daily Statistics.xlsx",
    "OrderEnquiryDownloadExcel.xlsx", "WarehouseStockDetailExcel.xlsx",
    "reg.xlsx", "Portbury MG - Daily Statistics.xlsx", "On-Site-Vehicles.xlsx",
    "VehicleMaterialInfoMaintenance.xlsx", "TYN haulier.xlsx",
    "NEW MG UK TRANSPORT SPLIT BY DEALER.xlsx", "FVL_Daily_KPI.csv",
    "MGExport.csv", "MgDailyReport.csv", "ECM_TRACK_Vehicle Tracking.xlsx",
]

def clean_vin(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def load_bundle(upload):
    if not upload or not upload.filename.lower().endswith(".zip"):
        raise ValueError("Choose a ZIP file containing the source reports.")
    files = {}
    with zipfile.ZipFile(upload) as archive:
        expanded_size = sum(item.file_size for item in archive.infolist())
        if expanded_size > 500 * 1024 * 1024:
            raise ValueError("The ZIP expands to more than 500 MB. Please use a smaller report set.")
        for name in archive.namelist():
            if not name.endswith("/"):
                files[name.replace("\\", "/").split("/")[-1].casefold()] = archive.read(name)
    missing = [name for name in REQUIRED_FILES if name.casefold() not in files]
    if missing:
        raise ValueError("ZIP is missing: " + ", ".join(missing))

    def excel(name, **kwargs):
        return pd.read_excel(io.BytesIO(files[name.casefold()]), **kwargs)
    def csv(name, **kwargs):
        return pd.read_csv(io.BytesIO(files[name.casefold()]), **kwargs)

    flx = excel("Delfin Summary.xlsx", sheet_name="MG Summary")
    damage = excel("Delfin Summary.xlsx", sheet_name="Damage")
    tyn = excel("POT MG - Daily Statistics.xlsx", sheet_name="Stock List",
                usecols=["VIN", "Arrival Date", "Scheduled collection date", "Haulier", "Gate out date", "Hold", "Remarks"])
    tyn = tyn[tyn["Gate out date"].isna()].copy()
    order = excel("OrderEnquiryDownloadExcel.xlsx", usecols=["VIN", "Dealer Name", "Dealer Code", "Region", "Delivery Destination", "Registration Date", "Tracking Status", "Vehicle Material Code", "Credit Hold"])
    stock = excel("WarehouseStockDetailExcel.xlsx", usecols=["vin", "vehicleMaterialCode", "vehicleStatus"])
    reg = excel("reg.xlsx", usecols=["VIN", "Dealer Name", "Dealer Code", "Region", "Delivery Destination", "Registration Date", "Tracking Status", "Vehicle Material Code", "Credit Hold"])
    bca = excel("Portbury MG - Daily Statistics.xlsx", sheet_name="Stock List_2", skiprows=4, usecols=["Contract", "Chassis Number", "Status", "Compound", "Days on Site"])
    lbh = excel("On-Site-Vehicles.xlsx", usecols=["VIN"])
    vmc = excel("VehicleMaterialInfoMaintenance.xlsx")
    tyn_h = excel("TYN haulier.xlsx")
    normal_h = excel("NEW MG UK TRANSPORT SPLIT BY DEALER.xlsx", usecols=["Dealer_Code", "Haulier"])

    fvl = csv("FVL_Daily_KPI.csv", usecols=["Notification_Date", "VIN", "Delivery_Code", "Collected", "Delivery_Date"])
    anji = csv("MGExport.csv", encoding="windows-1252")
    anji = anji.rename(columns={"Notification Date":"Notification_Date", "Delivery Code":"Delivery_Code", "Collection Date":"Collected", "Delivery Date":"Delivery_Date"})[["VIN", "Notification_Date", "Delivery_Code", "Collected", "Delivery_Date"]]
    ac = csv("MgDailyReport.csv")
    ac = ac.rename(columns={" Vin":"VIN", " Delivery_Code":"Delivery_Code", "Delivered":"Delivery_Date"})[["VIN", "Notification_Date", "Delivery_Code", "Collected", "Delivery_Date"]]
    ecm = excel("ECM_TRACK_Vehicle Tracking.xlsx", usecols="B,E,J:M")
    ecm = ecm.rename(columns={"Notification Date":"Notification_Date", "Consignee Dealer Code":"Delivery_Code", "Despatch Date":"Collected", "Delivery Date":"Delivery_Date"})[["VIN", "Notification_Date", "Delivery_Code", "Collected", "Delivery_Date"]]
    deliveries = []
    for frame, identity in [(fvl,"FVL"),(ecm,"ECM"),(ac,"AC"),(anji,"Anji")]:
        frame = frame.copy()
        frame["VIN"] = frame["VIN"].map(clean_vin)
        frame["identity"] = identity
        frame["Notification_Date"] = pd.to_datetime(frame["Notification_Date"], errors="coerce", dayfirst=True)
        for col in ["Collected", "Delivery_Date"]:
            frame[col] = frame[col].replace("0001-01-01", pd.NA)
            frame[col] = pd.to_datetime(frame[col], errors="coerce", dayfirst=True)
        deliveries.append(frame)
    delivery = pd.concat(deliveries, ignore_index=True).sort_values("Notification_Date", ascending=False).drop_duplicates("VIN", keep="first")

    stock = stock.rename(columns={"vin":"VIN", "vehicleMaterialCode":"Vehicle Material Code", "vehicleStatus":"Tracking Status"})
    stock_info = pd.concat([order, reg, stock[["VIN", "Tracking Status", "Vehicle Material Code"]]], ignore_index=True)
    stock_info["VIN"] = stock_info["VIN"].map(clean_vin)
    stock_info = stock_info.merge(vmc[["Vehicle Material Code", "Series", "Model Year", "Gear", "Trim Version", "Interior", "Paint&Colour"]], how="left", on="Vehicle Material Code")

    d = damage[["Vin", "Notes"]].rename(columns={"Vin":"VIN", "Notes":"Status"}).copy()
    d["VIN"] = d["VIN"].map(lambda x: clean_vin(x)[:17])
    d = d.drop_duplicates("VIN")
    ips = flx[["VIN"]].copy()
    ips["VIN"] = ips["VIN"].map(clean_vin)
    ips = ips.merge(d, how="left", on="VIN")
    ips["port"] = "IPS"
    bca = bca.rename(columns={"Chassis Number":"VIN"})[["VIN", "Status"]].copy()
    bca["port"] = "BCA"
    ty = tyn[["VIN"]].copy(); ty["port"] = "TYN"
    lb = lbh[["VIN"]].copy(); lb["port"] = "LBH"
    ports = pd.concat([bca, ips, ty, lb], ignore_index=True)
    ports["VIN"] = ports["VIN"].map(clean_vin)
    ports = ports.merge(stock_info, how="left", on="VIN")
    normal = dict(zip(normal_h["Dealer_Code"].astype(str).str.strip(), normal_h["Haulier"]))
    tyn_dict = dict(zip(tyn_h["delivery code"].astype(str).str.strip(), tyn_h["haulier"]))
    def haulier(row):
        destination = str(row.get("Delivery Destination", ""))[:6].strip()
        return tyn_dict.get(destination) if row["port"] == "TYN" else normal.get(destination)
    ports["haulier"] = ports.apply(haulier, axis=1)
    ports.loc[(ports["port"] == "IPS") & (ports["haulier"] == "ANJI"), "haulier"] = "FVL"
    ports = ports.merge(delivery, how="left", on="VIN")
    return ports, delivery

def save_shared_reports(digest, total_port, df_delivery):
    temporary = SHARED_DATA_PATH.with_name(SHARED_DATA_PATH.name + ".tmp")
    with temporary.open("wb") as output:
        pickle.dump({"digest": digest, "total_port": total_port, "df_delivery": df_delivery}, output, protocol=pickle.HIGHEST_PROTOCOL)
    os.replace(temporary, SHARED_DATA_PATH)

def load_saved_shared_reports():
    if not SHARED_DATA_PATH.exists():
        return
    try:
        with SHARED_DATA_PATH.open("rb") as source:
            data = pickle.load(source)
        if not all(key in data for key in ("digest", "total_port", "df_delivery")):
            raise ValueError("Saved report data is incomplete")
        SHARED_DATA.update(data)
        app.logger.info("Loaded the shared report data from persistent storage")
    except Exception:
        app.logger.exception("Could not load the saved shared report data")

def records(frame):
    # Cast before replacing nulls: otherwise pandas keeps float columns as
    # float64 and silently turns None back into NaN, which is invalid JSON.
    out = frame.astype(object).where(pd.notna(frame), None)
    for col in out.columns:
        out[col] = out[col].map(lambda v: v.isoformat() if isinstance(v, (pd.Timestamp, datetime, date)) else (v.item() if hasattr(v, "item") else v))
    return {"columns": list(out.columns), "rows": out.values.tolist()}

def json_safe(value):
    """Convert pandas/numpy nulls and scalars into strict JSON values."""
    if value is None or value is pd.NA or value is pd.NaT:
        return None
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(key): json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_safe(item) for item in value]
    if hasattr(value, "item"):
        return json_safe(value.item())
    if isinstance(value, float) and not math.isfinite(value):
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return value

def strict_json_response(payload):
    return Response(json.dumps(json_safe(payload), allow_nan=False), mimetype="application/json")

@contextmanager
def db_connect():
    connection = sqlite3.connect(DATABASE_PATH)
    connection.row_factory = sqlite3.Row
    try:
        yield connection
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()

def hash_password(password, salt=None):
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), 600_000).hex()
    return f"pbkdf2_sha256$600000${salt}${digest}"

def verify_password(password, encoded):
    try:
        algorithm, rounds, salt, expected = encoded.split("$", 3)
        actual = hashlib.pbkdf2_hmac(algorithm.removeprefix("pbkdf2_"), password.encode("utf-8"), bytes.fromhex(salt), int(rounds)).hex()
        return hmac.compare_digest(actual, expected)
    except (ValueError, TypeError):
        return False

def initialize_database():
    with db_connect() as connection:
        connection.execute("""CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL COLLATE NOCASE UNIQUE,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""")
        admin_name = os.environ.get("ADMIN_USERNAME", "leo.chen").strip()
        admin_hash = os.environ.get("ADMIN_PASSWORD_HASH")
        if not admin_hash and os.environ.get("ADMIN_PASSWORD"):
            admin_hash = hash_password(os.environ["ADMIN_PASSWORD"])
        if admin_name and admin_hash:
            connection.execute("INSERT OR IGNORE INTO users(username,password_hash,role) VALUES(?,?, 'admin')", (admin_name, admin_hash))

initialize_database()
load_saved_shared_reports()

@app.before_request
def protect_forms_and_api_posts():
    if request.method in {"GET", "HEAD"}:
        session.setdefault("csrf_token", secrets.token_urlsafe(32))
        return None
    expected = session.get("csrf_token", "")
    supplied = request.form.get("csrf_token", "") or request.headers.get("X-CSRF-Token", "")
    if not expected or not hmac.compare_digest(expected, supplied):
        if request.path.startswith("/api/"):
            return strict_json_response({"error": "Page security token expired. Refresh the page and try again."}), 400
        return "Page security token expired. Refresh the page and try again.", 400

@app.context_processor
def inject_csrf_token():
    return {"csrf_token": session.get("csrf_token", "")}

def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            if request.path.startswith("/api/"):
                return strict_json_response({"error": "Please sign in to use the VIN search."}), 401
            return render_template("login.html", error="Please sign in to continue."), 401
        with db_connect() as connection:
            user = connection.execute("SELECT id, username, role, active FROM users WHERE id=?", (session["user_id"],)).fetchone()
        if not user or not user["active"]:
            session.clear()
            return render_template("login.html", error="This account is unavailable. Contact the administrator."), 401
        request.current_user = user
        return view(*args, **kwargs)
    return wrapped

def admin_required(view):
    @wraps(view)
    @login_required
    def wrapped(*args, **kwargs):
        if request.current_user["role"] != "admin":
            return "Administrator access required.", 403
        return view(*args, **kwargs)
    return wrapped

@app.get("/")
@login_required
def index():
    return render_template("index.html", files=REQUIRED_FILES)

@app.route("/healthz", methods=["GET", "HEAD"])
def healthz():
    return "ok", 200

@app.route("/login", methods=["GET", "POST"])
@limiter.limit("10 per hour", methods=["POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        with db_connect() as connection:
            user = connection.execute("SELECT * FROM users WHERE username=? COLLATE NOCASE", (username,)).fetchone()
        if not user or not user["active"] or not verify_password(password, user["password_hash"]):
            return render_template("login.html", error="Username or password is incorrect."), 401
        session.clear()
        session["user_id"] = user["id"]
        session["csrf_token"] = secrets.token_urlsafe(32)
        session.permanent = True
        return redirect("/")
    return render_template("login.html")

@app.route("/register", methods=["GET", "POST"])
@limiter.limit("5 per hour", methods=["POST"])
def register():
    if request.method == "POST":
        with db_connect() as connection:
            has_admin = connection.execute("SELECT 1 FROM users WHERE role='admin' AND active=1 LIMIT 1").fetchone()
        if not has_admin:
            return render_template("register.html", error="Registration is not available until the administrator finishes setup."), 503
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if not re.fullmatch(r"[A-Za-z0-9._-]{3,40}", username):
            return render_template("register.html", error="Use 3–40 letters, numbers, dots, underscores, or hyphens."), 400
        if len(password) < 10:
            return render_template("register.html", error="Choose a password with at least 10 characters."), 400
        try:
            with db_connect() as connection:
                connection.execute("INSERT INTO users(username,password_hash) VALUES(?,?)", (username, hash_password(password)))
        except sqlite3.IntegrityError:
            return render_template("register.html", error="That username is already registered."), 400
        return render_template("login.html", message="Account created. Sign in to continue.")
    return render_template("register.html")

@app.post("/logout")
@login_required
def logout():
    session.clear()
    return redirect("/login")

@app.get("/admin/users")
@admin_required
def admin_users():
    with db_connect() as connection:
        users = connection.execute("SELECT id,username,role,active,created_at FROM users ORDER BY username COLLATE NOCASE").fetchall()
    return render_template("admin_users.html", users=users)

@app.post("/admin/users")
@admin_required
def manage_user():
    action = request.form.get("action")
    user_id = request.form.get("user_id", type=int)
    current_id = request.current_user["id"]
    with db_connect() as connection:
        if action == "create":
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            role = request.form.get("role", "user")
            if not re.fullmatch(r"[A-Za-z0-9._-]{3,40}", username) or len(password) < 10 or role not in {"user", "admin"}:
                return "Enter a valid username, a password of at least 10 characters, and a valid role.", 400
            try:
                connection.execute("INSERT INTO users(username,password_hash,role) VALUES(?,?,?)", (username, hash_password(password), role))
            except sqlite3.IntegrityError:
                return "That username is already registered.", 400
        elif action == "reset" and user_id:
            password = request.form.get("password", "")
            if len(password) < 10:
                return "New passwords must be at least 10 characters.", 400
            connection.execute("UPDATE users SET password_hash=? WHERE id=?", (hash_password(password), user_id))
        elif action in {"deactivate", "activate", "make_admin", "make_user", "delete"} and user_id:
            target = connection.execute("SELECT id,role,active FROM users WHERE id=?", (user_id,)).fetchone()
            if not target:
                return "User not found.", 404
            removes_admin = target["role"] == "admin" and action in {"deactivate", "make_user", "delete"} and target["active"]
            if removes_admin:
                admins = connection.execute("SELECT COUNT(*) FROM users WHERE role='admin' AND active=1").fetchone()[0]
                if admins <= 1:
                    return "The last active administrator cannot be removed or deactivated.", 400
            if user_id == current_id and action in {"deactivate", "make_user", "delete"}:
                return "You cannot deactivate or remove your own administrator account.", 400
            if action == "deactivate": connection.execute("UPDATE users SET active=0 WHERE id=?", (user_id,))
            elif action == "activate": connection.execute("UPDATE users SET active=1 WHERE id=?", (user_id,))
            elif action == "make_admin": connection.execute("UPDATE users SET role='admin' WHERE id=?", (user_id,))
            elif action == "make_user": connection.execute("UPDATE users SET role='user' WHERE id=?", (user_id,))
            elif action == "delete": connection.execute("DELETE FROM users WHERE id=?", (user_id,))
        else:
            return "Unknown account action.", 400
    return redirect("/admin/users")

@app.get("/api/status")
@login_required
def shared_status():
    with DATA_LOCK:
        ready = SHARED_DATA["total_port"] is not None and SHARED_DATA["df_delivery"] is not None
    return strict_json_response({"ready": ready, "message": "Shared reports are ready." if ready else "The administrator needs to upload the reports first."})

@app.post("/api/upload")
@limiter.limit("4 per hour")
@admin_required
def upload_reports():
    try:
        bundle = request.files.get("bundle")
        if not bundle:
            raise ValueError("Choose a ZIP file containing the source reports.")
        content = bundle.read()
        digest = hashlib.sha256(content).hexdigest()
        with DATA_LOCK:
            if SHARED_DATA["digest"] == digest and SHARED_DATA["total_port"] is not None:
                return strict_json_response({"message": "These shared reports are already loaded.", "cached": True})
        bundle.stream = io.BytesIO(content)
        total_port, df_delivery = load_bundle(bundle)
        with DATA_LOCK:
            save_shared_reports(digest, total_port, df_delivery)
            SHARED_DATA.update({"digest": digest, "total_port": total_port, "df_delivery": df_delivery})
        return strict_json_response({"message": "Shared reports saved. All users can now search this data.", "cached": False})
    except (ValueError, zipfile.BadZipFile, KeyError, OSError) as exc:
        return strict_json_response({"error": str(exc)}), 400
    except Exception as exc:
        app.logger.exception("Report loading failed")
        return strict_json_response({"error": "Could not read the reports. Check the workbook sheets and column names. Details: " + str(exc)}), 400

@app.post("/api/search")
@limiter.limit("120 per minute")
@login_required
def search():
    try:
        raw = request.form.get("vins", "")
        vins = list(dict.fromkeys(v for v in re.split(r"[\s,;]+", raw.strip()) if v))
        if not vins:
            return strict_json_response({"error": "Enter at least one VIN."}), 400
        with DATA_LOCK:
            total_port = SHARED_DATA["total_port"]
            df_delivery = SHARED_DATA["df_delivery"]
        if total_port is None or df_delivery is None:
            return strict_json_response({"error": "The administrator needs to upload the reports before searching."}), 400
        port_rows = total_port[total_port["VIN"].isin(vins)]
        missing = [v for v in vins if v not in set(port_rows["VIN"])]
        delivered = df_delivery[df_delivery["VIN"].isin(missing)]
        found_delivered = set(delivered["VIN"])
        not_found = [v for v in missing if v not in found_delivered]
        return strict_json_response({"port": records(port_rows), "delivered": records(delivered), "missing":{"columns":["VIN"],"rows":[[v] for v in not_found]}, "requested":len(vins)})
    except (ValueError, zipfile.BadZipFile, KeyError, OSError) as exc:
        return strict_json_response({"error": str(exc)}), 400
    except Exception as exc:
        app.logger.exception("Search failed")
        return strict_json_response({"error": "Could not search the reports. Details: " + str(exc)}), 400

@app.post("/api/export")
@login_required
def export():
    payload = request.get_json(force=True)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for key, title in [("port", "port information"), ("delivered", "delivered units"), ("missing", "not found units")]:
            section = payload.get(key, {"columns": [], "rows": []})
            pd.DataFrame(section["rows"], columns=section["columns"]).to_excel(writer, sheet_name=title, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="VIN search results.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=os.environ.get("FLASK_DEBUG") == "1")
