import os
import pandas as pd
from flask import Flask, jsonify, request
from sqlalchemy import (
    create_engine, Column, Integer, String, DECIMAL, Enum, ForeignKey,
    TIMESTAMP, text
)
from sqlalchemy.orm import declarative_base, sessionmaker, scoped_session, relationship
from datetime import datetime
from functools import lru_cache
from werkzeug.utils import secure_filename
from flask_cors import CORS   # <-- new import

# ─── Configuration ─────────────────────────────────────────────────────────────
DATABASE_URL = os.getenv(
    "DATABASE_URL"
)
UPLOAD_DIR = "/tmp/uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ─── Database Engine ──────────────────────────────────────────────────────────
engine = create_engine(
    DATABASE_URL,
    pool_size=5,
    max_overflow=5,
    pool_timeout=30,
    pool_recycle=300,  # recycle every 5 minutes
    pool_pre_ping=True,
    isolation_level="AUTOCOMMIT",
    future=True,
)

SessionLocal = sessionmaker(bind=engine, autocommit=False, autoflush=False)
Session = scoped_session(SessionLocal)
Base = declarative_base()

# ─── Models ───────────────────────────────────────────────────────────────────
class Employee(Base):
    __tablename__ = "employees"
    employee_id = Column(Integer, primary_key=True, autoincrement=True)
    emp_code = Column(String(64), unique=True, nullable=False)
    full_name = Column(String(255), nullable=False)
    status_name = Column(String(100), default="ปกติ")
    created_at = Column(TIMESTAMP, default=datetime.utcnow)
    items = relationship("SalaryItem", back_populates="employee")


class SalarySheet(Base):
    __tablename__ = "salary_sheets"
    sheet_id = Column(Integer, primary_key=True, autoincrement=True)
    month_year = Column(String(50), unique=True, nullable=False)
    created_at = Column(TIMESTAMP, default=datetime.utcnow)
    items = relationship("SalaryItem", back_populates="sheet")


class SalaryItem(Base):
    __tablename__ = "salary_items"
    item_id = Column(Integer, primary_key=True, autoincrement=True)
    sheet_id = Column(Integer, ForeignKey("salary_sheets.sheet_id"), nullable=False)
    employee_id = Column(Integer, ForeignKey("employees.employee_id"), nullable=False)
    item_group = Column(Enum("earnings", "deductions", "summary"), nullable=False)
    item_name = Column(String(255), nullable=False)
    amount = Column(DECIMAL(14, 2), default=0)
    sheet = relationship("SalarySheet", back_populates="items")
    employee = relationship("Employee", back_populates="items")


class SalaryItemMeta(Base):
    __tablename__ = "salary_item_meta"
    meta_id = Column(Integer, primary_key=True, autoincrement=True)
    item_name = Column(String(255), unique=True, nullable=False)
    item_group = Column(Enum("earnings", "deductions", "summary"), nullable=True)
    remark = Column(String(255))
    updated_at = Column(TIMESTAMP, default=datetime.utcnow, onupdate=datetime.utcnow)


Base.metadata.create_all(bind=engine)

# ─── Flask App ────────────────────────────────────────────────────────────────
app = Flask(__name__)
CORS(app,resources={r"/*":{"origins":"*"}})
@app.teardown_appcontext
def remove_session(exception=None):
    """Automatically clean up sessions to prevent sleeping connections."""
    Session.remove()

# ─── Cache Helper ─────────────────────────────────────────────────────────────
@lru_cache(maxsize=256)
def load_item_meta():
    """Return dict of {item_name: item_group} from DB."""
    session = Session()
    meta = session.execute(text("SELECT item_name, item_group FROM salary_item_meta")).fetchall()
    session.close()
    return {m[0]: m[1] for m in meta}


# ─── Health Check ─────────────────────────────────────────────────────────────
@app.route("/healthz")
def healthz():
    return jsonify({"status": "OK", "message": "Mena Payroll API is healthy ✅"}), 200


# ─────────────────────────────────────────────────────────────────────────────
# 1️⃣ GET & POST salary_data
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/salary_data/data", methods=["GET", "POST"])
def salary_data():
    session = Session()

    if request.method == "GET":
        month = request.args.get("month-year")
        emp_code = request.args.get("emp_id")
        if not month or not emp_code:
            return jsonify({"error": "month-year and emp_id required"}), 400

        sheet = session.query(SalarySheet).filter_by(month_year=month).first()
        emp = session.query(Employee).filter_by(emp_code=emp_code).first()
        if not sheet or not emp:
            session.close()
            return jsonify([])

        items = session.query(SalaryItem).filter_by(
            sheet_id=sheet.sheet_id, employee_id=emp.employee_id
        ).all()
        grouped = {"earnings": {}, "deductions": {}, "summary": {}}
        for i in items:
            grouped[i.item_group][i.item_name] = f"{float(i.amount):.2f}"

        result = [{
            "Sheet": sheet.month_year,
            "รหัสพนักงาน": emp.emp_code,
            "ชื่อ - นามสกุล": emp.full_name,
            "สถานะคนลาออก": emp.status_name,
            "datalist": grouped,
        }]
        session.close()
        return jsonify(result)

    # ────────────────────────────────
    # POST → insert or update (smart upsert)
    # ────────────────────────────────
    data = request.get_json(force=True)
    month = data.get("month-year")
    emp_code = data.get("emp_id")
    full_name = data.get("full_name", "")
    status = data.get("status", "ปกติ")
    datalist = data.get("datalist", {})

    if not month or not emp_code:
        session.close()
        return jsonify({"error": "month-year and emp_id required"}), 400

    try:
        # Ensure sheet exists
        sheet = session.query(SalarySheet).filter_by(month_year=month).first()
        if not sheet:
            sheet = SalarySheet(month_year=month)
            session.add(sheet)
            session.flush()

        # Upsert employee (safe)
        session.execute(text("""
            INSERT INTO employees (emp_code, full_name, status_name, created_at)
            VALUES (:code, :name, :status, NOW())
            ON DUPLICATE KEY UPDATE full_name=:name, status_name=:status
        """), {"code": emp_code, "name": full_name, "status": status})

        emp = session.query(Employee).filter_by(emp_code=emp_code).first()

        # Delete old salary items
        session.query(SalaryItem).filter_by(sheet_id=sheet.sheet_id, employee_id=emp.employee_id).delete()

        # Reload meta map
        meta_map = {
            row.item_name: row.item_group
            for row in session.query(SalaryItemMeta.item_name, SalaryItemMeta.item_group)
        }

        # Insert updated salary items
        for group, items in datalist.items():
            for name, val in items.items():
                try:
                    amount = float(val)
                except (ValueError, TypeError):
                    continue

                g = meta_map.get(name, group)
                session.add(SalaryItem(
                    sheet_id=sheet.sheet_id,
                    employee_id=emp.employee_id,
                    item_group=g,
                    item_name=name,
                    amount=amount
                ))

        session.commit()
        return jsonify({
            "status": "updated",
            "emp_id": emp_code,
            "month": month
        }), 201

    except Exception as e:
        session.rollback()
        return jsonify({"error": f"DB error: {str(e)}"}), 500

    finally:
        session.close()
        load_item_meta.cache_clear()



# ─────────────────────────────────────────────────────────────────────────────
# 2️⃣ Upload Excel endpoint (Thai→English month + batch commit)
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/upload_excel", methods=["POST"])
def upload_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Empty filename"}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_DIR, filename)
    file.save(filepath)

    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        return jsonify({"error": f"Failed to read Excel: {e}"}), 400

    # ─────────────────────────────────────────────
    # 🧹 Clean columns and fix naming
    # ─────────────────────────────────────────────
    df.columns = df.columns.astype(str).str.strip()
    df.columns = [c if not c.startswith("_") else f"Unnamed{c}" for c in df.columns]
    df = df.dropna(axis=1, how="all")  # remove empty columns

    # ─────────────────────────────────────────────
    # 🗓 Convert Thai month (e.g. พ.ย.2568 → November2025)
    # ─────────────────────────────────────────────
    prefix_map = {
        "ม.ค.": "January", "ก.พ.": "February", "มี.ค.": "March", "เม.ย.": "April",
        "พ.ค.": "May", "มิ.ย.": "June", "ก.ค.": "July", "ส.ค.": "August",
        "ก.ย.": "September", "ต.ค.": "October", "พ.ย.": "November", "ธ.ค.": "December"
    }

    if "Sheet" in df.columns:
        s = df["Sheet"].astype(str).str.replace(r"\s+", "", regex=True)
        df[["prefix", "year_th"]] = s.str.extract(r"^(\D+)(\d{4})$")
        df["Sheet"] = df["prefix"].map(prefix_map).fillna(df["prefix"]) + (
            (df["year_th"].astype(float)).astype(int).astype(str)
        )

    month_value = str(df.iloc[0].get("Sheet", "Unknown")).strip()

    # ─────────────────────────────────────────────
    # 💾 Start DB session
    # ─────────────────────────────────────────────
    session = Session()
    inserted_rows = 0

    try:
        # Ensure sheet record exists
        sheet = session.query(SalarySheet).filter_by(month_year=month_value).first()
        if not sheet:
            sheet = SalarySheet(month_year=month_value)
            session.add(sheet)
            session.flush()

        # Load meta mapping once
        meta_map = {
            row.item_name: row.item_group
            for row in session.query(SalaryItemMeta.item_name, SalaryItemMeta.item_group)
        }

        TOP_LEVEL = ["Sheet", "รหัสพนักงาน", "ชื่อ-นามสกุล", "สถานะคนลาออก", "prefix", "year_th"]

        # Iterate employees
        for _, row in df.iterrows():
            emp_code = str(row.get("รหัสพนักงาน", "")).strip()
            full_name = str(row.get("ชื่อ-นามสกุล", "")).strip()
            status = str(row.get("สถานะคนลาออก", "ปกติ")).strip()

            if not emp_code or emp_code.lower() in ["nan", "none"]:
                continue

            # ✅ Upsert employee (safe & non-locking)
            session.execute(text("""
                INSERT INTO employees (emp_code, full_name, status_name, created_at)
                VALUES (:code, :name, :status, NOW())
                ON DUPLICATE KEY UPDATE full_name=:name, status_name=:status
            """), {"code": emp_code, "name": full_name, "status": status})

            emp = session.query(Employee).filter_by(emp_code=emp_code).first()

            # Clear existing salary items
            session.query(SalaryItem).filter_by(
                sheet_id=sheet.sheet_id, employee_id=emp.employee_id
            ).delete()

            # Insert salary items
            for col in df.columns:
                if col in TOP_LEVEL:
                    continue

                val = row.get(col)
                if pd.isna(val):
                    continue
                try:
                    amount = float(val)
                except Exception:
                    continue

                group = meta_map.get(col, "earnings")  # default if missing
                session.add(SalaryItem(
                    sheet_id=sheet.sheet_id,
                    employee_id=emp.employee_id,
                    item_group=group,
                    item_name=col,
                    amount=amount
                ))

            inserted_rows += 1

        session.commit()

    except Exception as e:
        session.rollback()
        return jsonify({"error": f"DB error: {str(e)}"}), 500

    finally:
        session.close()
        load_item_meta.cache_clear()

    return jsonify({
        "status": "success",
        "sheet": month_value,
        "rows_inserted": inserted_rows
    }), 201

# ─────────────────────────────────────────────────────────────────────────────
# 3️⃣ salary_items/meta CRUD
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/salary_items/meta", methods=["GET", "POST", "DELETE"])
def salary_item_meta():
    session = Session()

    if request.method == "GET":
        rows = session.execute(
            text("SELECT meta_id, item_name, item_group, remark, updated_at FROM salary_item_meta ORDER BY item_name ASC")
        ).fetchall()
        session.close()
        return jsonify([
            {
                "meta_id": r[0],
                "item_name": r[1],
                "item_group": r[2],
                "remark": r[3],
                "updated_at": r[4].strftime("%Y-%m-%d %H:%M:%S"),
            } for r in rows
        ])

    if request.method == "POST":
        data = request.get_json(force=True)
        name = data.get("item_name")
        group = data.get("item_group")
        remark = data.get("remark", "")
        if not name or group not in ["earnings", "deductions", "summary"]:
            session.close()
            return jsonify({"error": "Invalid payload"}), 400

        session.execute(
            text("""
                INSERT INTO salary_item_meta (item_name, item_group, remark)
                VALUES (:name, :group, :remark)
                ON DUPLICATE KEY UPDATE item_group=:group, remark=:remark
            """),
            {"name": name, "group": group, "remark": remark},
        )
        session.commit()
        session.close()
        load_item_meta.cache_clear()
        return jsonify({"status": "updated", "item_name": name, "item_group": group}), 201

    if request.method == "DELETE":
        data = request.get_json(force=True)
        name = data.get("item_name")
        if not name:
            session.close()
            return jsonify({"error": "item_name required"}), 400

        result = session.execute(
            text("DELETE FROM salary_item_meta WHERE item_name=:name"), {"name": name}
        )
        session.commit()
        session.close()
        load_item_meta.cache_clear()

        if result.rowcount == 0:
            return jsonify({"status": "not_found", "item_name": name}), 404
        return jsonify({"status": "deleted", "item_name": name}), 200


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")
