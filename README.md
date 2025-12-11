# 🧾 Mena Payroll API v2

A robust and scalable **Flask + SQLAlchemy + Pandas** backend for managing employee payroll data.  
It supports Excel uploads, auto-classifies salary components, and provides RESTful endpoints for both admin and system integrations.

---

## 🚀 Features

- **Excel Upload** – Upload payroll sheets directly to `/upload_excel` and automatically insert salary data into MySQL.  
- **Dynamic Metadata Mapping** – Uses `salary_item_meta` table for classifying columns (`earnings`, `deductions`, `summary`).  
- **Thai Month → English Conversion** – Converts month labels like `พ.ย.2568` → `November2025`.  
- **Smart Salary CRUD** – GET or POST to `/salary_data/data` for fetching or updating employee salary details.  
- **Auto Transaction Management** – Handles concurrent uploads safely with SQLAlchemy session pooling.  
- **Zero Hardcoded Logic** – All classification controlled by database metadata.  

---

## 🧱 Tech Stack

| Layer | Technology |
|-------|-------------|
| **Framework** | Flask 3.0 |
| **ORM** | SQLAlchemy 2.0 |
| **Database** | MySQL 8.x (PyMySQL) |
| **Excel Parser** | Pandas + OpenPyXL |
| **Language** | Python 3.12 |
| **Deployment** | Render / DigitalOcean / Docker ready |

---

## 🗂 Project Structure

```
api-payslip_v2/
│
├── app.py              # Main Flask app
├── requirements.txt    # Dependencies
├── uploads/            # Uploaded Excel files
├── README.md
└── .env                # Environment variables
```

---

## ⚙️ Environment Variables

Create a `.env` file in the project root with your database URL:

```bash
DATABASE_URL=mysql+pymysql://user:password@host:3306/be_database
```


---

## 📤 API Endpoints

### 🔹 Upload Payroll Excel

**POST** `/upload_excel`

| Field | Type | Required | Description |
|-------|------|-----------|-------------|
| file | File (.xlsx) | ✅ | Payroll Excel file |

**Response**
```json
{
  "status": "success",
  "sheet": "November2025",
  "rows_inserted": 125
}
```

---

### 🔹 Get / Update Salary Data

**GET**
```
/salary_data/data?month-year=November2025&emp_id=512052
```

**POST**
```
/salary_data/data
```

**Example Body**
```json
{
  "month-year": "November2025",
  "emp_id": "512052",
  "full_name": "สุที ปัชชาเขียว",
  "status": "ปกติ",
  "datalist": {
    "earnings": {
      "เงินเดือน": "4000.00",
      "ค่าเที่ยว": "15285.00"
    },
    "deductions": {
      "ประกันสังคม": "750.00"
    },
    "summary": {
      "รายได้สุทธิ": "17750.00"
    }
  }
}
```

**Response**
```json
{
  "status": "success",
  "emp_id": "512052",
  "month": "November2025"
}
```

---

### 🔹 Manage Salary Item Metadata

**GET / POST / DELETE** `/salary_items/meta`

| Method | Description |
|--------|--------------|
| **GET** | List all salary items |
| **POST** | Add or update classification |
| **DELETE** | Remove salary item |

**Example POST Body**
```json
{
  "item_name": "เงินเดือน",
  "item_group": "earnings",
  "remark": "Base salary"
}
```

---

## 🧮 Database Schema

```
salary_sheets (1) ──< salary_items >── (1) employees
                          │
                          └── salary_item_meta
```

| Table | Description |
|--------|-------------|
| **employees** | Employee master (code, name, status) |
| **salary_sheets** | Payroll month-year record |
| **salary_items** | Detailed earnings & deductions |
| **salary_item_meta** | Master classification table |

---

## 🧠 Design Highlights

- ✅ **Idempotent Uploads** – Re-uploading the same month overwrites cleanly.  
- ✅ **Cache Optimization** – Uses `@lru_cache` to minimize DB reads.  
- ✅ **Dynamic Schema** – Easily extendable for new item types.  
- ✅ **Thai Localization** – Supports Thai-language field names and Buddhist years.  
- ✅ **Enterprise Ready** – Works with Jenkins, Render, or Docker.  

---

## 📜 License

**MIT License © 2025 MenaTech Thailand**  
Developed by **Narongkorn (Plug)** – Business Intelligence & Backend Engineering.

---

## 🔮 Future Roadmap

- ✅ Excel Export Endpoint `/export_excel`  
- ✅ Auth Tokens for Admin Routes  
- ✅ Docker Compose for Local MySQL  
- ✅ RESTful dashboard (Flask-Admin or Streamlit)  
- ✅ CI/CD pipeline with Render / Jenkins
