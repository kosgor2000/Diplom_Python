# -*- coding: utf-8 -*-
import os
import sqlite3
import datetime
import random
import math
from pathlib import Path

DB_FILE = Path(__file__).resolve().parent / "data" / "komus_ims.sqlite3"

SCHEMA_SQL = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS settings(
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS users(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  username TEXT NOT NULL UNIQUE,
  role TEXT NOT NULL CHECK(role IN ('PROCUREMENT','WAREHOUSE','PLANNING','FINANCE','IT','MANAGER'))
);

CREATE TABLE IF NOT EXISTS suppliers(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE,
  lead_time_days INTEGER NOT NULL DEFAULT 7,
  reliability REAL NOT NULL DEFAULT 0.95
);

CREATE TABLE IF NOT EXISTS warehouses(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS products(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  sku TEXT NOT NULL UNIQUE,
  name TEXT NOT NULL,
  category TEXT NOT NULL,
  uom TEXT NOT NULL DEFAULT 'шт',
  unit_cost REAL NOT NULL DEFAULT 0.0,
  is_perishable INTEGER NOT NULL DEFAULT 0,
  shelf_life_days INTEGER,
  substitution_group TEXT,
  is_active INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS product_supplier(
  product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
  supplier_id INTEGER NOT NULL REFERENCES suppliers(id) ON DELETE CASCADE,
  lead_time_days INTEGER,
  min_order_qty REAL NOT NULL DEFAULT 1,
  PRIMARY KEY(product_id, supplier_id)
);

CREATE TABLE IF NOT EXISTS stock_lots(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
  warehouse_id INTEGER NOT NULL REFERENCES warehouses(id) ON DELETE CASCADE,
  lot_no TEXT NOT NULL,
  expiry_date TEXT, -- YYYY-MM-DD or NULL
  qty_on_hand REAL NOT NULL DEFAULT 0,
  qty_reserved REAL NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_stock_lots_prod_wh ON stock_lots(product_id, warehouse_id);

CREATE TABLE IF NOT EXISTS demand_history(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
  ddate TEXT NOT NULL, -- YYYY-MM-DD
  qty REAL NOT NULL DEFAULT 0,
  channel TEXT NOT NULL DEFAULT 'B2B' -- B2B/B2C/OMNI
);

CREATE INDEX IF NOT EXISTS idx_demand_prod_date ON demand_history(product_id, ddate);

CREATE TABLE IF NOT EXISTS sales_orders(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  order_no TEXT NOT NULL UNIQUE,
  customer TEXT NOT NULL,
  channel TEXT NOT NULL DEFAULT 'B2B',
  status TEXT NOT NULL DEFAULT 'NEW' CHECK(status IN ('NEW','RESERVED','PARTIAL','CANCELLED','SHIPPED')),
  created_at TEXT NOT NULL,
  promised_date TEXT
);

CREATE TABLE IF NOT EXISTS sales_order_lines(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  order_id INTEGER NOT NULL REFERENCES sales_orders(id) ON DELETE CASCADE,
  product_id INTEGER NOT NULL REFERENCES products(id),
  qty REAL NOT NULL,
  qty_reserved REAL NOT NULL DEFAULT 0,
  price REAL NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS reservations(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  order_line_id INTEGER NOT NULL REFERENCES sales_order_lines(id) ON DELETE CASCADE,
  stock_lot_id INTEGER NOT NULL REFERENCES stock_lots(id) ON DELETE CASCADE,
  qty REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS purchase_orders(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  po_no TEXT NOT NULL UNIQUE,
  supplier_id INTEGER NOT NULL REFERENCES suppliers(id),
  status TEXT NOT NULL DEFAULT 'DRAFT' CHECK(status IN ('DRAFT','SENT','RECEIVED','CANCELLED')),
  created_at TEXT NOT NULL,
  expected_date TEXT
);

CREATE TABLE IF NOT EXISTS purchase_order_lines(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  po_id INTEGER NOT NULL REFERENCES purchase_orders(id) ON DELETE CASCADE,
  product_id INTEGER NOT NULL REFERENCES products(id),
  qty REAL NOT NULL,
  unit_cost REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS receipts(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  po_id INTEGER REFERENCES purchase_orders(id),
  warehouse_id INTEGER NOT NULL REFERENCES warehouses(id),
  received_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS receipt_lines(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  receipt_id INTEGER NOT NULL REFERENCES receipts(id) ON DELETE CASCADE,
  product_id INTEGER NOT NULL REFERENCES products(id),
  lot_no TEXT NOT NULL,
  expiry_date TEXT,
  qty_received REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS forecast_results(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
  period_start TEXT NOT NULL,
  period_end TEXT NOT NULL,
  method TEXT NOT NULL,
  forecast_qty REAL NOT NULL,
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS replenishment_params(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
  supplier_id INTEGER REFERENCES suppliers(id),
  lead_time_days INTEGER NOT NULL,
  service_level REAL NOT NULL,
  demand_mean REAL NOT NULL,
  demand_std REAL NOT NULL,
  safety_stock REAL NOT NULL,
  rop REAL NOT NULL,
  min_level REAL NOT NULL,
  max_level REAL NOT NULL,
  abc_class TEXT,
  xyz_class TEXT,
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS tasks(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  title TEXT NOT NULL,
  task_type TEXT NOT NULL,
  related_entity TEXT,
  related_id INTEGER,
  priority TEXT NOT NULL DEFAULT 'MEDIUM' CHECK(priority IN ('LOW','MEDIUM','HIGH','CRITICAL')),
  status TEXT NOT NULL DEFAULT 'OPEN' CHECK(status IN ('OPEN','IN_PROGRESS','DONE','CANCELLED')),
  assigned_to INTEGER REFERENCES users(id),
  created_at TEXT NOT NULL,
  due_date TEXT,
  notes TEXT
);

CREATE TABLE IF NOT EXISTS notifications(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  created_at TEXT NOT NULL,
  level TEXT NOT NULL DEFAULT 'INFO' CHECK(level IN ('INFO','WARN','ERROR')),
  message TEXT NOT NULL,
  is_read INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS integration_inbox(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  source TEXT NOT NULL,
  received_at TEXT NOT NULL,
  payload_json TEXT NOT NULL,
  processed INTEGER NOT NULL DEFAULT 0,
  processed_at TEXT
);
"""

def connect():
    DB_FILE.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(SCHEMA_SQL)
    # Defaults
    defaults = {
        "default_service_level": "0.95",
        "forecast_window_days": "90",
        "forecast_horizon_days": "14",
        "review_period_days": "7",
        "default_forecast_method": "ES"  # MA/ES/CROSTON
    }
    for k, v in defaults.items():
        conn.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (k, v))
    conn.commit()

def _set_notification(conn, level: str, message: str):
    now = datetime.datetime.now().isoformat(timespec="seconds")
    conn.execute("INSERT INTO notifications(created_at, level, message) VALUES(?,?,?)", (now, level, message))

def seed_demo_data(conn: sqlite3.Connection) -> None:
    # Only seed if empty
    cur = conn.execute("SELECT COUNT(*) AS c FROM products")
    if cur.fetchone()["c"] > 0:
        return

    # Users
    users = [
        ("planner", "PLANNING"),
        ("buyer", "PROCUREMENT"),
        ("warehouse", "WAREHOUSE"),
        ("finance", "FINANCE"),
        ("it", "IT"),
        ("manager", "MANAGER"),
    ]
    conn.executemany("INSERT INTO users(username, role) VALUES(?,?)", users)

    # Suppliers
    suppliers = [
        ("ООО Поставщик-Центр", 7, 0.95),
        ("АО Логистик-Партнер", 10, 0.93),
        ("ООО Импорт-Снаб", 14, 0.90),
    ]
    conn.executemany("INSERT INTO suppliers(name, lead_time_days, reliability) VALUES(?,?,?)", suppliers)

    # Warehouses
    warehouses = [("Склад Москва",), ("Склад СПб",)]
    conn.executemany("INSERT INTO warehouses(name) VALUES(?)", warehouses)

    # Products (примерная номенклатура)
    products = [
        ("PAP-A4-80", "Бумага A4 80г/м2 (пачка)", "Бумага", "уп", 210.0, 0, None, "PAPER"),
        ("PEN-GEL-BL", "Ручка гелевая синяя", "Канцтовары", "шт", 18.0, 0, None, "PEN"),
        ("NOTE-A5", "Блокнот A5", "Канцтовары", "шт", 95.0, 0, None, "NOTE"),
        ("CLEAN-SPR", "Спрей чистящий (500 мл)", "Хозтовары", "шт", 145.0, 0, None, "CLEAN"),
        ("COFFEE-1KG", "Кофе зерновой 1кг", "Питание", "уп", 1250.0, 1, 365, "COFFEE"),
        ("TEA-100", "Чай пакетированный 100шт", "Питание", "уп", 420.0, 1, 540, "TEA"),
        ("TONER-HP", "Картридж тонер (аналог HP)", "Оргтехника", "шт", 1850.0, 0, None, "TONER"),
        ("FOLDER-A4", "Папка-регистратор A4", "Канцтовары", "шт", 160.0, 0, None, "FOLDER"),
        ("GLOVES-NIT", "Перчатки нитриловые (уп)", "Хозтовары", "уп", 350.0, 0, None, "GLOVES"),
        ("MARKER-BL", "Маркер перманентный черный", "Канцтовары", "шт", 55.0, 0, None, "MARKER"),
    ]
    conn.executemany("""
        INSERT INTO products(sku, name, category, uom, unit_cost, is_perishable, shelf_life_days, substitution_group)
        VALUES(?,?,?,?,?,?,?,?)
    """, products)

    # Map product suppliers (simple: 1 supplier per product)
    supplier_ids = [r["id"] for r in conn.execute("SELECT id FROM suppliers ORDER BY id").fetchall()]
    product_ids = [r["id"] for r in conn.execute("SELECT id FROM products ORDER BY id").fetchall()]
    ps = []
    for idx, pid in enumerate(product_ids):
        sid = supplier_ids[idx % len(supplier_ids)]
        lead = conn.execute("SELECT lead_time_days FROM suppliers WHERE id=?", (sid,)).fetchone()["lead_time_days"]
        ps.append((pid, sid, lead, 1))
    conn.executemany("INSERT INTO product_supplier(product_id, supplier_id, lead_time_days, min_order_qty) VALUES(?,?,?,?)", ps)

    # Initial stock lots with FEFO for perishable
    wh_ids = [r["id"] for r in conn.execute("SELECT id FROM warehouses ORDER BY id").fetchall()]
    now = datetime.date.today()
    for pid in product_ids:
        prod = conn.execute("SELECT is_perishable, shelf_life_days, unit_cost FROM products WHERE id=?", (pid,)).fetchone()
        for wh in wh_ids:
            lot_count = 2 if prod["is_perishable"] else 1
            for li in range(lot_count):
                lot_no = f"LOT-{pid:03d}-{wh:02d}-{li+1}"
                if prod["is_perishable"] and prod["shelf_life_days"]:
                    expiry = (now + datetime.timedelta(days=int(prod["shelf_life_days"] * (0.4 + 0.4*li)))).isoformat()
                else:
                    expiry = None
                qty = random.randint(20, 120) if prod["unit_cost"] < 500 else random.randint(5, 40)
                created = datetime.datetime.now().isoformat(timespec="seconds")
                conn.execute("""
                    INSERT INTO stock_lots(product_id, warehouse_id, lot_no, expiry_date, qty_on_hand, qty_reserved, created_at)
                    VALUES(?,?,?,?,?,?,?)
                """, (pid, wh, lot_no, expiry, qty, 0, created))

    # Demand history for last 120 days (simulate)
    start = now - datetime.timedelta(days=120)
    for pid in product_ids:
        base = random.uniform(0.2, 3.5)  # daily mean
        intermittent = random.random() < 0.2  # some intermittent items
        for d in range(121):
            date = start + datetime.timedelta(days=d)
            # seasonality weekly
            season = 1.0 + 0.25*math.sin(2*math.pi*(d/7.0))
            lam = base * season
            if intermittent and random.random() < 0.75:
                qty = 0.0
            else:
                qty = max(0.0, random.gauss(lam, lam*0.6))
            # occasional promo spikes
            if random.random() < 0.03:
                qty *= random.uniform(2.0, 4.0)
            channel = random.choice(["B2B","B2C","OMNI"])
            conn.execute("INSERT INTO demand_history(product_id, ddate, qty, channel) VALUES(?,?,?,?)",
                         (pid, date.isoformat(), float(round(qty, 2)), channel))

    _set_notification(conn, "INFO", "База данных создана и заполнена демонстрационными данными.")
    conn.commit()
