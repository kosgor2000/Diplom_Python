# -*- coding: utf-8 -*-
import sqlite3
import datetime
import math
import statistics
from typing import List, Dict, Tuple, Optional

# ----------------- Utilities -----------------
def _now_ts() -> str:
    return datetime.datetime.now().isoformat(timespec="seconds")

def get_setting(conn: sqlite3.Connection, key: str, default: str) -> str:
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row["value"] if row else default

def set_setting(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute("INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))
    conn.commit()

def add_notification(conn: sqlite3.Connection, level: str, message: str) -> None:
    conn.execute("INSERT INTO notifications(created_at, level, message, is_read) VALUES(?,?,?,0)", (_now_ts(), level, message))
    conn.commit()

# ----------------- Demand & Forecast -----------------
def fetch_daily_demand(conn: sqlite3.Connection, product_id: int, days: int) -> List[float]:
    end = datetime.date.today()
    start = end - datetime.timedelta(days=days-1)
    rows = conn.execute("""
        SELECT ddate, SUM(qty) AS qty
        FROM demand_history
        WHERE product_id=? AND ddate BETWEEN ? AND ?
        GROUP BY ddate
        ORDER BY ddate
    """, (product_id, start.isoformat(), end.isoformat())).fetchall()
    m = {r["ddate"]: float(r["qty"]) for r in rows}
    series = []
    for i in range(days):
        d = (start + datetime.timedelta(days=i)).isoformat()
        series.append(m.get(d, 0.0))
    return series

def forecast_ma(series: List[float], horizon: int, window: int=14) -> float:
    if not series:
        return 0.0
    w = min(window, len(series))
    avg = sum(series[-w:]) / w
    return avg * horizon

def forecast_es(series: List[float], horizon: int, alpha: float=0.3) -> float:
    if not series:
        return 0.0
    level = series[0]
    for x in series[1:]:
        level = alpha*x + (1-alpha)*level
    # naive: constant forecast = level
    return max(0.0, level) * horizon

def forecast_croston(series: List[float], horizon: int, alpha: float=0.1) -> float:
    # Croston for intermittent demand (very simple)
    if not series:
        return 0.0
    z = 0.0  # demand size
    p = 1.0  # interval
    q = 0.0
    a = alpha
    interval = 0
    first = True
    for x in series:
        interval += 1
        if x > 0:
            if first:
                z = x
                p = interval
                first = False
            else:
                z = z + a*(x - z)
                p = p + a*(interval - p)
            q = z / max(p, 1e-9)
            interval = 0
    if first:
        # all zeros
        return 0.0
    return max(0.0, q) * horizon

def compute_forecast(conn: sqlite3.Connection, product_id: int, method: str, window_days: int, horizon_days: int) -> float:
    series = fetch_daily_demand(conn, product_id, window_days)
    method = method.upper().strip()
    if method == "MA":
        fc = forecast_ma(series, horizon_days, window=min(14, window_days))
    elif method == "CROSTON":
        fc = forecast_croston(series, horizon_days)
    else:
        fc = forecast_es(series, horizon_days)
    return float(round(fc, 2))

# ----------------- ABC/XYZ -----------------
def compute_abc_xyz(conn: sqlite3.Connection, window_days: int=365) -> Dict[int, Tuple[str, str]]:
    end = datetime.date.today()
    start = end - datetime.timedelta(days=window_days-1)
    rows = conn.execute("""
        SELECT p.id AS product_id, p.unit_cost AS cost, SUM(d.qty) AS qty_sum
        FROM products p
        LEFT JOIN demand_history d ON d.product_id=p.id AND d.ddate BETWEEN ? AND ?
        WHERE p.is_active=1
        GROUP BY p.id, p.unit_cost
    """, (start.isoformat(), end.isoformat())).fetchall()
    values = []
    for r in rows:
        qty = float(r["qty_sum"] or 0.0)
        val = qty * float(r["cost"] or 0.0)
        values.append((int(r["product_id"]), val))
    values.sort(key=lambda x: x[1], reverse=True)
    total = sum(v for _, v in values) or 1.0
    cum = 0.0
    abc = {}
    for pid, v in values:
        cum += v
        share = cum / total
        if share <= 0.8:
            cls = "A"
        elif share <= 0.95:
            cls = "B"
        else:
            cls = "C"
        abc[pid] = cls

    # XYZ by coefficient of variation of weekly demand
    xyz = {}
    for pid, _ in values:
        # build weekly buckets for last 12 weeks
        weeks = 12
        series = []
        for w in range(weeks):
            w_end = end - datetime.timedelta(days=7*w)
            w_start = w_end - datetime.timedelta(days=6)
            qty = conn.execute("""
                SELECT SUM(qty) AS s FROM demand_history
                WHERE product_id=? AND ddate BETWEEN ? AND ?
            """, (pid, w_start.isoformat(), w_end.isoformat())).fetchone()["s"]
            series.append(float(qty or 0.0))
        mean = sum(series)/len(series) if series else 0.0
        std = statistics.pstdev(series) if len(series) > 1 else 0.0
        cv = (std/mean) if mean > 1e-9 else 999.0 if std > 0 else 0.0
        if cv <= 0.5:
            cls = "X"
        elif cv <= 1.0:
            cls = "Y"
        else:
            cls = "Z"
        xyz[pid] = cls

    out = {pid: (abc.get(pid, "C"), xyz.get(pid, "Z")) for pid, _ in values}
    return out

# ----------------- Replenishment parameters -----------------
_Z_MAP = {
    0.90: 1.2816,
    0.95: 1.6449,
    0.97: 1.8808,
    0.98: 2.0537,
    0.99: 2.3263,
}

def z_from_service_level(sl: float) -> float:
    # snap to nearest in map
    keys = sorted(_Z_MAP.keys())
    best = min(keys, key=lambda k: abs(k - sl))
    return _Z_MAP[best]

def get_available_qty(conn: sqlite3.Connection, product_id: int, warehouse_id: Optional[int]=None) -> float:
    if warehouse_id is None:
        row = conn.execute("""
            SELECT SUM(qty_on_hand - qty_reserved) AS a
            FROM stock_lots
            WHERE product_id=?
        """, (product_id,)).fetchone()
    else:
        row = conn.execute("""
            SELECT SUM(qty_on_hand - qty_reserved) AS a
            FROM stock_lots
            WHERE product_id=? AND warehouse_id=?
        """, (product_id, warehouse_id)).fetchone()
    return float(row["a"] or 0.0)

def compute_replenishment_for_product(
    conn: sqlite3.Connection,
    product_id: int,
    supplier_id: int,
    lead_time_days: int,
    service_level: float,
    window_days: int,
    horizon_days: int,
    review_period_days: int,
    method: str
) -> Dict[str, float]:
    series = fetch_daily_demand(conn, product_id, window_days)
    # demand mean/std per day
    mean = sum(series)/len(series) if series else 0.0
    std = statistics.pstdev(series) if len(series) > 1 else 0.0
    z = z_from_service_level(service_level)
    # Safety stock with LT variability ignored: z * sigma * sqrt(LT)
    safety = z * std * math.sqrt(max(lead_time_days, 1))
    rop = mean * lead_time_days + safety
    min_level = rop
    max_level = rop + mean * review_period_days
    # forecast (for info)
    fc = compute_forecast(conn, product_id, method, window_days, horizon_days)
    return {
        "demand_mean": float(round(mean, 4)),
        "demand_std": float(round(std, 4)),
        "safety_stock": float(round(safety, 2)),
        "rop": float(round(rop, 2)),
        "min_level": float(round(min_level, 2)),
        "max_level": float(round(max_level, 2)),
        "forecast_horizon": float(round(fc, 2)),
    }

def recompute_all_parameters(conn: sqlite3.Connection) -> int:
    sl = float(get_setting(conn, "default_service_level", "0.95"))
    window_days = int(get_setting(conn, "forecast_window_days", "90"))
    horizon_days = int(get_setting(conn, "forecast_horizon_days", "14"))
    review_days = int(get_setting(conn, "review_period_days", "7"))
    method = get_setting(conn, "default_forecast_method", "ES")
    abcxyz = compute_abc_xyz(conn, window_days=365)

    products = conn.execute("SELECT id FROM products WHERE is_active=1 ORDER BY sku").fetchall()
    inserted = 0
    for pr in products:
        pid = int(pr["id"])
        # pick supplier mapping (first)
        row = conn.execute("""
            SELECT ps.supplier_id, COALESCE(ps.lead_time_days, s.lead_time_days) AS lt
            FROM product_supplier ps
            JOIN suppliers s ON s.id=ps.supplier_id
            WHERE ps.product_id=?
            ORDER BY ps.supplier_id
            LIMIT 1
        """, (pid,)).fetchone()
        if not row:
            continue
        sid = int(row["supplier_id"])
        lt = int(row["lt"] or 7)

        r = compute_replenishment_for_product(conn, pid, sid, lt, sl, window_days, horizon_days, review_days, method)
        abc, xyz = abcxyz.get(pid, ("C","Z"))
        conn.execute("""
            INSERT INTO replenishment_params(
              product_id, supplier_id, lead_time_days, service_level,
              demand_mean, demand_std, safety_stock, rop, min_level, max_level,
              abc_class, xyz_class, created_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            pid, sid, lt, sl,
            r["demand_mean"], r["demand_std"], r["safety_stock"], r["rop"], r["min_level"], r["max_level"],
            abc, xyz, _now_ts()
        ))
        inserted += 1
    conn.commit()
    add_notification(conn, "INFO", f"Пересчитаны параметры пополнения (строк: {inserted}).")
    return inserted

def build_replenishment_plan(conn: sqlite3.Connection, warehouse_id: Optional[int]=None) -> List[Dict]:
    # use latest params per product
    rows = conn.execute("""
        SELECT rp.*,
               p.sku, p.name, p.unit_cost,
               s.name AS supplier_name
        FROM replenishment_params rp
        JOIN (
          SELECT product_id, MAX(created_at) AS mx
          FROM replenishment_params
          GROUP BY product_id
        ) x ON x.product_id=rp.product_id AND x.mx=rp.created_at
        JOIN products p ON p.id=rp.product_id
        LEFT JOIN suppliers s ON s.id=rp.supplier_id
        WHERE p.is_active=1
        ORDER BY p.sku
    """).fetchall()

    plan = []
    for r in rows:
        pid = int(r["product_id"])
        avail = get_available_qty(conn, pid, warehouse_id=warehouse_id)
        min_level = float(r["min_level"])
        max_level = float(r["max_level"])
        if avail < min_level:
            qty_to_order = max_level - avail
        else:
            qty_to_order = 0.0
        plan.append({
            "product_id": pid,
            "sku": r["sku"],
            "name": r["name"],
            "supplier": r["supplier_name"] or "",
            "available": float(round(avail, 2)),
            "min_level": float(round(min_level, 2)),
            "max_level": float(round(max_level, 2)),
            "qty_to_order": float(round(max(qty_to_order, 0.0), 2)),
            "abc": r["abc_class"] or "",
            "xyz": r["xyz_class"] or "",
        })
    return plan

# ----------------- ATP & Reservations (FEFO/FIFO) -----------------
def atp_check(conn: sqlite3.Connection, product_id: int, qty: float, warehouse_id: Optional[int]=None) -> Tuple[bool, float]:
    avail = get_available_qty(conn, product_id, warehouse_id=warehouse_id)
    return (avail >= qty, avail)

def suggest_substitutions(conn: sqlite3.Connection, product_id: int, qty: float, warehouse_id: Optional[int]=None) -> List[Dict]:
    row = conn.execute("SELECT substitution_group FROM products WHERE id=?", (product_id,)).fetchone()
    if not row or not row["substitution_group"]:
        return []
    grp = row["substitution_group"]
    candidates = conn.execute("""
        SELECT id, sku, name
        FROM products
        WHERE is_active=1 AND substitution_group=? AND id<>?
        ORDER BY sku
    """, (grp, product_id)).fetchall()
    out = []
    for c in candidates:
        ok, avail = atp_check(conn, int(c["id"]), qty, warehouse_id)
        if avail > 0:
            out.append({"product_id": int(c["id"]), "sku": c["sku"], "name": c["name"], "available": float(round(avail,2))})
    out.sort(key=lambda x: x["available"], reverse=True)
    return out

def _fetch_lots_for_reservation(conn: sqlite3.Connection, product_id: int, warehouse_id: Optional[int]) -> List[sqlite3.Row]:
    # FEFO: by expiry_date asc (NULL last), then created_at asc
    if warehouse_id is None:
        rows = conn.execute("""
            SELECT *
            FROM stock_lots
            WHERE product_id=? AND (qty_on_hand - qty_reserved) > 0
            ORDER BY (expiry_date IS NULL) ASC, expiry_date ASC, created_at ASC
        """, (product_id,)).fetchall()
    else:
        rows = conn.execute("""
            SELECT *
            FROM stock_lots
            WHERE product_id=? AND warehouse_id=? AND (qty_on_hand - qty_reserved) > 0
            ORDER BY (expiry_date IS NULL) ASC, expiry_date ASC, created_at ASC
        """, (product_id, warehouse_id)).fetchall()
    return rows

def reserve_stock_for_order_line(conn: sqlite3.Connection, order_line_id: int, warehouse_id: Optional[int]=None) -> Dict:
    line = conn.execute("""
        SELECT sol.id, sol.product_id, sol.qty, sol.qty_reserved, p.sku, p.name
        FROM sales_order_lines sol
        JOIN products p ON p.id=sol.product_id
        WHERE sol.id=?
    """, (order_line_id,)).fetchone()
    if not line:
        return {"ok": False, "message": "Строка заказа не найдена."}
    need = float(line["qty"]) - float(line["qty_reserved"])
    if need <= 0:
        return {"ok": True, "message": "Резерв уже выполнен."}

    ok, avail = atp_check(conn, int(line["product_id"]), need, warehouse_id)
    if not ok:
        return {"ok": False, "message": f"Недостаточно доступного остатка. Доступно: {avail:.2f}, нужно: {need:.2f}"}

    lots = _fetch_lots_for_reservation(conn, int(line["product_id"]), warehouse_id)
    allocated = 0.0
    for lot in lots:
        free = float(lot["qty_on_hand"]) - float(lot["qty_reserved"])
        if free <= 0:
            continue
        take = min(free, need - allocated)
        if take <= 0:
            break
        # update lot reserved
        conn.execute("UPDATE stock_lots SET qty_reserved = qty_reserved + ? WHERE id=?", (take, int(lot["id"])))
        # reservation record
        conn.execute("INSERT INTO reservations(order_line_id, stock_lot_id, qty) VALUES(?,?,?)", (int(line["id"]), int(lot["id"]), take))
        allocated += take
        if allocated >= need - 1e-9:
            break

    conn.execute("UPDATE sales_order_lines SET qty_reserved = qty_reserved + ? WHERE id=?", (allocated, int(line["id"])))
    conn.commit()
    return {"ok": True, "message": f"Зарезервировано {allocated:.2f} по FEFO/FIFO."}

# ----------------- Orders & PO/Receipts -----------------
def create_sales_order(conn: sqlite3.Connection, customer: str, channel: str, promised_date: Optional[str]=None) -> int:
    # order_no simple
    order_no = f"SO-{datetime.datetime.now().strftime('%Y%m%d-%H%M%S')}"
    conn.execute("""
        INSERT INTO sales_orders(order_no, customer, channel, status, created_at, promised_date)
        VALUES(?,?,?,?,?,?)
    """, (order_no, customer, channel, "NEW", _now_ts(), promised_date))
    oid = conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"]
    conn.commit()
    return int(oid)

def add_sales_order_line(conn: sqlite3.Connection, order_id: int, product_id: int, qty: float, price: float) -> int:
    conn.execute("""
        INSERT INTO sales_order_lines(order_id, product_id, qty, qty_reserved, price)
        VALUES(?,?,?,?,?)
    """, (order_id, product_id, qty, 0.0, price))
    lid = conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"]
    conn.commit()
    return int(lid)

def create_purchase_order(conn: sqlite3.Connection, supplier_id: int, expected_date: Optional[str]=None) -> int:
    po_no = f"PO-{datetime.datetime.now().strftime('%Y%m%d-%H%M%S')}"
    conn.execute("""
        INSERT INTO purchase_orders(po_no, supplier_id, status, created_at, expected_date)
        VALUES(?,?,?,?,?)
    """, (po_no, supplier_id, "DRAFT", _now_ts(), expected_date))
    po_id = conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"]
    conn.commit()
    return int(po_id)

def add_purchase_order_line(conn: sqlite3.Connection, po_id: int, product_id: int, qty: float, unit_cost: float) -> int:
    conn.execute("""
        INSERT INTO purchase_order_lines(po_id, product_id, qty, unit_cost)
        VALUES(?,?,?,?)
    """, (po_id, product_id, qty, unit_cost))
    lid = conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"]
    conn.commit()
    return int(lid)

def receive_purchase_order(conn: sqlite3.Connection, po_id: int, warehouse_id: int, lines: List[Dict]) -> int:
    # lines: {product_id, lot_no, expiry_date, qty}
    conn.execute("INSERT INTO receipts(po_id, warehouse_id, received_at) VALUES(?,?,?)", (po_id, warehouse_id, _now_ts()))
    rid = int(conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"])
    for ln in lines:
        pid = int(ln["product_id"])
        qty = float(ln["qty"])
        lot = str(ln["lot_no"])
        exp = ln.get("expiry_date")
        conn.execute("""
            INSERT INTO receipt_lines(receipt_id, product_id, lot_no, expiry_date, qty_received)
            VALUES(?,?,?,?,?)
        """, (rid, pid, lot, exp, qty))
        # add into stock_lots as new lot
        conn.execute("""
            INSERT INTO stock_lots(product_id, warehouse_id, lot_no, expiry_date, qty_on_hand, qty_reserved, created_at)
            VALUES(?,?,?,?,?,?,?)
        """, (pid, warehouse_id, lot, exp, qty, 0.0, _now_ts()))
    conn.execute("UPDATE purchase_orders SET status='RECEIVED' WHERE id=?", (po_id,))
    conn.commit()
    add_notification(conn, "INFO", f"Приёмка выполнена по PO id={po_id}.")
    return rid

# ----------------- Workflow & KPI -----------------
def create_task(conn: sqlite3.Connection, title: str, task_type: str, priority: str="MEDIUM", assigned_to: Optional[int]=None, due_date: Optional[str]=None, related_entity: Optional[str]=None, related_id: Optional[int]=None, notes: str="") -> int:
    conn.execute("""
        INSERT INTO tasks(title, task_type, related_entity, related_id, priority, status, assigned_to, created_at, due_date, notes)
        VALUES(?,?,?,?,?,?,?,?,?,?)
    """, (title, task_type, related_entity, related_id, priority, "OPEN", assigned_to, _now_ts(), due_date, notes))
    tid = int(conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"])
    conn.commit()
    return tid

def compute_kpi(conn: sqlite3.Connection) -> Dict[str, float]:
    # OOS: active sku with available == 0 and demand in last 30 days > 0
    end = datetime.date.today()
    start = end - datetime.timedelta(days=29)
    rows = conn.execute("""
        SELECT p.id, p.sku,
               COALESCE(SUM(d.qty),0) AS d30,
               COALESCE((SELECT SUM(sl.qty_on_hand - sl.qty_reserved) FROM stock_lots sl WHERE sl.product_id=p.id),0) AS avail
        FROM products p
        LEFT JOIN demand_history d ON d.product_id=p.id AND d.ddate BETWEEN ? AND ?
        WHERE p.is_active=1
        GROUP BY p.id, p.sku
    """, (start.isoformat(), end.isoformat())).fetchall()
    total = len(rows) or 1
    oos = 0
    for r in rows:
        if float(r["avail"]) <= 0.0001 and float(r["d30"]) > 0.0001:
            oos += 1
    osa = 1.0 - (oos / total)

    # Inventory value and rough turnover: (COGS in 90 days annualized) / avg inventory value
    inv_val = conn.execute("""
        SELECT COALESCE(SUM((qty_on_hand - qty_reserved) * p.unit_cost),0) AS v
        FROM stock_lots sl
        JOIN products p ON p.id=sl.product_id
        WHERE p.is_active=1
    """).fetchone()["v"]
    inv_val = float(inv_val or 0.0)

    start90 = end - datetime.timedelta(days=89)
    cogs90 = conn.execute("""
        SELECT COALESCE(SUM(d.qty * p.unit_cost),0) AS c
        FROM demand_history d
        JOIN products p ON p.id=d.product_id
        WHERE d.ddate BETWEEN ? AND ?
    """, (start90.isoformat(), end.isoformat())).fetchone()["c"]
    cogs90 = float(cogs90 or 0.0)
    annual_cogs = cogs90 * (365/90.0) if cogs90 > 0 else 0.0
    turnover = (annual_cogs / inv_val) if inv_val > 1e-9 else 0.0

    # Obsolete estimate: items with demand 0 in last 90 days but have inventory
    start90 = end - datetime.timedelta(days=89)
    obs_rows = conn.execute("""
        SELECT p.id,
               COALESCE(SUM(d.qty),0) AS d90,
               COALESCE((SELECT SUM(sl.qty_on_hand - sl.qty_reserved) FROM stock_lots sl WHERE sl.product_id=p.id),0) AS avail,
               p.unit_cost
        FROM products p
        LEFT JOIN demand_history d ON d.product_id=p.id AND d.ddate BETWEEN ? AND ?
        WHERE p.is_active=1
        GROUP BY p.id, p.unit_cost
    """, (start90.isoformat(), end.isoformat())).fetchall()
    obs_val = 0.0
    for r in obs_rows:
        if float(r["d90"]) <= 0.0001 and float(r["avail"]) > 0.0001:
            obs_val += float(r["avail"]) * float(r["unit_cost"])
    return {
        "sku_total": float(total),
        "oos_sku": float(oos),
        "osa": float(round(osa, 4)),
        "inventory_value": float(round(inv_val, 2)),
        "turnover": float(round(turnover, 4)),
        "obsolete_value": float(round(obs_val, 2)),
    }

# ----------------- Integration inbox -----------------
def push_integration_message(conn: sqlite3.Connection, source: str, payload: dict) -> int:
    conn.execute("""
        INSERT INTO integration_inbox(source, received_at, payload_json, processed)
        VALUES(?,?,?,0)
    """, (source, _now_ts(), __import__("json").dumps(payload, ensure_ascii=False), 0))
    mid = int(conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"])
    conn.commit()
    return mid

def process_integration_message(conn: sqlite3.Connection, msg_id: int) -> str:
    row = conn.execute("SELECT * FROM integration_inbox WHERE id=?", (msg_id,)).fetchone()
    if not row:
        return "Сообщение не найдено."
    if int(row["processed"]) == 1:
        return "Сообщение уже обработано."
    payload = __import__("json").loads(row["payload_json"])
    # Supported: {"type":"supplier_eta_update","supplier":"...","po_no":"...","expected_date":"YYYY-MM-DD"}
    mtype = payload.get("type", "")
    if mtype == "supplier_eta_update":
        po_no = payload.get("po_no")
        expected_date = payload.get("expected_date")
        if po_no and expected_date:
            conn.execute("UPDATE purchase_orders SET expected_date=? WHERE po_no=?", (expected_date, po_no))
            add_notification(conn, "WARN", f"Обновлена ETA поставки по {po_no}: {expected_date}.")
            # create task
            create_task(conn, f"Проверить влияние ETA по {po_no} на план пополнения", "ETA_CHANGE", priority="HIGH")
            res = f"ETA обновлена для {po_no}."
        else:
            res = "Недостаточно данных для ETA обновления."
    else:
        res = f"Неизвестный тип сообщения: {mtype}"
    conn.execute("UPDATE integration_inbox SET processed=1, processed_at=? WHERE id=?", (_now_ts(), msg_id))
    conn.commit()
    return res
