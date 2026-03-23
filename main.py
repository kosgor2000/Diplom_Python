# -*- coding: utf-8 -*-
import os
import sqlite3
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import db
import services
from ui_widgets import LabeledEntry, LabeledCombo, make_tree, show_info

APP_TITLE = "Komus IMS — управление запасами (Python + Tkinter + SQLite)"

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1280x760")
        self.minsize(1180, 700)

        self.conn = db.connect()
        db.init_db(self.conn)
        db.seed_demo_data(self.conn)

        # Precompute replenishment params if none
        c = self.conn.execute("SELECT COUNT(*) AS c FROM replenishment_params").fetchone()["c"]
        if int(c) == 0:
            services.recompute_all_parameters(self.conn)

        self._build_ui()
        self.refresh_all()

    # ---------------- UI ----------------
    def _build_ui(self):
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Label(top, text=APP_TITLE, font=("Times New Roman", 14, "bold")).pack(side="left")

        ttk.Button(top, text="Пересчитать параметры", command=self.on_recompute).pack(side="right", padx=6)
        ttk.Button(top, text="Обновить", command=self.refresh_all).pack(side="right", padx=6)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=8)

        self.tab_dashboard = ttk.Frame(self.notebook)
        self.tab_catalog = ttk.Frame(self.notebook)
        self.tab_stock = ttk.Frame(self.notebook)
        self.tab_orders = ttk.Frame(self.notebook)
        self.tab_replen = ttk.Frame(self.notebook)
        self.tab_po = ttk.Frame(self.notebook)
        self.tab_tasks = ttk.Frame(self.notebook)
        self.tab_integration = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_dashboard, text="KPI / Дашборд")
        self.notebook.add(self.tab_catalog, text="Номенклатура (SKU)")
        self.notebook.add(self.tab_stock, text="Склад / Партии")
        self.notebook.add(self.tab_orders, text="Заказы / ATP")
        self.notebook.add(self.tab_replen, text="Прогноз и пополнение")
        self.notebook.add(self.tab_po, text="Закупки (PO) / Приёмка")
        self.notebook.add(self.tab_tasks, text="Задачи / Уведомления")
        self.notebook.add(self.tab_integration, text="Интеграции (эмуляция)")
        self.notebook.add(self.tab_settings, text="Настройки")

        self._build_dashboard()
        self._build_catalog()
        self._build_stock()
        self._build_orders()
        self._build_replen()
        self._build_po()
        self._build_tasks()
        self._build_integration()
        self._build_settings()

    # ---------------- Dashboard ----------------
    def _build_dashboard(self):
        f = self.tab_dashboard
        left = ttk.Frame(f)
        left.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        self.kpi_vars = {k: tk.StringVar(value="") for k in ["sku_total","oos_sku","osa","inventory_value","turnover","obsolete_value"]}

        grid = ttk.Frame(left)
        grid.pack(anchor="nw", fill="x")

        def row(lbl, key):
            r = ttk.Frame(grid)
            r.pack(fill="x", pady=2)
            ttk.Label(r, text=lbl, width=30).pack(side="left")
            ttk.Label(r, textvariable=self.kpi_vars[key], width=24).pack(side="left")

        row("Всего активных SKU:", "sku_total")
        row("SKU в дефиците (OOS):", "oos_sku")
        row("OSA (доля доступных):", "osa")
        row("Стоимость запасов, руб.:", "inventory_value")
        row("Оборачиваемость (оценка):", "turnover")
        row("Неликвиды/без спроса, руб.:", "obsolete_value")

        ttk.Separator(left).pack(fill="x", pady=10)

        ttk.Label(left, text="Позиции в дефиците (OOS за 30 дней):", font=("Times New Roman", 14, "bold")).pack(anchor="w", pady=(0,6))
        cols = ["SKU","Наименование","Спрос 30д","Доступно"]
        self.tree_oos, vsb = make_tree(left, cols, widths=[140, 420, 120, 120])
        self.tree_oos.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")

    def _refresh_dashboard(self):
        k = services.compute_kpi(self.conn)
        for kk, v in k.items():
            if kk in self.kpi_vars:
                self.kpi_vars[kk].set(str(v))

        # OOS list
        for i in self.tree_oos.get_children():
            self.tree_oos.delete(i)
        end = datetime.date.today()
        start = end - datetime.timedelta(days=29)
        rows = self.conn.execute("""
            SELECT p.sku, p.name,
                   COALESCE(SUM(d.qty),0) AS d30,
                   COALESCE((SELECT SUM(sl.qty_on_hand - sl.qty_reserved) FROM stock_lots sl WHERE sl.product_id=p.id),0) AS avail
            FROM products p
            LEFT JOIN demand_history d ON d.product_id=p.id AND d.ddate BETWEEN ? AND ?
            WHERE p.is_active=1
            GROUP BY p.id
            HAVING avail <= 0.0001 AND d30 > 0.0001
            ORDER BY d30 DESC
        """, (start.isoformat(), end.isoformat())).fetchall()
        for r in rows:
            self.tree_oos.insert("", "end", values=(r["sku"], r["name"], round(float(r["d30"]),2), round(float(r["avail"]),2)))

    # ---------------- Catalog ----------------
    def _build_catalog(self):
        f = self.tab_catalog
        top = ttk.Frame(f); top.pack(fill="x", padx=10, pady=10)
        ttk.Button(top, text="Добавить SKU", command=self.on_add_product).pack(side="left")
        ttk.Button(top, text="Обновить список", command=self._refresh_catalog).pack(side="left", padx=6)

        cols = ["ID","SKU","Наименование","Категория","Ед.","Себестоимость","Годн.","Срок,д","Группа замен","Активен"]
        self.tree_products, vsb = make_tree(f, cols, widths=[60,120,340,140,60,120,80,80,120,70])
        self.tree_products.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)
        vsb.pack(side="left", fill="y", pady=10)

        right = ttk.Frame(f); right.pack(side="left", fill="y", padx=10, pady=10)
        ttk.Button(right, text="Изменить активность", command=self.on_toggle_product).pack(fill="x", pady=3)
        ttk.Button(right, text="Экспорт в CSV", command=self.on_export_products).pack(fill="x", pady=3)

    def _refresh_catalog(self):
        for i in self.tree_products.get_children():
            self.tree_products.delete(i)
        rows = self.conn.execute("""
            SELECT id, sku, name, category, uom, unit_cost, is_perishable, shelf_life_days, substitution_group, is_active
            FROM products
            ORDER BY sku
        """).fetchall()
        for r in rows:
            self.tree_products.insert("", "end", values=(
                r["id"], r["sku"], r["name"], r["category"], r["uom"],
                round(float(r["unit_cost"]),2),
                "да" if int(r["is_perishable"])==1 else "нет",
                r["shelf_life_days"] if r["shelf_life_days"] is not None else "",
                r["substitution_group"] or "",
                "да" if int(r["is_active"])==1 else "нет"
            ))

    def on_add_product(self):
        win = tk.Toplevel(self); win.title("Добавить SKU"); win.geometry("560x420")
        frm = ttk.Frame(win); frm.pack(fill="both", expand=True, padx=12, pady=12)

        e_sku = LabeledEntry(frm, "SKU:"); e_sku.pack(fill="x")
        e_name = LabeledEntry(frm, "Наименование:"); e_name.pack(fill="x")
        e_cat = LabeledEntry(frm, "Категория:"); e_cat.pack(fill="x")
        e_uom = LabeledEntry(frm, "Ед.изм.:"); e_uom.var.set("шт"); e_uom.pack(fill="x")
        e_cost = LabeledEntry(frm, "Себестоимость (руб.):"); e_cost.var.set("0"); e_cost.pack(fill="x")
        e_per = LabeledCombo(frm, "Скоропорт (0/1):", ["0","1"]); e_per.var.set("0"); e_per.pack(fill="x")
        e_life = LabeledEntry(frm, "Срок годности (дней, пусто если нет):"); e_life.pack(fill="x")
        e_grp = LabeledEntry(frm, "Группа замен (опционально):"); e_grp.pack(fill="x")

        def save():
            try:
                sku = e_sku.var.get().strip()
                name = e_name.var.get().strip()
                cat = e_cat.var.get().strip()
                uom = e_uom.var.get().strip() or "шт"
                cost = float(e_cost.var.get().strip() or "0")
                per = int(e_per.var.get().strip() or "0")
                life = e_life.var.get().strip()
                life_val = int(life) if life else None
                grp = e_grp.var.get().strip() or None
                if not sku or not name or not cat:
                    raise ValueError("SKU/Наименование/Категория обязательны.")
                self.conn.execute("""
                    INSERT INTO products(sku,name,category,uom,unit_cost,is_perishable,shelf_life_days,substitution_group,is_active)
                    VALUES(?,?,?,?,?,?,?,?,1)
                """, (sku, name, cat, uom, cost, per, life_val, grp))
                self.conn.commit()
                services.add_notification(self.conn, "INFO", f"Добавлен SKU {sku}.")
                self._refresh_catalog()
                win.destroy()
            except Exception as ex:
                messagebox.showerror("Ошибка", str(ex))
        ttk.Button(frm, text="Сохранить", command=save).pack(pady=10)
        ttk.Button(frm, text="Отмена", command=win.destroy).pack()

    def on_toggle_product(self):
        sel = self.tree_products.selection()
        if not sel:
            return
        item = self.tree_products.item(sel[0])["values"]
        pid = int(item[0])
        cur = self.conn.execute("SELECT is_active, sku FROM products WHERE id=?", (pid,)).fetchone()
        newv = 0 if int(cur["is_active"])==1 else 1
        self.conn.execute("UPDATE products SET is_active=? WHERE id=?", (newv, pid))
        self.conn.commit()
        services.add_notification(self.conn, "WARN", f"Изменена активность SKU {cur['sku']} => {newv}.")
        self._refresh_catalog()

    def on_export_products(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")], title="Экспорт номенклатуры")
        if not path:
            return
        rows = self.conn.execute("SELECT sku,name,category,uom,unit_cost,is_perishable,shelf_life_days,substitution_group,is_active FROM products ORDER BY sku").fetchall()
        import csv
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["sku","name","category","uom","unit_cost","is_perishable","shelf_life_days","substitution_group","is_active"])
            for r in rows:
                w.writerow([r["sku"], r["name"], r["category"], r["uom"], r["unit_cost"], r["is_perishable"], r["shelf_life_days"], r["substitution_group"], r["is_active"]])
        messagebox.showinfo("Готово", "Экспорт выполнен.")

    # ---------------- Stock ----------------
    def _build_stock(self):
        f = self.tab_stock
        top = ttk.Frame(f); top.pack(fill="x", padx=10, pady=10)

        self.wh_var = tk.StringVar(value="Все склады")
        whs = ["Все склады"] + [r["name"] for r in self.conn.execute("SELECT name FROM warehouses ORDER BY name").fetchall()]
        ttk.Label(top, text="Склад:").pack(side="left")
        self.wh_combo = ttk.Combobox(top, textvariable=self.wh_var, values=whs, state="readonly", width=22)
        self.wh_combo.pack(side="left", padx=6)
        ttk.Button(top, text="Показать", command=self._refresh_stock).pack(side="left", padx=6)
        ttk.Button(top, text="Добавить партию вручную", command=self.on_add_lot).pack(side="left", padx=6)

        cols = ["ID","SKU","Наименование","Склад","Партия","Годен до","Остаток","Резерв","Доступно"]
        self.tree_lots, vsb = make_tree(f, cols, widths=[60,120,330,150,150,110,100,100,100])
        self.tree_lots.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)
        vsb.pack(side="left", fill="y", pady=10)

    def _refresh_stock(self):
        for i in self.tree_lots.get_children():
            self.tree_lots.delete(i)
        wh = self.wh_var.get()
        if wh == "Все склады":
            rows = self.conn.execute("""
                SELECT sl.id, p.sku, p.name, w.name AS wh, sl.lot_no, sl.expiry_date,
                       sl.qty_on_hand, sl.qty_reserved, (sl.qty_on_hand-sl.qty_reserved) AS avail
                FROM stock_lots sl
                JOIN products p ON p.id=sl.product_id
                JOIN warehouses w ON w.id=sl.warehouse_id
                ORDER BY p.sku, (sl.expiry_date IS NULL) ASC, sl.expiry_date ASC, sl.created_at ASC
            """).fetchall()
        else:
            rows = self.conn.execute("""
                SELECT sl.id, p.sku, p.name, w.name AS wh, sl.lot_no, sl.expiry_date,
                       sl.qty_on_hand, sl.qty_reserved, (sl.qty_on_hand-sl.qty_reserved) AS avail
                FROM stock_lots sl
                JOIN products p ON p.id=sl.product_id
                JOIN warehouses w ON w.id=sl.warehouse_id
                WHERE w.name=?
                ORDER BY p.sku, (sl.expiry_date IS NULL) ASC, sl.expiry_date ASC, sl.created_at ASC
            """, (wh,)).fetchall()
        for r in rows:
            self.tree_lots.insert("", "end", values=(
                r["id"], r["sku"], r["name"], r["wh"], r["lot_no"], r["expiry_date"] or "",
                round(float(r["qty_on_hand"]),2), round(float(r["qty_reserved"]),2), round(float(r["avail"]),2)
            ))

    def on_add_lot(self):
        win = tk.Toplevel(self); win.title("Добавить партию"); win.geometry("620x420")
        frm = ttk.Frame(win); frm.pack(fill="both", expand=True, padx=12, pady=12)

        products = self.conn.execute("SELECT id, sku, name FROM products WHERE is_active=1 ORDER BY sku").fetchall()
        prod_map = {f"{p['sku']} — {p['name']}": int(p["id"]) for p in products}
        whs = self.conn.execute("SELECT id, name FROM warehouses ORDER BY name").fetchall()
        wh_map = {w["name"]: int(w["id"]) for w in whs}

        e_prod = LabeledCombo(frm, "Товар:", list(prod_map.keys()), width=42); e_prod.pack(fill="x")
        e_wh = LabeledCombo(frm, "Склад:", list(wh_map.keys()), width=42); e_wh.pack(fill="x")
        e_lot = LabeledEntry(frm, "Партия (lot_no):", width=42); e_lot.pack(fill="x")
        e_exp = LabeledEntry(frm, "Годен до (YYYY-MM-DD или пусто):", width=42); e_exp.pack(fill="x")
        e_qty = LabeledEntry(frm, "Количество:", width=42); e_qty.pack(fill="x"); e_qty.var.set("0")

        def save():
            try:
                pid = prod_map[e_prod.var.get()]
                wid = wh_map[e_wh.var.get()]
                lot = e_lot.var.get().strip()
                exp = e_exp.var.get().strip() or None
                qty = float(e_qty.var.get().strip() or "0")
                if not lot or qty <= 0:
                    raise ValueError("lot_no и количество обязательны (qty>0).")
                self.conn.execute("""
                    INSERT INTO stock_lots(product_id, warehouse_id, lot_no, expiry_date, qty_on_hand, qty_reserved, created_at)
                    VALUES(?,?,?,?,?,?,?)
                """, (pid, wid, lot, exp, qty, 0.0, datetime.datetime.now().isoformat(timespec="seconds")))
                self.conn.commit()
                services.add_notification(self.conn, "INFO", f"Добавлена партия {lot}.")
                self._refresh_stock()
                win.destroy()
            except Exception as ex:
                messagebox.showerror("Ошибка", str(ex))

        ttk.Button(frm, text="Сохранить", command=save).pack(pady=10)
        ttk.Button(frm, text="Отмена", command=win.destroy).pack()

    # ---------------- Orders / ATP ----------------
    def _build_orders(self):
        f = self.tab_orders
        top = ttk.Frame(f); top.pack(fill="x", padx=10, pady=10)

        ttk.Button(top, text="Создать заказ", command=self.on_create_order).pack(side="left")
        ttk.Button(top, text="Обновить", command=self._refresh_orders).pack(side="left", padx=6)

        cols = ["ID","Номер","Клиент","Канал","Статус","Создан","Обещ. дата"]
        self.tree_orders, vsb = make_tree(f, cols, widths=[60,160,220,90,110,170,110])
        self.tree_orders.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)
        vsb.pack(side="left", fill="y", pady=10)

        right = ttk.Frame(f); right.pack(side="left", fill="both", padx=10, pady=10)
        ttk.Label(right, text="Строки заказа:", font=("Times New Roman", 14, "bold")).pack(anchor="w")
        cols2 = ["LineID","SKU","Наименование","Кол-во","Резерв","Цена"]
        self.tree_order_lines, vsb2 = make_tree(right, cols2, widths=[70,120,260,90,90,90])
        self.tree_order_lines.pack(side="top", fill="both", expand=True)
        vsb2.pack(side="top", fill="y")

        btns = ttk.Frame(right); btns.pack(fill="x", pady=6)
        ttk.Button(btns, text="ATP проверка", command=self.on_atp_check).pack(side="left")
        ttk.Button(btns, text="Резервировать (FEFO)", command=self.on_reserve_line).pack(side="left", padx=6)
        ttk.Button(btns, text="Подбор замен", command=self.on_substitutions).pack(side="left", padx=6)

        self.tree_orders.bind("<<TreeviewSelect>>", lambda e: self._refresh_order_lines())

    def _refresh_orders(self):
        for i in self.tree_orders.get_children():
            self.tree_orders.delete(i)
        rows = self.conn.execute("""
            SELECT id, order_no, customer, channel, status, created_at, promised_date
            FROM sales_orders
            ORDER BY created_at DESC
            LIMIT 200
        """).fetchall()
        for r in rows:
            self.tree_orders.insert("", "end", values=(r["id"], r["order_no"], r["customer"], r["channel"], r["status"], r["created_at"], r["promised_date"] or ""))
        self._refresh_order_lines()

    def _selected_order_id(self):
        sel = self.tree_orders.selection()
        if not sel:
            return None
        return int(self.tree_orders.item(sel[0])["values"][0])

    def _refresh_order_lines(self):
        for i in self.tree_order_lines.get_children():
            self.tree_order_lines.delete(i)
        oid = self._selected_order_id()
        if not oid:
            return
        rows = self.conn.execute("""
            SELECT sol.id AS line_id, p.sku, p.name, sol.qty, sol.qty_reserved, sol.price
            FROM sales_order_lines sol
            JOIN products p ON p.id=sol.product_id
            WHERE sol.order_id=?
            ORDER BY sol.id
        """, (oid,)).fetchall()
        for r in rows:
            self.tree_order_lines.insert("", "end", values=(r["line_id"], r["sku"], r["name"], r["qty"], r["qty_reserved"], r["price"]))

    def on_create_order(self):
        win = tk.Toplevel(self); win.title("Создать заказ"); win.geometry("720x520")
        frm = ttk.Frame(win); frm.pack(fill="both", expand=True, padx=12, pady=12)

        e_cust = LabeledEntry(frm, "Клиент:", width=46); e_cust.pack(fill="x")
        e_chan = LabeledCombo(frm, "Канал:", ["B2B","B2C","OMNI"], width=18); e_chan.var.set("B2B"); e_chan.pack(fill="x")
        e_prom = LabeledEntry(frm, "Обещанная дата (YYYY-MM-DD, опционально):", width=18); e_prom.pack(fill="x")

        products = self.conn.execute("SELECT id, sku, name, unit_cost FROM products WHERE is_active=1 ORDER BY sku").fetchall()
        prod_map = {f"{p['sku']} — {p['name']}": (int(p["id"]), float(p["unit_cost"])) for p in products}

        ttk.Separator(frm).pack(fill="x", pady=8)
        ttk.Label(frm, text="Строки заказа:", font=("Times New Roman", 14, "bold")).pack(anchor="w", pady=(0,4))

        lines = []

        cols = ["SKU","Кол-во","Цена"]
        tree, vsb = make_tree(frm, cols, widths=[420, 100, 100])
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")

        addf = ttk.Frame(frm); addf.pack(fill="x", pady=8)
        e_prod = LabeledCombo(addf, "Товар:", list(prod_map.keys()), width=42); e_prod.pack(fill="x")
        e_qty = LabeledEntry(addf, "Кол-во:", width=12); e_qty.var.set("1"); e_qty.pack(fill="x")
        e_price = LabeledEntry(addf, "Цена (руб.):", width=12); e_price.var.set("0"); e_price.pack(fill="x")

        def add_line():
            try:
                key = e_prod.var.get()
                pid, cost = prod_map[key]
                qty = float(e_qty.var.get().strip() or "0")
                price = float(e_price.var.get().strip() or str(cost))
                if qty <= 0:
                    raise ValueError("Кол-во должно быть >0.")
                lines.append((pid, qty, price, key.split(" — ")[0]))
                tree.insert("", "end", values=(key, qty, price))
            except Exception as ex:
                messagebox.showerror("Ошибка", str(ex))

        ttk.Button(addf, text="Добавить строку", command=add_line).pack(pady=6)

        def save_order():
            try:
                cust = e_cust.var.get().strip() or "Клиент"
                chan = e_chan.var.get().strip() or "B2B"
                prom = e_prom.var.get().strip() or None
                oid = services.create_sales_order(self.conn, cust, chan, prom)
                for pid, qty, price, _sku in lines:
                    services.add_sales_order_line(self.conn, oid, pid, qty, price)
                services.add_notification(self.conn, "INFO", f"Создан заказ SO id={oid} со строками: {len(lines)}.")
                self._refresh_orders()
                win.destroy()
            except Exception as ex:
                messagebox.showerror("Ошибка", str(ex))

        bot = ttk.Frame(frm); bot.pack(fill="x", pady=8)
        ttk.Button(bot, text="Сохранить заказ", command=save_order).pack(side="left")
        ttk.Button(bot, text="Закрыть", command=win.destroy).pack(side="left", padx=6)

    def _selected_line_id(self):
        sel = self.tree_order_lines.selection()
        if not sel:
            return None
        return int(self.tree_order_lines.item(sel[0])["values"][0])

    def on_atp_check(self):
        lid = self._selected_line_id()
        if not lid:
            return
        line = self.conn.execute("""
            SELECT sol.product_id, sol.qty, sol.qty_reserved, p.sku, p.name
            FROM sales_order_lines sol
            JOIN products p ON p.id=sol.product_id
            WHERE sol.id=?
        """, (lid,)).fetchone()
        need = float(line["qty"]) - float(line["qty_reserved"])
        ok, avail = services.atp_check(self.conn, int(line["product_id"]), need)
        msg = f"SKU {line['sku']} — {line['name']}\nНужно к резерву: {need:.2f}\nДоступно: {avail:.2f}\nРезультат: {'OK' if ok else 'НЕ ХВАТАЕТ'}"
        show_info(self, "ATP проверка", msg)

    def on_reserve_line(self):
        lid = self._selected_line_id()
        if not lid:
            return
        res = services.reserve_stock_for_order_line(self.conn, lid)
        if res.get("ok"):
            services.add_notification(self.conn, "INFO", res.get("message",""))
            self._refresh_orders()
            self._refresh_stock()
            self._refresh_dashboard()
        else:
            messagebox.showwarning("Резерв", res.get("message",""))

    def on_substitutions(self):
        lid = self._selected_line_id()
        if not lid:
            return
        line = self.conn.execute("""
            SELECT sol.product_id, sol.qty, sol.qty_reserved, p.sku, p.name
            FROM sales_order_lines sol
            JOIN products p ON p.id=sol.product_id
            WHERE sol.id=?
        """, (lid,)).fetchone()
        need = float(line["qty"]) - float(line["qty_reserved"])
        subs = services.suggest_substitutions(self.conn, int(line["product_id"]), need)
        if not subs:
            show_info(self, "Подбор замен", "Подходящих замен не найдено (нет группы замен или нет доступного остатка).")
            return
        txt = "Возможные замены (по группе замен):\n\n"
        for s in subs[:10]:
            txt += f"{s['sku']} — {s['name']}; доступно: {s['available']}\n"
        show_info(self, "Подбор замен", txt)

    # ---------------- Replenishment ----------------
    def _build_replen(self):
        f = self.tab_replen
        top = ttk.Frame(f); top.pack(fill="x", padx=10, pady=10)
        ttk.Button(top, text="Сформировать план пополнения", command=self._refresh_plan).pack(side="left")
        ttk.Button(top, text="Создать PO по плану (выбранные)", command=self.on_create_po_from_plan).pack(side="left", padx=6)

        cols = ["SKU","Наименование","Поставщик","Доступно","Min","Max","К заказу","ABC","XYZ"]
        self.tree_plan, vsb = make_tree(f, cols, widths=[120,330,200,90,90,90,90,60,60])
        self.tree_plan.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)
        vsb.pack(side="left", fill="y", pady=10)

        right = ttk.Frame(f); right.pack(side="left", fill="y", padx=10, pady=10)
        ttk.Button(right, text="Экспорт плана в CSV", command=self.on_export_plan).pack(fill="x", pady=3)
        ttk.Button(right, text="Создать задачи по OOS/низким остаткам", command=self.on_create_tasks_by_plan).pack(fill="x", pady=3)
        ttk.Separator(right).pack(fill="x", pady=10)
        ttk.Label(right, text="Подсказка:", font=("Times New Roman", 14, "bold")).pack(anchor="w")
        ttk.Label(right, text="План рассчитывается по последним параметрам\n(min/max, ROP, safety stock).\nЕсли «К заказу» > 0 —\nдоступный остаток ниже min.").pack(anchor="w")

    def _refresh_plan(self):
        self.plan = services.build_replenishment_plan(self.conn)
        for i in self.tree_plan.get_children():
            self.tree_plan.delete(i)
        for r in self.plan:
            self.tree_plan.insert("", "end", values=(r["sku"], r["name"], r["supplier"], r["available"], r["min_level"], r["max_level"], r["qty_to_order"], r["abc"], r["xyz"]))

    def on_export_plan(self):
        if not hasattr(self, "plan"):
            self._refresh_plan()
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")], title="Экспорт плана пополнения")
        if not path:
            return
        import csv
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["sku","name","supplier","available","min","max","qty_to_order","abc","xyz"])
            for r in self.plan:
                w.writerow([r["sku"], r["name"], r["supplier"], r["available"], r["min_level"], r["max_level"], r["qty_to_order"], r["abc"], r["xyz"]])
        messagebox.showinfo("Готово", "Экспорт выполнен.")

    def on_create_tasks_by_plan(self):
        if not hasattr(self, "plan"):
            self._refresh_plan()
        created = 0
        # Assign to buyer by default
        buyer = self.conn.execute("SELECT id FROM users WHERE role='PROCUREMENT' ORDER BY id LIMIT 1").fetchone()
        buyer_id = int(buyer["id"]) if buyer else None
        for r in self.plan:
            if float(r["qty_to_order"]) > 0:
                services.create_task(self.conn, f"Проверить пополнение: {r['sku']} ({r['qty_to_order']})", "REPLENISHMENT", priority="HIGH", assigned_to=buyer_id, notes="Сформировано автоматически по плану пополнения.")
                created += 1
        services.add_notification(self.conn, "INFO", f"Созданы задачи по плану пополнения: {created}.")
        self._refresh_tasks()

    def on_create_po_from_plan(self):
        if not hasattr(self, "plan"):
            self._refresh_plan()
        # Group by supplier using latest params
        sel = self.tree_plan.selection()
        if not sel:
            messagebox.showinfo("PO", "Выбери строки плана (Ctrl/Shift) и повтори.")
            return
        # map sku -> product_id and supplier
        by_supplier = {}
        for s in sel:
            vals = self.tree_plan.item(s)["values"]
            sku = vals[0]
            qty = float(vals[6])
            if qty <= 0:
                continue
            row = self.conn.execute("""
                SELECT p.id AS product_id,
                       rp.supplier_id, s.name AS supplier_name,
                       p.unit_cost
                FROM products p
                JOIN (
                  SELECT product_id, MAX(created_at) AS mx FROM replenishment_params GROUP BY product_id
                ) x ON x.product_id=p.id
                JOIN replenishment_params rp ON rp.product_id=p.id AND rp.created_at=x.mx
                LEFT JOIN suppliers s ON s.id=rp.supplier_id
                WHERE p.sku=?
            """, (sku,)).fetchone()
            if not row or row["supplier_id"] is None:
                continue
            sid = int(row["supplier_id"])
            by_supplier.setdefault(sid, []).append((int(row["product_id"]), qty, float(row["unit_cost"])))

        if not by_supplier:
            messagebox.showwarning("PO", "Нет строк с qty_to_order > 0 или не определён поставщик.")
            return

        created = 0
        for sid, lines in by_supplier.items():
            po_id = services.create_purchase_order(self.conn, sid)
            for pid, qty, cost in lines:
                services.add_purchase_order_line(self.conn, po_id, pid, qty, cost)
            # mark as SENT to emulate
            self.conn.execute("UPDATE purchase_orders SET status='SENT' WHERE id=?", (po_id,))
            self.conn.commit()
            created += 1
        services.add_notification(self.conn, "INFO", f"Созданы PO по плану (кол-во документов: {created}).")
        self._refresh_po()

    # ---------------- Purchase Orders ----------------
    def _build_po(self):
        f = self.tab_po
        top = ttk.Frame(f); top.pack(fill="x", padx=10, pady=10)
        ttk.Button(top, text="Создать PO вручную", command=self.on_create_po_manual).pack(side="left")
        ttk.Button(top, text="Обновить", command=self._refresh_po).pack(side="left", padx=6)
        ttk.Button(top, text="Выполнить приёмку PO", command=self.on_receive_po).pack(side="left", padx=6)

        cols = ["ID","Номер","Поставщик","Статус","Создан","ETA"]
        self.tree_po, vsb = make_tree(f, cols, widths=[60,170,260,110,170,110])
        self.tree_po.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)
        vsb.pack(side="left", fill="y", pady=10)

        right = ttk.Frame(f); right.pack(side="left", fill="both", padx=10, pady=10)
        ttk.Label(right, text="Строки PO:", font=("Times New Roman", 14, "bold")).pack(anchor="w")
        cols2 = ["LineID","SKU","Наименование","Кол-во","Себестоимость"]
        self.tree_po_lines, vsb2 = make_tree(right, cols2, widths=[70,120,260,90,110])
        self.tree_po_lines.pack(side="top", fill="both", expand=True)
        vsb2.pack(side="top", fill="y")

        self.tree_po.bind("<<TreeviewSelect>>", lambda e: self._refresh_po_lines())

    def _refresh_po(self):
        for i in self.tree_po.get_children():
            self.tree_po.delete(i)
        rows = self.conn.execute("""
            SELECT po.id, po.po_no, s.name AS supplier, po.status, po.created_at, po.expected_date
            FROM purchase_orders po
            JOIN suppliers s ON s.id=po.supplier_id
            ORDER BY po.created_at DESC
            LIMIT 200
        """).fetchall()
        for r in rows:
            self.tree_po.insert("", "end", values=(r["id"], r["po_no"], r["supplier"], r["status"], r["created_at"], r["expected_date"] or ""))
        self._refresh_po_lines()

    def _selected_po_id(self):
        sel = self.tree_po.selection()
        if not sel:
            return None
        return int(self.tree_po.item(sel[0])["values"][0])

    def _refresh_po_lines(self):
        for i in self.tree_po_lines.get_children():
            self.tree_po_lines.delete(i)
        po_id = self._selected_po_id()
        if not po_id:
            return
        rows = self.conn.execute("""
            SELECT pol.id AS line_id, p.sku, p.name, pol.qty, pol.unit_cost
            FROM purchase_order_lines pol
            JOIN products p ON p.id=pol.product_id
            WHERE pol.po_id=?
            ORDER BY pol.id
        """, (po_id,)).fetchall()
        for r in rows:
            self.tree_po_lines.insert("", "end", values=(r["line_id"], r["sku"], r["name"], r["qty"], r["unit_cost"]))

    def on_create_po_manual(self):
        win = tk.Toplevel(self); win.title("Создать PO"); win.geometry("760x560")
        frm = ttk.Frame(win); frm.pack(fill="both", expand=True, padx=12, pady=12)

        suppliers = self.conn.execute("SELECT id, name FROM suppliers ORDER BY name").fetchall()
        sup_map = {s["name"]: int(s["id"]) for s in suppliers}

        e_sup = LabeledCombo(frm, "Поставщик:", list(sup_map.keys()), width=44); e_sup.pack(fill="x")
        e_eta = LabeledEntry(frm, "ETA (YYYY-MM-DD, опционально):", width=18); e_eta.pack(fill="x")

        products = self.conn.execute("SELECT id, sku, name, unit_cost FROM products WHERE is_active=1 ORDER BY sku").fetchall()
        prod_map = {f"{p['sku']} — {p['name']}": (int(p["id"]), float(p["unit_cost"])) for p in products}

        ttk.Separator(frm).pack(fill="x", pady=8)
        ttk.Label(frm, text="Строки PO:", font=("Times New Roman", 14, "bold")).pack(anchor="w", pady=(0,4))

        cols = ["SKU","Кол-во","Себестоимость"]
        tree, vsb = make_tree(frm, cols, widths=[460, 100, 120])
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")

        addf = ttk.Frame(frm); addf.pack(fill="x", pady=8)
        e_prod = LabeledCombo(addf, "Товар:", list(prod_map.keys()), width=44); e_prod.pack(fill="x")
        e_qty = LabeledEntry(addf, "Кол-во:", width=12); e_qty.var.set("1"); e_qty.pack(fill="x")
        e_cost = LabeledEntry(addf, "Себестоимость:", width=12); e_cost.var.set("0"); e_cost.pack(fill="x")

        lines = []

        def add_line():
            try:
                pid, base_cost = prod_map[e_prod.var.get()]
                qty = float(e_qty.var.get().strip() or "0")
                cost = float(e_cost.var.get().strip() or str(base_cost))
                if qty <= 0:
                    raise ValueError("Кол-во должно быть >0.")
                lines.append((pid, qty, cost, e_prod.var.get()))
                tree.insert("", "end", values=(e_prod.var.get(), qty, cost))
            except Exception as ex:
                messagebox.showerror("Ошибка", str(ex))

        ttk.Button(addf, text="Добавить строку", command=add_line).pack(pady=6)

        def save_po():
            try:
                sid = sup_map[e_sup.var.get()]
                eta = e_eta.var.get().strip() or None
                po_id = services.create_purchase_order(self.conn, sid, eta)
                for pid, qty, cost, _ in lines:
                    services.add_purchase_order_line(self.conn, po_id, pid, qty, cost)
                self.conn.execute("UPDATE purchase_orders SET status='SENT' WHERE id=?", (po_id,))
                self.conn.commit()
                services.add_notification(self.conn, "INFO", f"Создан PO id={po_id}.")
                self._refresh_po()
                win.destroy()
            except Exception as ex:
                messagebox.showerror("Ошибка", str(ex))

        bot = ttk.Frame(frm); bot.pack(fill="x", pady=8)
        ttk.Button(bot, text="Сохранить PO", command=save_po).pack(side="left")
        ttk.Button(bot, text="Закрыть", command=win.destroy).pack(side="left", padx=6)

    def on_receive_po(self):
        po_id = self._selected_po_id()
        if not po_id:
            return
        po = self.conn.execute("SELECT status, po_no FROM purchase_orders WHERE id=?", (po_id,)).fetchone()
        if not po or po["status"] in ("RECEIVED","CANCELLED"):
            messagebox.showinfo("Приёмка", "PO уже получен или отменён.")
            return

        win = tk.Toplevel(self); win.title("Приёмка PO"); win.geometry("780x560")
        frm = ttk.Frame(win); frm.pack(fill="both", expand=True, padx=12, pady=12)

        whs = self.conn.execute("SELECT id, name FROM warehouses ORDER BY name").fetchall()
        wh_map = {w["name"]: int(w["id"]) for w in whs}
        e_wh = LabeledCombo(frm, "Склад приёмки:", list(wh_map.keys()), width=44); e_wh.pack(fill="x")

        lines = self.conn.execute("""
            SELECT pol.product_id, p.sku, p.name, pol.qty, p.is_perishable, p.shelf_life_days
            FROM purchase_order_lines pol
            JOIN products p ON p.id=pol.product_id
            WHERE pol.po_id=?
            ORDER BY p.sku
        """, (po_id,)).fetchall()

        cols = ["SKU","Наименование","Кол-во PO","Lot No","Годен до (опц.)","Принято"]
        tree, vsb = make_tree(frm, cols, widths=[120,260,90,140,140,90])
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")

        entries = []
        today = datetime.date.today()
        for ln in lines:
            lot = f"RCV-{po_id}-{ln['sku']}-{today.strftime('%m%d')}"
            exp = ""
            if int(ln["is_perishable"])==1 and ln["shelf_life_days"]:
                exp = (today + datetime.timedelta(days=int(ln["shelf_life_days"]))).isoformat()
            tree.insert("", "end", values=(ln["sku"], ln["name"], ln["qty"], lot, exp, ln["qty"]))
            entries.append({"product_id": int(ln["product_id"]), "lot_no": lot, "expiry_date": exp or None, "qty": float(ln["qty"])})

        def do_receive():
            try:
                wid = wh_map[e_wh.var.get()]
                services.receive_purchase_order(self.conn, po_id, wid, entries)
                services.add_notification(self.conn, "INFO", f"Приёмка выполнена по {po['po_no']}.")
                self._refresh_po()
                self._refresh_stock()
                self._refresh_dashboard()
                win.destroy()
            except Exception as ex:
                messagebox.showerror("Ошибка", str(ex))

        ttk.Button(frm, text="Выполнить приёмку", command=do_receive).pack(pady=8)
        ttk.Button(frm, text="Закрыть", command=win.destroy).pack()

    # ---------------- Tasks / Notifications ----------------
    def _build_tasks(self):
        f = self.tab_tasks
        top = ttk.Frame(f); top.pack(fill="x", padx=10, pady=10)
        ttk.Button(top, text="Создать задачу", command=self.on_create_task_manual).pack(side="left")
        ttk.Button(top, text="Обновить", command=self._refresh_tasks).pack(side="left", padx=6)
        ttk.Button(top, text="Пометить уведомления прочитанными", command=self.on_mark_notifications_read).pack(side="right")

        mid = ttk.Panedwindow(f, orient="horizontal")
        mid.pack(fill="both", expand=True, padx=10, pady=10)

        left = ttk.Frame(mid); right = ttk.Frame(mid)
        mid.add(left, weight=2); mid.add(right, weight=1)

        ttk.Label(left, text="Задачи:", font=("Times New Roman", 14, "bold")).pack(anchor="w")
        cols = ["ID","Заголовок","Тип","Приоритет","Статус","Назначен","Создан","Срок"]
        self.tree_tasks, vsb = make_tree(left, cols, widths=[60,380,140,90,100,120,170,110])
        self.tree_tasks.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")
        btn = ttk.Frame(left); btn.pack(fill="x", pady=6)
        ttk.Button(btn, text="В работу", command=lambda: self.on_set_task_status("IN_PROGRESS")).pack(side="left")
        ttk.Button(btn, text="Завершить", command=lambda: self.on_set_task_status("DONE")).pack(side="left", padx=6)

        ttk.Label(right, text="Уведомления:", font=("Times New Roman", 14, "bold")).pack(anchor="w")
        cols2 = ["ID","Время","Уровень","Сообщение","Проч."]
        self.tree_notif, vsb2 = make_tree(right, cols2, widths=[50,150,70,360,60])
        self.tree_notif.pack(side="left", fill="both", expand=True)
        vsb2.pack(side="left", fill="y")

    def _refresh_tasks(self):
        for i in self.tree_tasks.get_children():
            self.tree_tasks.delete(i)
        rows = self.conn.execute("""
            SELECT t.id, t.title, t.task_type, t.priority, t.status,
                   COALESCE(u.username,'') AS assigned,
                   t.created_at, t.due_date
            FROM tasks t
            LEFT JOIN users u ON u.id=t.assigned_to
            ORDER BY CASE t.priority WHEN 'CRITICAL' THEN 1 WHEN 'HIGH' THEN 2 WHEN 'MEDIUM' THEN 3 ELSE 4 END,
                     t.created_at DESC
            LIMIT 300
        """).fetchall()
        for r in rows:
            self.tree_tasks.insert("", "end", values=(r["id"], r["title"], r["task_type"], r["priority"], r["status"], r["assigned"], r["created_at"], r["due_date"] or ""))

        for i in self.tree_notif.get_children():
            self.tree_notif.delete(i)
        nrows = self.conn.execute("""
            SELECT id, created_at, level, message, is_read
            FROM notifications
            ORDER BY created_at DESC
            LIMIT 200
        """).fetchall()
        for r in nrows:
            self.tree_notif.insert("", "end", values=(r["id"], r["created_at"], r["level"], r["message"], "да" if int(r["is_read"])==1 else "нет"))

    def on_create_task_manual(self):
        win = tk.Toplevel(self); win.title("Создать задачу"); win.geometry("560x420")
        frm = ttk.Frame(win); frm.pack(fill="both", expand=True, padx=12, pady=12)

        e_title = LabeledEntry(frm, "Заголовок:", width=44); e_title.pack(fill="x")
        e_type = LabeledEntry(frm, "Тип (например REPLENISHMENT):", width=20); e_type.pack(fill="x")
        e_pri = LabeledCombo(frm, "Приоритет:", ["LOW","MEDIUM","HIGH","CRITICAL"], width=18); e_pri.var.set("MEDIUM"); e_pri.pack(fill="x")
        e_due = LabeledEntry(frm, "Срок (YYYY-MM-DD, опционально):", width=18); e_due.pack(fill="x")

        users = self.conn.execute("SELECT id, username FROM users ORDER BY username").fetchall()
        u_map = {"": None, **{u["username"]: int(u["id"]) for u in users}}
        e_user = LabeledCombo(frm, "Исполнитель:", list(u_map.keys()), width=18); e_user.var.set(""); e_user.pack(fill="x")

        txt = tk.Text(frm, height=6, wrap="word")
        ttk.Label(frm, text="Комментарий:").pack(anchor="w", pady=(8,2))
        txt.pack(fill="both", expand=True)

        def save():
            try:
                title = e_title.var.get().strip()
                ttype = e_type.var.get().strip() or "GENERAL"
                pri = e_pri.var.get().strip()
                due = e_due.var.get().strip() or None
                assigned = u_map.get(e_user.var.get())
                notes = txt.get("1.0", "end").strip()
                if not title:
                    raise ValueError("Заголовок обязателен.")
                services.create_task(self.conn, title, ttype, priority=pri, assigned_to=assigned, due_date=due, notes=notes)
                services.add_notification(self.conn, "INFO", f"Создана задача: {title}.")
                self._refresh_tasks()
                win.destroy()
            except Exception as ex:
                messagebox.showerror("Ошибка", str(ex))

        ttk.Button(frm, text="Сохранить", command=save).pack(pady=10)
        ttk.Button(frm, text="Закрыть", command=win.destroy).pack()

    def on_set_task_status(self, status):
        sel = self.tree_tasks.selection()
        if not sel:
            return
        tid = int(self.tree_tasks.item(sel[0])["values"][0])
        self.conn.execute("UPDATE tasks SET status=? WHERE id=?", (status, tid))
        self.conn.commit()
        services.add_notification(self.conn, "INFO", f"Изменен статус задачи {tid} => {status}.")
        self._refresh_tasks()

    def on_mark_notifications_read(self):
        self.conn.execute("UPDATE notifications SET is_read=1 WHERE is_read=0")
        self.conn.commit()
        self._refresh_tasks()

    # ---------------- Integration ----------------
    def _build_integration(self):
        f = self.tab_integration
        top = ttk.Frame(f); top.pack(fill="x", padx=10, pady=10)
        ttk.Button(top, text="Сгенерировать сообщение ETA (пример)", command=self.on_generate_eta_message).pack(side="left")
        ttk.Button(top, text="Обработать выбранное", command=self.on_process_message).pack(side="left", padx=6)
        ttk.Button(top, text="Обновить", command=self._refresh_inbox).pack(side="left", padx=6)

        cols = ["ID","Источник","Время","Обработано","Payload"]
        self.tree_inbox, vsb = make_tree(f, cols, widths=[60,120,160,90,760])
        self.tree_inbox.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)
        vsb.pack(side="left", fill="y", pady=10)

    def _refresh_inbox(self):
        for i in self.tree_inbox.get_children():
            self.tree_inbox.delete(i)
        rows = self.conn.execute("""
            SELECT id, source, received_at, processed, payload_json
            FROM integration_inbox
            ORDER BY received_at DESC
            LIMIT 200
        """).fetchall()
        for r in rows:
            payload = r["payload_json"]
            if len(payload) > 220:
                payload = payload[:220] + "..."
            self.tree_inbox.insert("", "end", values=(r["id"], r["source"], r["received_at"], "да" if int(r["processed"])==1 else "нет", payload))

    def on_generate_eta_message(self):
        # pick any SENT PO
        po = self.conn.execute("SELECT po_no FROM purchase_orders WHERE status='SENT' ORDER BY created_at DESC LIMIT 1").fetchone()
        if not po:
            messagebox.showinfo("Интеграции", "Нет отправленных PO (status=SENT). Создайте PO и повторите.")
            return
        eta = (datetime.date.today() + datetime.timedelta(days=7)).isoformat()
        payload = {"type":"supplier_eta_update", "po_no": po["po_no"], "expected_date": eta}
        services.push_integration_message(self.conn, "EDI", payload)
        services.add_notification(self.conn, "INFO", f"Добавлено входящее сообщение ETA для {po['po_no']}.")
        self._refresh_inbox()
        self._refresh_tasks()

    def on_process_message(self):
        sel = self.tree_inbox.selection()
        if not sel:
            return
        msg_id = int(self.tree_inbox.item(sel[0])["values"][0])
        res = services.process_integration_message(self.conn, msg_id)
        messagebox.showinfo("Результат", res)
        self._refresh_inbox()
        self._refresh_po()
        self._refresh_tasks()

    # ---------------- Settings ----------------
    def _build_settings(self):
        f = self.tab_settings
        frm = ttk.Frame(f); frm.pack(fill="both", expand=True, padx=10, pady=10)

        self.set_sl = LabeledEntry(frm, "Service level (например 0.95):", width=12); self.set_sl.pack(fill="x")
        self.set_win = LabeledEntry(frm, "Окно истории спроса (дней):", width=12); self.set_win.pack(fill="x")
        self.set_hor = LabeledEntry(frm, "Горизонт прогноза (дней):", width=12); self.set_hor.pack(fill="x")
        self.set_rev = LabeledEntry(frm, "Период обзора (дней):", width=12); self.set_rev.pack(fill="x")
        self.set_method = LabeledCombo(frm, "Метод прогноза:", ["ES","MA","CROSTON"], width=12); self.set_method.pack(fill="x")

        ttk.Button(frm, text="Сохранить настройки", command=self.on_save_settings).pack(pady=10)
        ttk.Button(frm, text="Справка по прототипу", command=self.on_help).pack()

    def _refresh_settings(self):
        self.set_sl.var.set(services.get_setting(self.conn, "default_service_level", "0.95"))
        self.set_win.var.set(services.get_setting(self.conn, "forecast_window_days", "90"))
        self.set_hor.var.set(services.get_setting(self.conn, "forecast_horizon_days", "14"))
        self.set_rev.var.set(services.get_setting(self.conn, "review_period_days", "7"))
        self.set_method.var.set(services.get_setting(self.conn, "default_forecast_method", "ES"))

    def on_save_settings(self):
        try:
            sl = float(self.set_sl.var.get().strip() or "0.95")
            if sl < 0.5 or sl >= 1.0:
                raise ValueError("Service level должен быть в диапазоне [0.5; 1.0).")
            services.set_setting(self.conn, "default_service_level", str(sl))
            services.set_setting(self.conn, "forecast_window_days", str(int(self.set_win.var.get().strip() or "90")))
            services.set_setting(self.conn, "forecast_horizon_days", str(int(self.set_hor.var.get().strip() or "14")))
            services.set_setting(self.conn, "review_period_days", str(int(self.set_rev.var.get().strip() or "7")))
            services.set_setting(self.conn, "default_forecast_method", self.set_method.var.get().strip() or "ES")
            services.add_notification(self.conn, "INFO", "Настройки сохранены.")
            messagebox.showinfo("Готово", "Настройки сохранены. Пересчитайте параметры пополнения.")
        except Exception as ex:
            messagebox.showerror("Ошибка", str(ex))

    def on_help(self):
        msg = (
            "Прототип реализует ключевые требования ТЗ:\n"
            "1) Прогноз спроса (ES/MA/Croston) и параметры пополнения (ROP, safety stock, min/max)\n"
            "2) ABC/XYZ классификация\n"
            "3) ATP-проверка и резервирование по FEFO/FIFO\n"
            "4) Партии/сроки годности (лотирование)\n"
            "5) Задачи/уведомления (workflow)\n"
            "6) Интеграции (эмуляция входящих сообщений EDI/API)\n"
            "7) KPI: OOS/OSA, оборачиваемость (оценка), неликвиды\n\n"
            "Для диплома: в Главе 2 можно описывать окна приложения,\n"
            "приводить скриншоты и примеры SQL/логики расчетов."
        )
        show_info(self, "Справка", msg)

    # ---------------- Global refresh/actions ----------------
    def refresh_all(self):
        self._refresh_settings()
        self._refresh_dashboard()
        self._refresh_catalog()
        self._refresh_stock()
        self._refresh_orders()
        self._refresh_plan()
        self._refresh_po()
        self._refresh_tasks()
        self._refresh_inbox()

    def on_recompute(self):
        services.recompute_all_parameters(self.conn)
        self._refresh_plan()
        self._refresh_dashboard()
        self._refresh_tasks()

if __name__ == "__main__":
    app = App()
    app.mainloop()
