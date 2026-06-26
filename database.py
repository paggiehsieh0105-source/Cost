"""
階段3：資料儲存與持久化 (SQLite)

資料表設計對應 CLAUDE.md 第六節：
    projects, packaging_materials, project_packaging,
    project_labor, project_subcontract, project_shipping, settings

提供 CRUD 函式，並能把 CostProject 物件存入/讀出資料庫，
讓 cost_model.calculate_total_cost() 可以直接拿讀出來的資料計算。
"""

import sqlite3
from contextlib import contextmanager
from typing import Optional

from cost_model import (
    Settings, PackagingMaterial, PackagingCatalog, CostProject,
    RawMaterial, PackagingLine, LaborLine, ShippingLine, CartonLine,
)

SCHEMA = """
CREATE TABLE IF NOT EXISTS settings (
    id INTEGER PRIMARY KEY CHECK (id = 1),      -- 全廠只有一筆共用設定
    margin_mode TEXT NOT NULL DEFAULT '成本基礎(加成法)',
    margin_rate REAL NOT NULL DEFAULT 0.30,
    overhead_rate REAL NOT NULL DEFAULT 0.15,
    tax_rate REAL NOT NULL DEFAULT 0.05,
    tax_mode TEXT NOT NULL DEFAULT '未稅',
    monthly_salary REAL NOT NULL DEFAULT 30000,
    work_days_per_month INTEGER NOT NULL DEFAULT 22,
    work_hours_per_day REAL NOT NULL DEFAULT 8,
    logistics_rate REAL NOT NULL DEFAULT 150,
    pallet_rate REAL NOT NULL DEFAULT 280,
    tier_margin_low REAL NOT NULL DEFAULT 0.25,
    tier_margin_high REAL NOT NULL DEFAULT 0.35,
    raw_material_loss_rate REAL NOT NULL DEFAULT 0,
    labor_loss_rate REAL NOT NULL DEFAULT 0,
    shipping_loss_rate REAL NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS shipping_regions (
    region_name TEXT PRIMARY KEY,
    unit_rate REAL NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS carton_sizes (
    size_name TEXT PRIMARY KEY,
    unit_rate REAL NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS packaging_materials (
    material_id TEXT PRIMARY KEY,
    material_name TEXT NOT NULL,
    spec TEXT,
    unit TEXT,
    unit_price REAL NOT NULL DEFAULT 0,
    vendor TEXT,
    note TEXT,
    updated_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS projects (
    project_id TEXT PRIMARY KEY,
    customer_name TEXT,
    product_name TEXT,
    monthly_quantity REAL DEFAULT 0,
    raw_material_qty REAL DEFAULT 1,
    raw_material_unit_price REAL DEFAULT 0,
    raw_material_loss_rate REAL DEFAULT 0,
    shipping_order_quantity REAL,
    tier_price_low REAL,
    tier_price_mid REAL,
    tier_price_high REAL,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS project_packaging (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
    material_id TEXT NOT NULL REFERENCES packaging_materials(material_id),
    quantity REAL NOT NULL DEFAULT 1,
    loss_rate REAL NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS project_labor (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
    process_name TEXT NOT NULL,
    batch_quantity REAL NOT NULL,
    batch_time_minutes REAL NOT NULL,
    headcount INTEGER NOT NULL DEFAULT 1,
    hourly_wage_override REAL
);

CREATE TABLE IF NOT EXISTS project_shipping (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
    method TEXT NOT NULL,             -- 專車 / 物流 / 棧板
    quantity REAL NOT NULL,
    region TEXT                       -- 僅 method='專車' 時使用
);

CREATE TABLE IF NOT EXISTS project_carton (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
    size TEXT NOT NULL REFERENCES carton_sizes(size_name),
    quantity REAL NOT NULL
);
"""

DEFAULT_REGIONS = ["中彰投雲林", "竹苗嘉", "桃園台北台南高雄", "宜蘭屏東", "台東花蓮"]
DEFAULT_CARTON_SIZES = {
    "小箱(20x20x20cm)": 15,
    "中箱(30x30x30cm)": 25,
    "大箱(40x40x40cm)": 35,
}


@contextmanager
def get_connection(db_path: str = "cost_calculator.db"):
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON")
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db(db_path: str = "cost_calculator.db") -> None:
    """建立所有資料表；若不存在才建立，並寫入預設設定與專車區域。
    若資料庫是用舊版程式建立的（缺少新欄位），會自動補上欄位，不需刪除重建。

    重要：「專車區域」「紙箱尺寸」的預設值，只會在該資料表『完全是空的』時才寫入
    （也就是只有第一次建資料庫那一刻才會發生）。如果你之後手動刪除了某個區域/尺寸，
    之後重新整理頁面、重啟程式都不會把它復原——這是刻意設計成這樣，
    否則每次 init_db() 都呼叫的話，已經刪除的預設項目會被「INSERT OR IGNORE」重新插回去。
    """
    with get_connection(db_path) as conn:
        conn.executescript(SCHEMA)
        conn.execute(
            "INSERT OR IGNORE INTO settings (id) VALUES (1)"
        )

        region_count = conn.execute("SELECT COUNT(*) AS c FROM shipping_regions").fetchone()["c"]
        if region_count == 0:
            for region in DEFAULT_REGIONS:
                conn.execute(
                    "INSERT OR IGNORE INTO shipping_regions (region_name, unit_rate) VALUES (?, 0)",
                    (region,),
                )

        carton_count = conn.execute("SELECT COUNT(*) AS c FROM carton_sizes").fetchone()["c"]
        if carton_count == 0:
            for size, rate in DEFAULT_CARTON_SIZES.items():
                conn.execute(
                    "INSERT OR IGNORE INTO carton_sizes (size_name, unit_rate) VALUES (?, ?)",
                    (size, rate),
                )

        _migrate_missing_columns(conn)


def _migrate_missing_columns(conn) -> None:
    """檢查既有資料表是否缺少新版欄位，缺少就用 ALTER TABLE 補上（保留舊資料）。"""
    expected_columns = {
        "projects": [
            ("shipping_order_quantity", "REAL"),
            ("tier_price_low", "REAL"),
            ("tier_price_mid", "REAL"),
            ("tier_price_high", "REAL"),
        ],
        "settings": [
            ("tier_margin_low", "REAL NOT NULL DEFAULT 0.25"),
            ("tier_margin_high", "REAL NOT NULL DEFAULT 0.35"),
            ("raw_material_loss_rate", "REAL NOT NULL DEFAULT 0"),
            ("labor_loss_rate", "REAL NOT NULL DEFAULT 0"),
            ("shipping_loss_rate", "REAL NOT NULL DEFAULT 0"),
        ],
    }
    for table, columns in expected_columns.items():
        existing = {row["name"] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}
        for col_name, col_type in columns:
            if col_name not in existing:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {col_name} {col_type}")

    # ---- 舊版 project_packaging 用 material_name 當外鍵，新版改用 material_id ----
    # 補上 material_id 欄位，並用舊的 material_name 對照「包材主檔」回填
    # （回填依據：早期 material_name 是唯一的，所以這個對照在當時是可靠的）
    pkg_cols = {row["name"] for row in conn.execute("PRAGMA table_info(project_packaging)").fetchall()}
    if "material_id" not in pkg_cols:
        conn.execute("ALTER TABLE project_packaging ADD COLUMN material_id TEXT")
        if "material_name" in pkg_cols:
            conn.execute(
                """UPDATE project_packaging
                   SET material_id = (
                       SELECT pm.material_id FROM packaging_materials pm
                       WHERE pm.material_name = project_packaging.material_name
                       LIMIT 1
                   )
                   WHERE material_id IS NULL"""
            )
        pkg_cols.add("material_id")

    # 重要修正：早期的 project_packaging 表，material_name 欄位是 NOT NULL（不能留空）。
    # 上面那段只是「多加一個 material_id 欄位」，並沒有真正解除 material_name 的 NOT NULL限制，
    # 導致新版程式存檔時(已經不會再填material_name了)，照樣會撞到這條舊規則，
    # 丟出 sqlite3.IntegrityError: NOT NULL constraint failed: project_packaging.material_name。
    # SQLite不支援直接「移除某欄位的NOT NULL」，唯一作法是整張表重建：
    # 建一張新表(只有目前真正需要的欄位) → 把舊資料搬過去 → 刪除舊表 → 把新表改回原名。
    if "material_name" in pkg_cols:
        conn.executescript("""
            ALTER TABLE project_packaging RENAME TO project_packaging_old;

            CREATE TABLE project_packaging (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
                material_id TEXT NOT NULL REFERENCES packaging_materials(material_id),
                quantity REAL NOT NULL DEFAULT 1,
                loss_rate REAL NOT NULL DEFAULT 0
            );

            INSERT INTO project_packaging (id, project_id, material_id, quantity, loss_rate)
            SELECT id, project_id, material_id, quantity, loss_rate
            FROM project_packaging_old
            WHERE material_id IS NOT NULL;

            DROP TABLE project_packaging_old;
        """)


# ============================================================
# 設定（全廠共用，唯一一筆）
# ============================================================

def load_settings(db_path: str = "cost_calculator.db") -> Settings:
    with get_connection(db_path) as conn:
        row = conn.execute("SELECT * FROM settings WHERE id = 1").fetchone()
        regions = conn.execute("SELECT region_name, unit_rate FROM shipping_regions").fetchall()
        cartons = conn.execute("SELECT size_name, unit_rate FROM carton_sizes").fetchall()
    if row is None:
        raise RuntimeError("settings 資料表是空的，請先呼叫 init_db()")
    return Settings(
        margin_mode=row["margin_mode"],
        margin_rate=row["margin_rate"],
        overhead_rate=row["overhead_rate"],
        tax_rate=row["tax_rate"],
        tax_mode=row["tax_mode"],
        monthly_salary=row["monthly_salary"],
        work_days_per_month=row["work_days_per_month"],
        work_hours_per_day=row["work_hours_per_day"],
        logistics_rate=row["logistics_rate"],
        pallet_rate=row["pallet_rate"],
        tier_margin_low=row["tier_margin_low"],
        tier_margin_high=row["tier_margin_high"],
        raw_material_loss_rate=row["raw_material_loss_rate"],
        labor_loss_rate=row["labor_loss_rate"],
        shipping_loss_rate=row["shipping_loss_rate"],
        region_rates={r["region_name"]: r["unit_rate"] for r in regions},
        carton_rates={c["size_name"]: c["unit_rate"] for c in cartons},
    )


def save_settings(settings: Settings, db_path: str = "cost_calculator.db") -> None:
    with get_connection(db_path) as conn:
        conn.execute(
            """UPDATE settings SET
                margin_mode=?, margin_rate=?, overhead_rate=?, tax_rate=?, tax_mode=?,
                monthly_salary=?, work_days_per_month=?, work_hours_per_day=?,
                logistics_rate=?, pallet_rate=?, tier_margin_low=?, tier_margin_high=?,
                raw_material_loss_rate=?, labor_loss_rate=?, shipping_loss_rate=?
               WHERE id = 1""",
            (settings.margin_mode, settings.margin_rate, settings.overhead_rate,
             settings.tax_rate, settings.tax_mode, settings.monthly_salary,
             settings.work_days_per_month, settings.work_hours_per_day,
             settings.logistics_rate, settings.pallet_rate,
             settings.tier_margin_low, settings.tier_margin_high,
             settings.raw_material_loss_rate, settings.labor_loss_rate, settings.shipping_loss_rate),
        )
        # 重要修正：原本這裡只有 INSERT...ON CONFLICT（新增/更新），從來沒有刪除過任何資料，
        # 所以使用者在「共用設定」頁的表格裡刪除一列、按下儲存，資料庫裡那一列根本沒被刪掉，
        # 只是沒被覆蓋而已，畫面上才會看起來「怎麼刪都刪不掉」。
        # 修正方式：採用「先清空、再整批寫入」，讓資料庫的內容跟畫面上的表格完全同步。
        conn.execute("DELETE FROM shipping_regions")
        for region, rate in settings.region_rates.items():
            conn.execute(
                "INSERT INTO shipping_regions (region_name, unit_rate) VALUES (?, ?)",
                (region, rate),
            )
        conn.execute("DELETE FROM carton_sizes")
        for size, rate in settings.carton_rates.items():
            conn.execute(
                "INSERT INTO carton_sizes (size_name, unit_rate) VALUES (?, ?)",
                (size, rate),
            )


# ============================================================
# 包材主檔
# ============================================================

def upsert_packaging_material(material: PackagingMaterial, db_path: str = "cost_calculator.db") -> None:
    """
    新增或更新一筆包材主檔資料。

    重要：「包材編號」(material_id) 是唯一識別碼，「包材名稱」(material_name) 可以重複
    （例如不同批次、不同供應商但用同一個品名）。所以比對／更新一律以編號為準。
    """
    with get_connection(db_path) as conn:
        conn.execute(
            """INSERT INTO packaging_materials
                (material_id, material_name, spec, unit, unit_price, vendor, note)
               VALUES (?, ?, ?, ?, ?, ?, ?)
               ON CONFLICT(material_id) DO UPDATE SET
                material_name=excluded.material_name, spec=excluded.spec, unit=excluded.unit,
                unit_price=excluded.unit_price, vendor=excluded.vendor, note=excluded.note,
                updated_at=datetime('now')""",
            (material.material_id, material.name, material.spec, material.unit,
             material.unit_price, material.vendor, material.note),
        )


def load_packaging_catalog(db_path: str = "cost_calculator.db") -> PackagingCatalog:
    with get_connection(db_path) as conn:
        rows = conn.execute("SELECT * FROM packaging_materials").fetchall()
    materials = [
        PackagingMaterial(
            material_id=r["material_id"], name=r["material_name"], spec=r["spec"] or "",
            unit=r["unit"] or "", unit_price=r["unit_price"], vendor=r["vendor"] or "",
            note=r["note"] or "",
        )
        for r in rows
    ]
    return PackagingCatalog(materials)


# ============================================================
# 專案（成本試算案）：存入 / 讀出
# ============================================================

def save_project(project: CostProject, db_path: str = "cost_calculator.db") -> None:
    """寫入或覆蓋一筆專案及其所有明細列（先刪除舊明細再整批寫入，簡化覆寫邏輯）"""
    with get_connection(db_path) as conn:
        conn.execute(
            """INSERT INTO projects
                (project_id, customer_name, product_name, monthly_quantity,
                 raw_material_unit_price)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(project_id) DO UPDATE SET
                customer_name=excluded.customer_name, product_name=excluded.product_name,
                monthly_quantity=excluded.monthly_quantity,
                raw_material_unit_price=excluded.raw_material_unit_price""",
            (project.project_id, project.customer_name, project.product_name,
             project.monthly_quantity, project.raw_material.unit_price),
        )

        for table in ("project_packaging", "project_labor", "project_shipping", "project_carton"):
            conn.execute(f"DELETE FROM {table} WHERE project_id = ?", (project.project_id,))

        for line in project.packaging_lines:
            conn.execute(
                "INSERT INTO project_packaging (project_id, material_id, quantity, loss_rate) "
                "VALUES (?, ?, ?, ?)",
                (project.project_id, line.material_id, line.quantity, line.loss_rate),
            )
        for line in project.labor_lines:
            conn.execute(
                "INSERT INTO project_labor "
                "(project_id, process_name, batch_quantity, batch_time_minutes, headcount, hourly_wage_override) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (project.project_id, line.process_name, line.batch_quantity,
                 line.batch_time_minutes, line.headcount, line.hourly_wage_override),
            )
        for line in project.shipping_lines:
            conn.execute(
                "INSERT INTO project_shipping (project_id, method, quantity, region) VALUES (?, ?, ?, ?)",
                (project.project_id, line.method, line.quantity, line.region),
            )
        for line in project.carton_lines:
            conn.execute(
                "INSERT INTO project_carton (project_id, size, quantity) VALUES (?, ?, ?)",
                (project.project_id, line.size, line.quantity),
            )


def load_project(project_id: str, db_path: str = "cost_calculator.db") -> CostProject:
    with get_connection(db_path) as conn:
        p = conn.execute("SELECT * FROM projects WHERE project_id = ?", (project_id,)).fetchone()
        if p is None:
            raise KeyError(f"找不到專案 {project_id}")
        packaging = conn.execute(
            "SELECT * FROM project_packaging WHERE project_id = ?", (project_id,)
        ).fetchall()
        labor = conn.execute(
            "SELECT * FROM project_labor WHERE project_id = ?", (project_id,)
        ).fetchall()
        shipping = conn.execute(
            "SELECT * FROM project_shipping WHERE project_id = ?", (project_id,)
        ).fetchall()
        carton = conn.execute(
            "SELECT * FROM project_carton WHERE project_id = ?", (project_id,)
        ).fetchall()

    return CostProject(
        project_id=p["project_id"],
        customer_name=p["customer_name"],
        product_name=p["product_name"],
        monthly_quantity=p["monthly_quantity"],
        raw_material=RawMaterial(
            unit_price=p["raw_material_unit_price"],
        ),
        packaging_lines=[
            PackagingLine(r["material_id"], r["quantity"], r["loss_rate"]) for r in packaging
        ],
        labor_lines=[
            LaborLine(r["process_name"], r["batch_quantity"], r["batch_time_minutes"],
                      r["headcount"], r["hourly_wage_override"])
            for r in labor
        ],
        shipping_lines=[
            ShippingLine(r["method"], r["quantity"], r["region"]) for r in shipping
        ],
        carton_lines=[
            CartonLine(r["size"], r["quantity"]) for r in carton
        ],
    )


def list_projects(db_path: str = "cost_calculator.db") -> list[dict]:
    with get_connection(db_path) as conn:
        rows = conn.execute(
            "SELECT project_id, customer_name, product_name, monthly_quantity, created_at "
            "FROM projects ORDER BY created_at DESC"
        ).fetchall()
    return [dict(r) for r in rows]


def delete_project(project_id: str, db_path: str = "cost_calculator.db") -> None:
    with get_connection(db_path) as conn:
        conn.execute("DELETE FROM projects WHERE project_id = ?", (project_id,))
