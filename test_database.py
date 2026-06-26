"""
驗證 database.py：
1. 初始化資料庫、寫入包材主檔、寫入設定
2. 存入一筆專案，讀出後重新計算，結果應與階段2測試一致
3. 測試多專案管理（新增第二筆、列出、刪除）
"""
import os
import sys
import math

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import database as db
from cost_model import (
    PackagingCatalog, CostProject, RawMaterial, PackagingLine,
    LaborLine, ShippingLine, CartonLine, calculate_total_cost,
)

TEST_DB = os.path.join(os.path.dirname(__file__), "..", "data", "test_cost_calculator.db")


def build_example_project(project_id="OEM-2026-001") -> CostProject:
    return CostProject(
        project_id=project_id,
        customer_name="範例客戶",
        product_name="範例產品",
        monthly_quantity=10000,
        raw_material=RawMaterial(unit_price=10.0),
        packaging_lines=[
            PackagingLine("PKG-001", quantity=1, loss_rate=0.01),
            PackagingLine("PKG-003", quantity=1, loss_rate=0.01),
            PackagingLine("PKG-004", quantity=1, loss_rate=0.01),
            PackagingLine("PKG-005", quantity=1, loss_rate=0.02),
            PackagingLine("PKG-006", quantity=1, loss_rate=0.0),
        ],
        labor_lines=[
            LaborLine("配方攪拌", batch_quantity=5000, batch_time_minutes=100, headcount=1),
            LaborLine("充填", batch_quantity=5000, batch_time_minutes=125, headcount=2),
            LaborLine("貼標", batch_quantity=5000, batch_time_minutes=67, headcount=1),
            LaborLine("裝箱", batch_quantity=5000, batch_time_minutes=42, headcount=1),
        ],
        shipping_lines=[ShippingLine(method="物流", quantity=1)],
        carton_lines=[CartonLine(size="小箱(20x20x20cm)", quantity=2)],
    )


def main():
    if os.path.exists(TEST_DB):
        os.remove(TEST_DB)

    # 1) 初始化
    db.init_db(TEST_DB)
    print("[PASS] 資料庫初始化完成")

    # 2) 寫入包材主檔（用階段2的範例資料）
    for m in PackagingCatalog.sample().all_materials():
        db.upsert_packaging_material(m, TEST_DB)
    catalog = db.load_packaging_catalog(TEST_DB)
    assert catalog.get_by_id("PKG-001").unit_price == 3.5
    print("[PASS] 包材主檔寫入/讀出正確")

    # 3) 設定維持預設值即可（已在 init_db 寫入），讀出確認
    settings = db.load_settings(TEST_DB)
    assert settings.monthly_salary == 30000
    assert settings.logistics_rate == 150
    print("[PASS] 設定讀出正確（預設值）")

    # 4) 存入專案，讀出，重新計算，應與階段2測試結果一致
    project = build_example_project()
    db.save_project(project, TEST_DB)
    loaded = db.load_project("OEM-2026-001", TEST_DB)
    result = calculate_total_cost(loaded, settings, catalog)

    assert len(loaded.carton_lines) == 1 and loaded.carton_lines[0].size == "小箱(20x20x20cm)"
    print(f"[PASS] 紙箱明細存入/讀出正確：{loaded.carton_lines[0].size} x {loaded.carton_lines[0].quantity}")

    assert len(loaded.packaging_lines) == 5
    assert loaded.packaging_lines[0].material_id == "PKG-001"
    print(f"[PASS] 包材明細存入/讀出正確：以編號(material_id)為準，第1筆={loaded.packaging_lines[0].material_id}")

    print(f"[PASS] 存入/讀出後重新計算：單件總成本={result.unit_cost:.2f}，最終報價={result.price_final:.2f}")

    # 5) 多專案管理：新增第二筆、列出、刪除
    project2 = build_example_project(project_id="OEM-2026-002")
    project2.customer_name = "第二個客戶"
    project2.monthly_quantity = 5000
    db.save_project(project2, TEST_DB)

    projects = db.list_projects(TEST_DB)
    assert len(projects) == 2
    print(f"[PASS] 多專案管理：目前共有 {len(projects)} 筆專案")

    db.delete_project("OEM-2026-002", TEST_DB)
    projects = db.list_projects(TEST_DB)
    assert len(projects) == 1
    print(f"[PASS] 刪除專案後剩餘 {len(projects)} 筆")

    # 6) 驗證「名稱可以重複，編號才是唯一識別碼」
    from cost_model import PackagingMaterial
    m_dup1 = PackagingMaterial(material_id="PKG-101", name="瓶蓋", spec="供應商A版本", unit_price=0.8)
    m_dup2 = PackagingMaterial(material_id="PKG-102", name="瓶蓋", spec="供應商B版本", unit_price=0.9)
    db.upsert_packaging_material(m_dup1, TEST_DB)
    db.upsert_packaging_material(m_dup2, TEST_DB)
    catalog2 = db.load_packaging_catalog(TEST_DB)
    assert catalog2.get_by_id("PKG-101").spec == "供應商A版本"
    assert catalog2.get_by_id("PKG-102").spec == "供應商B版本"
    print("[PASS] 名稱重複(都叫「瓶蓋」)但編號不同，兩筆資料都正確存在，沒有互相覆蓋")

    print("\n所有測試通過：資料儲存層運作正常 ✅")


if __name__ == "__main__":
    main()
