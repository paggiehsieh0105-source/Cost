"""
驗證 cost_model.py 的計算結果（2026-06-25 重大調整後的版本）。

本次調整重點：
1. 原料(RawMaterial) 移除「件數」欄位，只剩「整組配方成本」，直接當單件原料成本使用
2. 移除「外包加工費」整個模組
3. 物流／紙箱費用一律用「訂單量」(monthly_quantity) 當分攤分母，不再有其他分母選項
4. 新增「原物料／人工／物流」三個損耗率(%)，在算完各自小計後乘以(1+損耗率)當整體緩衝
"""
import math
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from cost_model import (
    Settings, PackagingCatalog, CostProject, RawMaterial,
    PackagingLine, LaborLine, ShippingLine, CartonLine,
    calculate_total_cost,
)


def build_example_project(monthly_quantity=1000) -> CostProject:
    return CostProject(
        project_id="OEM-2026-001",
        customer_name="範例客戶",
        product_name="範例產品",
        monthly_quantity=monthly_quantity,
        raw_material=RawMaterial(unit_price=10.0),  # 不再有件數，直接是單件原料成本
        packaging_lines=[
            PackagingLine("PKG-001", quantity=1, loss_rate=0.01),  # 瓶身 3.5*1.01=3.535
            PackagingLine("PKG-003", quantity=1, loss_rate=0.01),  # 噴頭 2.0*1.01=2.02
        ],
        labor_lines=[
            LaborLine("配方攪拌", batch_quantity=5000, batch_time_minutes=100, headcount=1),
            LaborLine("充填", batch_quantity=5000, batch_time_minutes=125, headcount=2),
        ],
        shipping_lines=[ShippingLine(method="物流", quantity=monthly_quantity)],
    )


def assert_close(actual, expected, label, tol=0.01):
    assert math.isclose(actual, expected, abs_tol=tol), (
        f"[FAIL] {label}: got {actual:.6f}, expected {expected:.6f}"
    )
    print(f"[PASS] {label}: {actual:.4f} (expected: {expected:.4f})")


def main():
    settings = Settings()  # 預設三個損耗率都是0
    catalog = PackagingCatalog.sample()
    project = build_example_project(monthly_quantity=1000)

    result = calculate_total_cost(project, settings, catalog)

    # ---- 原物料成本：單件原料成本 + 包材，無損耗率時直接加總 ----
    expected_raw = 10.0 + 3.535 + 2.02  # = 15.555
    assert_close(result.raw_material_total, expected_raw, "原物料成本(無損耗率時)")

    # ---- 人工成本：跟原本邏輯一致，不受本次調整影響 ----
    # 配方攪拌：100/5000分鐘*時薪/60*1人；充填：125/5000*時薪/60*2人
    hourly_wage = settings.hourly_wage
    expected_labor = (100/5000)*(hourly_wage/60)*1 + (125/5000)*(hourly_wage/60)*2
    assert_close(result.labor_total, expected_labor, "人工成本(無損耗率時)")

    expected_overhead = (expected_raw + expected_labor) * settings.overhead_rate
    assert_close(result.overhead, expected_overhead, "製造費用")

    # ---- 物流：固定用「訂單量」當分母 ----
    # 物流線設定 quantity=monthly_quantity(1000)，每件150元 → 總額150000，分攤回每件還是150元
    expected_shipping_total = 1000 * settings.logistics_rate
    assert_close(result.shipping_total, expected_shipping_total, "物流運費總額")
    assert_close(result.shipping_cost_per_unit, settings.logistics_rate, "物流運費分攤後(用訂單量當分母)")

    expected_unit_cost = expected_raw + expected_labor + expected_overhead + settings.logistics_rate
    assert_close(result.unit_cost, expected_unit_cost, "單件總成本(已移除外包加工費)")

    print("\n所有基本測試通過 ✅")

    # ---- 額外測試：三個損耗率(%)套用在各自的小計上 ----
    print("\n--- 額外測試：原物料／人工／物流 損耗率 ---")
    settings_with_loss = Settings(
        raw_material_loss_rate=0.05,  # 原物料多抓5%
        labor_loss_rate=0.02,         # 人工多抓2%
        shipping_loss_rate=0.03,      # 物流多抓3%
    )
    result_loss = calculate_total_cost(project, settings_with_loss, catalog)

    assert_close(result_loss.raw_material_total, expected_raw * 1.05, "原物料成本(含5%損耗率)")
    assert_close(result_loss.labor_total, expected_labor * 1.02, "人工成本(含2%損耗率)")
    assert_close(result_loss.shipping_cost_per_unit, settings.logistics_rate * 1.03, "物流運費分攤後(含3%損耗率)")
    print("[PASS] 三個損耗率都正確套用在各自的成本小計上")

    # ---- 額外測試：紙箱費用也是用「訂單量」當分攤分母 ----
    print("\n--- 額外測試：紙箱用量 = 箱數量 ÷ 訂單量 ---")
    project2 = build_example_project(monthly_quantity=1000)
    project2.shipping_lines = []  # 這次只測紙箱，不疊加物流線
    project2.carton_lines = [
        CartonLine(size="小箱(20x20x20cm)", quantity=50),   # 50箱 × 15元 = 750元
    ]
    result2 = calculate_total_cost(project2, settings, catalog)
    expected_carton_total = 50 * 15  # 750
    expected_carton_per_unit = expected_carton_total / 1000  # 750 / 1000訂單量 = 0.75元/件
    assert_close(result2.carton_total, expected_carton_total, "紙箱費用總額(50箱×15元)")
    assert_close(result2.shipping_cost_per_unit, expected_carton_per_unit, "紙箱分攤後(箱數×單價 ÷ 訂單量)")
    print(f"[PASS] 紙箱費用正確用「訂單量」當分母：{expected_carton_total}元 ÷ 1000件 = {expected_carton_per_unit:.3f}元/件")

    # ---- 額外測試：切換售價基礎(毛利率法)，級距報價公式要一致 ----
    print("\n--- 額外測試：級距報價(直接輸入毛利率) ---")
    assert_close(result.tier_margin_low, 0.25, "低量毛利率(直接輸入)")
    assert_close(result.tier_margin_high, 0.35, "高量毛利率(直接輸入)")
    assert_close(result.tier_price_low, expected_unit_cost * 1.25, "低量級距報價(成本×1.25)")
    assert_close(result.tier_price_high, expected_unit_cost * 1.35, "高量級距報價(成本×1.35)")
    print("[PASS] 級距報價正確套用直接輸入的毛利率")

    print("\n所有測試通過：Python計算模型運作正常 ✅")


if __name__ == "__main__":
    main()
