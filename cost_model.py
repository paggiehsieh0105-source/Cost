"""
代工成本計算模型 (階段2：後端資料計算模型)

完整公式邏輯說明請參考專案根目錄 CLAUDE.md。
本模組的目標：與 Excel 雛形（代工成本試算表.xlsx）的計算結果逐項對應。
"""

from dataclasses import dataclass, field
from typing import Literal, Optional


# ============================================================
# 共用參數 / 主檔
# ============================================================

@dataclass
class Settings:
    """對應 Excel「設定」工作表"""
    margin_mode: Literal["成本基礎(加成法)", "售價基礎(毛利率法)"] = "成本基礎(加成法)"
    margin_rate: float = 0.30
    overhead_rate: float = 0.15
    tax_rate: float = 0.05
    tax_mode: Literal["未稅", "含稅"] = "未稅"

    # 人工時薪換算
    monthly_salary: float = 30000
    work_days_per_month: int = 22
    work_hours_per_day: float = 8

    # 物流／棧板固定費率
    logistics_rate: float = 150
    pallet_rate: float = 280

    # 專車依縣市分區單價
    region_rates: dict = field(default_factory=lambda: {
        "中彰投雲林": 0,
        "竹苗嘉": 0,
        "桃園台北台南高雄": 0,
        "宜蘭屏東": 0,
        "台東花蓮": 0,
    })

    # 紙箱依尺寸的費率（每個紙箱多少錢，尺寸名稱可自訂）
    carton_rates: dict = field(default_factory=lambda: {
        "小箱(20x20x20cm)": 15,
        "中箱(30x30x30cm)": 25,
        "大箱(40x40x40cm)": 35,
    })

    # 級距報價：低量／高量 的毛利率直接由使用者輸入（不是用「增減幅度」自動算），
    # 中量(主檔)則沿用上面的 margin_rate
    tier_margin_low: float = 0.25
    tier_margin_high: float = 0.35

    # 損耗率（％）：分別套用在原物料／人工／物流，算完各自的成本小計之後，
    # 再乘以 (1+損耗率) 當作一個整體的合理性緩衝，例如原物料損耗率5%代表
    # 「算出來的原物料成本，實際因為製程損耗，大概還要多抓5%才夠」
    raw_material_loss_rate: float = 0.0
    labor_loss_rate: float = 0.0
    shipping_loss_rate: float = 0.0

    @property
    def hourly_wage(self) -> float:
        """換算時薪 = 月薪 ÷ (工作天數 × 每日工時)"""
        return self.monthly_salary / (self.work_days_per_month * self.work_hours_per_day)


@dataclass
class PackagingMaterial:
    """包材主檔單筆資料，對應 Excel「包材資料表」工作表"""
    material_id: str
    name: str
    spec: str = ""
    unit: str = ""
    unit_price: float = 0.0
    vendor: str = ""
    note: str = ""


class PackagingCatalog:
    """
    包材主檔集合。

    重要設計：包材「編號」(material_id) 是唯一識別碼，「名稱」(name) 可以重複
    （例如不同供應商、不同批次但用同一個品名）。所以查價、查料一律以編號為準，
    名稱只用來搜尋／顯示，不能拿名稱去查到「唯一」一筆資料（因為名稱可能對應到多筆）。
    """

    def __init__(self, materials: Optional[list[PackagingMaterial]] = None):
        self._by_id: dict[str, PackagingMaterial] = {}
        self._names_by_id: dict[str, str] = {}  # 方便顯示：編號 -> 名稱
        for m in (materials or []):
            self.add(m)

    def add(self, material: PackagingMaterial) -> None:
        self._by_id[material.material_id] = material
        self._names_by_id[material.material_id] = material.name

    def get_by_id(self, material_id: str) -> "PackagingMaterial":
        if material_id not in self._by_id:
            raise KeyError(f"包材主檔找不到編號「{material_id}」，請先在包材資料表新增")
        return self._by_id[material_id]

    def id_to_name(self, material_id: str) -> Optional[str]:
        return self._names_by_id.get(material_id)

    def all_names(self) -> list[str]:
        """所有包材名稱（不重複），用於下拉選單。若名稱重複，這裡只會列一次。"""
        seen = []
        for m in self._by_id.values():
            if m.name not in seen:
                seen.append(m.name)
        return seen

    def find_by_name(self, name: str) -> Optional["PackagingMaterial"]:
        """依名稱查詢第一筆符合的資料。如果名稱重複對應到多筆，只會回傳第一筆找到的。"""
        for m in self._by_id.values():
            if m.name == name:
                return m
        return None

    def find_all_by_name(self, name: str) -> list["PackagingMaterial"]:
        """依名稱查詢所有符合的資料（名稱重複時會回傳多筆）。"""
        return [m for m in self._by_id.values() if m.name == name]

    def all_ids(self) -> list[str]:
        return list(self._by_id.keys())

    def all_materials(self) -> list["PackagingMaterial"]:
        return list(self._by_id.values())

    @classmethod
    def sample(cls) -> "PackagingCatalog":
        """CLAUDE.md 第四節的範例資料"""
        data = [
            ("PKG-001", "瓶身-30ml噴瓶", "PET透明 30ml", "個", 3.5),
            ("PKG-002", "瓶身-50ml噴瓶", "PET透明 50ml", "個", 4.2),
            ("PKG-003", "噴頭", "標準噴頭 18mm", "個", 2.0),
            ("PKG-004", "瓶蓋", "平蓋 18mm", "個", 0.8),
            ("PKG-005", "標籤", "貼紙標籤 防水", "張", 0.8),
            ("PKG-006", "外箱", "B浪外箱(分攤至單件)", "箱", 0.5),
            ("PKG-007", "說明書", "彩色說明書", "張", 0.3),
        ]
        return cls([PackagingMaterial(*row) for row in data])


# ============================================================
# 單一專案的成本明細項目
# ============================================================

@dataclass
class RawMaterial:
    """
    原料（簡化為單筆，整組配方視為一個總價）

    unit_price = 整組配方成本，已經是「做一件成品」需要用掉的整組原料總價。

    公式：原料成本小計 = 整組配方成本（不再乘件數，件數的概念留給最後
    「成本加總與報價試算結果」去乘訂單量算總營收即可，這裡只算「單件」原料成本）
    """
    unit_price: float = 0.0

    def subtotal(self) -> float:
        return self.unit_price


@dataclass
class PackagingLine:
    """
    包材明細單列，單價透過 material_id（包材主檔的唯一編號）查詢

    quantity = 用量，**是「一件成品」實際使用量**（例如一個瓶身、一個瓶蓋通常填1），
    不是訂單總數量！訂單總共要採購多少包材，由 monthly_quantity × quantity 換算，
    不需要使用者自己把訂單量乘進這裡。
    """
    material_id: str
    quantity: float = 1.0
    loss_rate: float = 0.0

    def subtotal(self, catalog: PackagingCatalog) -> float:
        material = catalog.get_by_id(self.material_id)
        return self.quantity * material.unit_price * (1 + self.loss_rate)


@dataclass
class LaborLine:
    """人工製程單列，每件加工時間由批量回推"""
    process_name: str
    batch_quantity: float
    batch_time_minutes: float
    headcount: int = 1
    hourly_wage_override: Optional[float] = None  # 個別製程可覆寫時薪

    def per_unit_minutes(self) -> float:
        if not self.batch_quantity:
            return 0.0
        return self.batch_time_minutes / self.batch_quantity

    def subtotal(self, settings: Settings) -> float:
        wage = self.hourly_wage_override if self.hourly_wage_override is not None else settings.hourly_wage
        return self.per_unit_minutes() * (wage / 60) * self.headcount


@dataclass
class ShippingLine:
    """物流運費單列：專車／物流／棧板"""
    method: Literal["專車", "物流", "棧板"]
    quantity: float
    region: Optional[str] = None  # 僅 method="專車" 時需要

    def unit_fee(self, settings: Settings) -> float:
        if self.method == "專車":
            if not self.region:
                raise ValueError("運送方式為「專車」時必須指定 region（縣市分區）")
            if self.region not in settings.region_rates:
                raise KeyError(f"找不到專車區域「{self.region}」的費率設定")
            return settings.region_rates[self.region]
        elif self.method == "物流":
            return settings.logistics_rate
        elif self.method == "棧板":
            return settings.pallet_rate
        raise ValueError(f"未知的運送方式：{self.method}")

    def subtotal(self, settings: Settings) -> float:
        return self.quantity * self.unit_fee(settings)


@dataclass
class CartonLine:
    """紙箱用量單列，依尺寸查詢「設定」工作表的紙箱費率"""
    size: str       # 紙箱尺寸名稱，須對應 settings.carton_rates 的鍵
    quantity: float  # 用了幾個這個尺寸的紙箱

    def unit_fee(self, settings: Settings) -> float:
        if self.size not in settings.carton_rates:
            raise KeyError(f"找不到紙箱尺寸「{self.size}」的費率設定")
        return settings.carton_rates[self.size]

    def subtotal(self, settings: Settings) -> float:
        return self.quantity * self.unit_fee(settings)


# ============================================================
# 專案 + 計算結果
# ============================================================

@dataclass
class CostProject:
    """單一代工成本試算案（對應 Excel「成本試算」工作表）"""
    project_id: str = ""
    customer_name: str = ""
    product_name: str = ""
    monthly_quantity: float = 0  # 訂單量；同時也是物流/紙箱費用分攤的分母
    raw_material: RawMaterial = field(default_factory=RawMaterial)
    packaging_lines: list[PackagingLine] = field(default_factory=list)
    labor_lines: list[LaborLine] = field(default_factory=list)
    shipping_lines: list[ShippingLine] = field(default_factory=list)
    carton_lines: list[CartonLine] = field(default_factory=list)


@dataclass
class CostBreakdown:
    """計算結果（各分項小計＋最終報價）"""
    raw_material_total: float
    labor_total: float
    overhead: float
    shipping_total: float          # 物流運費「總額」（含紙箱，整批/整單，未分攤前）
    carton_total: float            # 紙箱費用「總額」（已併入shipping_total，這裡單獨列出方便檢視）
    shipping_cost_per_unit: float  # 物流運費(含紙箱)「分攤到每件」的金額（= 已併入單件總成本）
    unit_cost: float
    price_pretax: float
    margin_amount: float
    price_final: float
    monthly_revenue: float
    # 級距報價（依「設定」工作表的級距毛利率增減幅度＋目前的毛利率計算模式自動算出）
    tier_margin_low: float
    tier_margin_mid: float
    tier_margin_high: float
    tier_price_low: float
    tier_price_mid: float
    tier_price_high: float


def calculate_total_cost(
    project: CostProject,
    settings: Settings,
    catalog: PackagingCatalog,
) -> CostBreakdown:
    """
    主計算函式，邏輯逐項對應 CLAUDE.md 第二節公式。
    """
    # 二、原物料成本（原料 + 包材），算完小計後套用「原物料損耗率」當整體緩衝
    raw_material_subtotal = project.raw_material.subtotal()
    packaging_total = sum(line.subtotal(catalog) for line in project.packaging_lines)
    raw_material_total = (raw_material_subtotal + packaging_total) * (1 + settings.raw_material_loss_rate)

    # 三、直接人工成本，算完小計後套用「人工損耗率」當整體緩衝
    labor_total = sum(line.subtotal(settings) for line in project.labor_lines)
    labor_total *= (1 + settings.labor_loss_rate)

    # 四、製造費用（比例攤提，用套用損耗率之後的原物料+人工去算）
    overhead = (raw_material_total + labor_total) * settings.overhead_rate

    # 五、物流運費（含紙箱費用）
    # 重要：每列的小計 = 數量 × 每件/每箱費率，這是「整批/整單」的費用總額，
    # 不能直接當成「單件成本」的一部分！必須先除以訂單量，換算成「每件分攤多少」。
    # 分母統一使用「訂單量」(project.monthly_quantity)，不再用各列數量加總或另一個獨立欄位，
    # 設計更單純：你填的「數量」只是用來算這一列自己的費用，分攤永遠用訂單量去除。
    shipping_lines_total = sum(line.subtotal(settings) for line in project.shipping_lines)
    carton_total = sum(line.subtotal(settings) for line in project.carton_lines)
    shipping_total = shipping_lines_total + carton_total

    if project.monthly_quantity:
        shipping_cost_per_unit = shipping_total / project.monthly_quantity
    else:
        shipping_cost_per_unit = 0.0
    shipping_cost_per_unit *= (1 + settings.shipping_loss_rate)

    # 六、成本加總與報價試算
    unit_cost = raw_material_total + labor_total + overhead + shipping_cost_per_unit

    if settings.margin_mode == "成本基礎(加成法)":
        price_pretax = unit_cost * (1 + settings.margin_rate)
    else:  # 售價基礎(毛利率法)
        price_pretax = unit_cost / (1 - settings.margin_rate)

    margin_amount = price_pretax - unit_cost

    if settings.tax_mode == "含稅":
        price_final = price_pretax * (1 + settings.tax_rate)
    else:
        price_final = price_pretax

    monthly_revenue = price_final * project.monthly_quantity

    # 級距報價：低量／高量 毛利率各增減 tier_margin_step，套用跟主檔一樣的 margin_mode 公式
    def _price_from_margin(margin_rate: float) -> float:
        if settings.margin_mode == "成本基礎(加成法)":
            pretax = unit_cost * (1 + margin_rate)
        else:  # 售價基礎(毛利率法)；毛利率不能 >= 1，超過就夾住在0.95避免除以負數或除以0
            safe_rate = min(margin_rate, 0.95)
            pretax = unit_cost / (1 - safe_rate)
        if settings.tax_mode == "含稅":
            return pretax * (1 + settings.tax_rate)
        return pretax

    tier_margin_mid = settings.margin_rate
    tier_margin_low = settings.tier_margin_low
    tier_margin_high = settings.tier_margin_high

    tier_price_mid = price_final
    tier_price_low = _price_from_margin(tier_margin_low)
    tier_price_high = _price_from_margin(tier_margin_high)

    return CostBreakdown(
        raw_material_total=raw_material_total,
        labor_total=labor_total,
        overhead=overhead,
        shipping_total=shipping_total,
        carton_total=carton_total,
        shipping_cost_per_unit=shipping_cost_per_unit,
        unit_cost=unit_cost,
        price_pretax=price_pretax,
        margin_amount=margin_amount,
        price_final=price_final,
        monthly_revenue=monthly_revenue,
        tier_margin_low=tier_margin_low,
        tier_margin_mid=tier_margin_mid,
        tier_margin_high=tier_margin_high,
        tier_price_low=tier_price_low,
        tier_price_mid=tier_price_mid,
        tier_price_high=tier_price_high,
    )
