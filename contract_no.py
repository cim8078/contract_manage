# -*- coding: utf-8 -*-
"""合同编号规则（6 位）——统一编号的解析、校验与顺序号生成。

编号格式::

    第 1-2 位 年度（年后两位，如 2026 -> "26"）
    第 3   位 科室（中心）序号
    第 4   位 中心组室序号
    第 5-6 位 顺序号（不足两位左补 0；同一“年度+科室+组室”前缀下从 01 连续递增）

    示例：266401 = 2026 年 + 6（后勤中心）+ 4（物业组）+ 01（第 1 号）

本模块只负责“规则 -> 前缀”的解析；顺序号需由调用方查询对应表后经
``next_number()`` 计算（同一表内唯一）。未列明的科室（中心）/中心组室不生成、
不猜测编号——``resolve()`` 会抛出 :class:`NoRuleError`。

如需扩展：在 ``CENTER_RULES`` / ``BUREAU_SECTION_CODES`` 中补充中心组室编号，
在 ``UNIT_ALIASES`` 中登记单位名称前缀即可（科室/组室一律使用统一后名称）。

名称口径（2026-10 统一后，与 departments 表核对一致）::

    第 3 位  1=办公室 2=国有资产科 3=公共事务科 4=人事科 5=财务科
             6=后勤中心 7=资产中心 8=公车中心 9=政策法规科
    第 4 位  局本级科室(1/2/3/4/5/9)统一为 1；后勤中心 1~7、资产中心 1~4、公车中心 1

历史写法（如“物业管理组/工程建设组/人事与政策法规科/国有资产管理科”等）不再做自动归一，
一律按“未配置编号规则”处理，避免混用旧名称生成编号。
"""

import datetime

# ─── 常量 ───
MAX_SEQ = 99          # 当前 6 位规则下顺序号上限（01~99），超出需扩展编号规则
SEQ_DIGITS = 2
YEAR_DIGITS = 2

# ─── 一、第 3 位：科室（中心）序号 ───
# 局本级科室（第 3 位 1/2/3/4/5/9；其第 4 位固定为 1）
# 说明：原“人事与政策法规科”已拆分为“人事科”(4) 与“政策法规科”(9)。
BUREAU_SECTION_CODES = {
    "办公室": 1,
    "国有资产科": 2,
    "公共事务科": 3,
    "人事科": 4,
    "财务科": 5,
    "政策法规科": 9,
}
BUREAU_UNIT = "局本级"                       # 局本级的规范名称
BUREAU_GROUP_DIGIT = 1                       # 局本级科室的第 4 位统一为 1

# 中心（第 3 位 6/7/8）及其中心组室（第 4 位）——名称均为统一后名称
CENTER_RULES = {
    "后勤中心": (6, {
        "综合组": 1, "公车组": 2, "工程组": 3, "物业组": 4, "后勤组": 5,
        "市民之家专班": 6, "周转住房专班": 7,
    }),
    "资产中心": (7, {
        "综合组": 1, "房地资产组": 2, "通用资产组": 3, "集采组": 4,
    }),
    "公车中心": (8, {
        "现场负责组": 1,
    }),
}

# 未配置编号规则、但历史上出现过的名称（仅用于给出更明确的提示，绝不据此生成编号）
LEGACY_SECTION_NAMES = {
    "人事与政策法规科": "该科室已拆分为“人事科”(第3位 4) 与“政策法规科”(第3位 9)，部门表旧行已删除",
    "国有资产管理科": "该科室已统一为“国有资产科”(第3位 2)",
    "公共事务管理科": "该科室已统一为“公共事务科”(第3位 3)",
    "财务管理科": "该科室已统一为“财务科”(第3位 5)",
    "资产科": "该科室已统一为“国有资产科”(第3位 2)",
    "公车管理组": "该组室已统一为“公车组”(第4位 2)",
    "工程建设组": "该组室已统一为“工程组”(第4位 3)",
    "物业管理组": "该组室已统一为“物业组”(第4位 4)",
    "后勤管理组": "该组室已统一为“后勤组”(第4位 5)",
    "国资组": "资产中心未列明的中心组室",
    "办公用房组": "资产中心未列明的中心组室",
    "主任班子": "经业务确认不纳入自动编号，请人工填写合同编号",
    "局领导": "经业务确认不纳入自动编号，请人工填写合同编号",
}

# ─── 二、单位名称归一（部门表登记的单位名称 -> 规则中的中心名称）───
# 说明：部门表登记的单位名称为“湘潭市机关事务管理局XX”，规则中的“中心”名称使用简称；
#       此处仅做单位前缀归一。科室/组室名称一律使用统一后名称，不做旧名归一。
UNIT_ALIASES = {
    "湘潭市机关事务管理局本级": BUREAU_UNIT,
    "湘潭市机关事务管理局": BUREAU_UNIT,
    "湘潭市机关事务管理局后勤中心": "后勤中心",
    "湘潭市机关事务管理局资产中心": "资产中心",
    "湘潭市机关事务管理局公车中心": "公车中心",
    BUREAU_UNIT: BUREAU_UNIT,
    "后勤中心": "后勤中心",
    "资产中心": "资产中心",
    "公车中心": "公车中心",
}

# 科室/组室名称一律以统一后名称为准，不做旧名归一（旧名见 LEGACY_SECTION_NAMES，按“未配置规则”处理）
SECTION_ALIASES = {}


class NoRuleError(ValueError):
    """未配置 6 位编号规则（未列明的中心/组室不得自动生成或猜测编号）。"""


def normalize_unit(unit):
    """单位名称归一化；未知单位原样返回。"""
    u = str(unit or "").strip()
    return UNIT_ALIASES.get(u, u)


def normalize_section(section):
    """科室/组室名称归一化；未知名称原样返回。"""
    s = str(section or "").strip()
    return SECTION_ALIASES.get(s, s)


def split_department(dept):
    """拆分“所属部门”字符串（形如 单位名称-科室名称）为 (单位, 科室)。

    不含分隔符时，整体视为科室名称（兼容只登记组室名称的历史数据）。
    """
    s = str(dept or "").strip()
    if not s:
        return "", ""
    for sep in ("-", "－", "—", "/"):
        if sep in s:
            unit, sec = s.split(sep, 1)
            return unit.strip(), sec.strip()
    return "", s


def resolve(unit, section):
    """解析 (单位, 科室/组室) -> (第 3 位, 第 4 位)。

    无法确定编号时抛出 :class:`NoRuleError`（不猜测编号）。
    """
    u = normalize_unit(unit)
    s = normalize_section(section)
    # 提示后缀：历史/未列明名称给出明确说明，其余统一提示“先补充规则”
    suffix = ("（" + LEGACY_SECTION_NAMES[s] + "）") if s in LEGACY_SECTION_NAMES else "，请先补充编号规则后再生成编号"

    if u in CENTER_RULES:
        d3, groups = CENTER_RULES[u]
        if s in groups:
            return str(d3), str(groups[s])
        raise NoRuleError("中心「%s」的中心组室「%s」未配置编号规则%s" % (u, section or "（空）", suffix))

    if u == BUREAU_UNIT:
        if s in BUREAU_SECTION_CODES:
            return str(BUREAU_SECTION_CODES[s]), str(BUREAU_GROUP_DIGIT)
        raise NoRuleError("本科室「%s」未配置编号规则%s" % (section or "（空）", suffix))

    # 兼容：未登记单位、但科室名称本身属于局本级科室
    if not u and s in BUREAU_SECTION_CODES:
        return str(BUREAU_SECTION_CODES[s]), str(BUREAU_GROUP_DIGIT)

    raise NoRuleError("科室（中心）「%s」未配置编号规则%s" % (u or section or "（空）", suffix))


def year_digits(year=None):
    """年度后两位字符串，如 2026 -> "26"。"""
    y = int(year or datetime.datetime.now().year)
    return "%0*d" % (YEAR_DIGITS, y % 100)


def prefix_for(unit, section, year=None):
    """返回 4 位前缀（年度后两位 + 科室序号 + 中心组室序号）。"""
    d3, d4 = resolve(unit, section)
    return year_digits(year) + d3 + d4


def build_number(prefix4, seq):
    """按前缀与顺序号拼装 6 位编号字符串。"""
    return "%s%0*d" % (prefix4, SEQ_DIGITS, int(seq))


def next_number(existing_numbers, prefix4, max_seq=MAX_SEQ):
    """在既有编号集合中查找同前缀最大顺序号并加 1。

    仅统计长度 6、前 4 位等于 ``prefix4``、后 2 位为数字的编号；
    既有编号集合为空（该前缀下暂无记录）时从 ``01`` 开始；
    若生成的编号已存在（异常数据）则继续向后取，直至可用；
    顺序号超过 ``max_seq`` 时抛出 :class:`OverflowError`（需扩展编号规则）。
    """
    used = set()
    mx = 0
    for n in existing_numbers or []:
        s = str(n or "").strip()
        if not s:
            continue
        used.add(s)
        if len(s) == 4 + SEQ_DIGITS and s[:4] == prefix4 and s[4:].isdigit():
            mx = max(mx, int(s[4:]))
    seq = mx + 1
    while seq <= max_seq:
        cand = build_number(prefix4, seq)
        if cand not in used:
            return cand
        seq += 1
    raise OverflowError("前缀 %s 的顺序号已达上限 %d，需扩展编号规则" % (prefix4, max_seq))


def plan(dept=None, year=None, unit=None, section=None):
    """规则预览：返回 dict(prefix/unit/section/d3/d4/year)；无规则时抛 NoRuleError。"""
    if unit is None and section is None:
        unit, section = split_department(dept)
    u = normalize_unit(unit)
    s = normalize_section(section)
    d3, d4 = resolve(u, s)
    y = int(year or datetime.datetime.now().year)
    return {
        "year": y, "year_digits": year_digits(y), "unit": u, "section": s,
        "d3": d3, "d4": d4, "prefix": year_digits(y) + d3 + d4,
    }


def catalog():
    """规则目录（供文档/页面展示）：第 3 位科室（中心）及其第 4 位组室。"""
    rows = []
    for name, code in sorted(BUREAU_SECTION_CODES.items(), key=lambda kv: kv[1]):
        rows.append({"d3": str(code), "name": name, "groups": {"（第 4 位统一为 1）": 1}})
    for name, (code, groups) in sorted(CENTER_RULES.items(), key=lambda kv: kv[1][0]):
        rows.append({"d3": str(code), "name": name, "groups": dict(groups)})
    return rows
