# -*- coding: utf-8 -*-
"""按《系统开发文档》《数据字典表》口径**独立实现**的期望值计算器（oracle）。
用途：首页卡片 / 提醒 / 待收待支 / 统计的数值比对——避免直接把被测代码逻辑抄进断言。
所有函数只接收数据库行（dict），不做 IO。
"""
import datetime
import re


def num(v):
    """与 app._num 口径一致：仅接受纯数值字符串，round 2 位；其余返回 None。
    注意：不解析「约三万」这类文本、不去千分位逗号（项目现状即如此）。"""
    if v == "" or v is None:
        return None
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return None


def rent_totals(rent_rows):
    """{合同ID(int): {"count","due","paid"}}——与 app.asset_rent_totals() 同口径。"""
    out = {}
    for r in rent_rows:
        try:
            k = int(float(r.get("资产合同ID")))
        except (TypeError, ValueError):
            continue
        g = out.setdefault(k, {"count": 0, "due": 0.0, "paid": 0.0})
        g["count"] += 1
        g["due"] += float(r.get("应缴金额（元）") or 0)
        g["paid"] += float(r.get("实缴金额（元）") or 0)
    return out


def asset_key(a):
    try:
        return int(float(a.get("id")))
    except (TypeError, ValueError):
        return None


def receivable_diff(a, totals):
    """资产合同待收差额（元）= Σ应缴 − Σ实缴，与 read_assets 的计算列同口径。"""
    t = totals.get(asset_key(a)) or {"due": 0.0, "paid": 0.0}
    return round(t["due"] - t["paid"], 2)


def parse_date(v):
    s = str(v or "").strip()
    if not s:
        return None
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if not m:
        return None
    try:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def remind_days(row, default_days, field="提醒提前天数"):
    v = num(row.get(field))
    return default_days if v is None else int(v)


def is_due(row, date_field, default_days, today, days_field="提醒提前天数"):
    """到期判定：日期非空 且 今天 >= 日期 − 提前天数（空取全局默认）。"""
    d = parse_date(row.get(date_field))
    if not d:
        return None
    rd = remind_days(row, default_days, days_field)
    if today < d - datetime.timedelta(days=rd):
        return None
    return (d - today).days


def pending_unpaid_project(c):
    """待支付金额：优先取「未支付」（含 0）；为空时按 合同金额 − 已支付（空按 0）。"""
    up = num(c.get("未支付"))
    if up is not None:
        return up
    amt = num(c.get("合同金额（万元）")) or 0
    paid = num(c.get("已支付")) or 0
    return round(amt - paid, 2)


def archived_project(c):
    return str(c.get("是否归档", "") or "").strip() == "是"


def in_scope(dept, visible):
    return visible is None or str(dept or "").strip() in visible
