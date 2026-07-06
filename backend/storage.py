"""数据存储层。

主要数据（宾客名单）存 Excel xlsx 文件，次要数据（桌子属性、全局配置）存 csv 文件。
所有写入先落临时文件再原子替换，避免写一半导致文件损坏。
"""

import csv
import os
from pathlib import Path

from openpyxl import Workbook, load_workbook

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
GUESTS_XLSX = DATA_DIR / "guests.xlsx"
TABLES_CSV = DATA_DIR / "tables.csv"
CONFIG_CSV = DATA_DIR / "config.csv"

# guests.xlsx 表头（中文列名便于直接用 Excel 打开维护）
GUEST_HEADERS = ["ID", "姓名", "预算人数", "确认人数", "家属姓名", "桌号", "邀请函状态", "确认状态", "备注"]
# 旧版列名 → 新版列名（按表头读取，兼容旧文件平滑升级）
GUEST_HEADER_ALIASES = {"总人数": "预算人数"}
TABLE_HEADERS = ["桌号", "备注名", "容纳人数", "X坐标", "Y坐标"]
CONFIG_HEADERS = ["配置项", "值"]

INVITE_STATUSES = ["未发送", "已发送"]
CONFIRM_STATUSES = ["待确认", "已确认", "不参加"]

DEFAULT_CONFIG = {
    "default_capacity": 10,   # 默认单桌容纳人数
    "budget_total": 100,      # 人数预算
    "venue_width": 18.0,      # 会场宽度（米，横向，舞台所在边）
    "venue_depth": 25.0,      # 会场长度（米，纵向）
    "table_diameter": 1.8,    # 桌子直径（米）
}
CONFIG_INT_KEYS = {"default_capacity", "budget_total"}


# ---------- 通用 ----------

def _atomic_replace(tmp_path: Path, final_path: Path):
    os.replace(tmp_path, final_path)


def ensure_data_files():
    """数据文件不存在时初始化（含少量示例数据，便于首次查看效果）。"""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not CONFIG_CSV.exists():
        save_config(dict(DEFAULT_CONFIG))
    if not TABLES_CSV.exists():
        save_tables([
            # 主桌演示手动摆放（T 台右侧靠舞台）；数字桌不设坐标，由平面图按桌号自动排列
            {"table_no": "主桌", "label": "新人与主宾", "capacity": 12, "x": 12.5, "y": 5.5},
            {"table_no": "1", "label": "男方亲戚", "capacity": None, "x": None, "y": None},
            {"table_no": "2", "label": "女方亲戚", "capacity": None, "x": None, "y": None},
            {"table_no": "3", "label": "同事朋友", "capacity": None, "x": None, "y": None},
        ])
    if not GUESTS_XLSX.exists():
        save_guests([
            {"id": 1, "name": "张伟", "party_size": 3, "confirmed_size": 2, "family_names": "李娜,张小宝",
             "table_no": "1", "invite_status": "已发送", "confirm_status": "已确认", "note": "叔叔一家"},
            {"id": 2, "name": "王芳", "party_size": 2, "confirmed_size": 2, "family_names": "刘强",
             "table_no": "2", "invite_status": "已发送", "confirm_status": "待确认", "note": ""},
            {"id": 3, "name": "陈静", "party_size": 1, "confirmed_size": 1, "family_names": "",
             "table_no": "", "invite_status": "未发送", "confirm_status": "待确认", "note": "大学同学"},
        ])


# ---------- 宾客（xlsx） ----------

def load_guests() -> list[dict]:
    wb = load_workbook(GUESTS_XLSX)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    # 按表头名定位列，兼容旧版列名与用户在 Excel 里调整过的列顺序
    header = next(rows, None) or ()
    col = {}
    for i, h in enumerate(header):
        name = str(h or "").strip()
        col[GUEST_HEADER_ALIASES.get(name, name)] = i

    def cell(vals, name):
        i = col.get(name)
        return vals[i] if i is not None and i < len(vals) else None

    guests = []
    for row in rows:
        if row is None or cell(row, "ID") is None:
            continue
        party_size = int(cell(row, "预算人数") or 1)
        conf_raw = cell(row, "确认人数")
        conf_str = str(conf_raw).strip() if conf_raw is not None else ""
        guests.append({
            "id": int(cell(row, "ID")),
            "name": str(cell(row, "姓名") or "").strip(),
            "party_size": party_size,
            # 旧文件无此列 / 单元格留空 → 默认等于预算人数
            "confirmed_size": int(float(conf_str)) if conf_str else party_size,
            "family_names": str(cell(row, "家属姓名") or "").strip(),
            "table_no": str(cell(row, "桌号") or "").strip(),
            "invite_status": str(cell(row, "邀请函状态") or "未发送").strip(),
            "confirm_status": str(cell(row, "确认状态") or "待确认").strip(),
            "note": str(cell(row, "备注") or "").strip(),
        })
    return guests


def save_guests(guests: list[dict]):
    wb = Workbook()
    ws = wb.active
    ws.title = "宾客名单"
    ws.append(GUEST_HEADERS)
    for g in guests:
        ws.append([
            g["id"], g["name"], g["party_size"], g["confirmed_size"], g["family_names"],
            g["table_no"], g["invite_status"], g["confirm_status"], g["note"],
        ])
    # 设置列宽，直接用 Excel 打开时可读性更好
    widths = [6, 14, 10, 10, 24, 8, 12, 10, 24]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"
    tmp = GUESTS_XLSX.with_suffix(".xlsx.tmp")
    wb.save(tmp)
    _atomic_replace(tmp, GUESTS_XLSX)


def next_guest_id(guests: list[dict]) -> int:
    return max((g["id"] for g in guests), default=0) + 1


# ---------- 桌子（csv） ----------

def load_tables() -> list[dict]:
    tables = []
    with open(TABLES_CSV, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            def num(key):
                raw = (row.get(key) or "").strip()
                return float(raw) if raw else None

            cap = num("容纳人数")
            tables.append({
                "table_no": (row.get("桌号") or "").strip(),
                "label": (row.get("备注名") or "").strip(),
                "capacity": int(cap) if cap is not None else None,
                "x": num("X坐标"),   # 桌心坐标（米），空 = 未摆放，由前端自动布局
                "y": num("Y坐标"),
            })
    return [t for t in tables if t["table_no"]]


def _fmt_num(v):
    if v is None:
        return ""
    return int(v) if float(v).is_integer() else round(float(v), 1)


def save_tables(tables: list[dict]):
    tmp = TABLES_CSV.with_suffix(".csv.tmp")
    # utf-8-sig：保证用 Excel 直接打开 csv 不乱码
    with open(tmp, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(TABLE_HEADERS)
        for t in tables:
            writer.writerow([
                t["table_no"], t["label"],
                "" if t["capacity"] is None else t["capacity"],
                _fmt_num(t.get("x")), _fmt_num(t.get("y")),
            ])
    _atomic_replace(tmp, TABLES_CSV)


# ---------- 全局配置（csv） ----------

def load_config() -> dict:
    config = dict(DEFAULT_CONFIG)
    with open(CONFIG_CSV, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            key = (row.get("配置项") or "").strip()
            val = (row.get("值") or "").strip()
            if key in config and val:
                config[key] = int(float(val)) if key in CONFIG_INT_KEYS else float(val)
    return config


def save_config(config: dict):
    tmp = CONFIG_CSV.with_suffix(".csv.tmp")
    with open(tmp, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(CONFIG_HEADERS)
        for key, val in config.items():
            writer.writerow([key, _fmt_num(val) if key not in CONFIG_INT_KEYS else val])
    _atomic_replace(tmp, CONFIG_CSV)
