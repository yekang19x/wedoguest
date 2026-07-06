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
GUEST_HEADERS = ["ID", "姓名", "总人数", "家属姓名", "桌号", "邀请函状态", "确认状态", "备注"]
TABLE_HEADERS = ["桌号", "备注名", "容纳人数"]
CONFIG_HEADERS = ["配置项", "值"]

INVITE_STATUSES = ["未发送", "已发送"]
CONFIRM_STATUSES = ["待确认", "已确认", "不参加"]

DEFAULT_CONFIG = {"default_capacity": 10, "budget_total": 100}


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
            {"table_no": "主桌", "label": "新人与主宾", "capacity": 12},
            {"table_no": "1", "label": "男方亲戚", "capacity": None},
            {"table_no": "2", "label": "女方亲戚", "capacity": None},
            {"table_no": "3", "label": "同事朋友", "capacity": None},
        ])
    if not GUESTS_XLSX.exists():
        save_guests([
            {"id": 1, "name": "张伟", "party_size": 3, "family_names": "李娜,张小宝",
             "table_no": "1", "invite_status": "已发送", "confirm_status": "已确认", "note": "叔叔一家"},
            {"id": 2, "name": "王芳", "party_size": 2, "family_names": "刘强",
             "table_no": "2", "invite_status": "已发送", "confirm_status": "待确认", "note": ""},
            {"id": 3, "name": "陈静", "party_size": 1, "family_names": "",
             "table_no": "", "invite_status": "未发送", "confirm_status": "待确认", "note": "大学同学"},
        ])


# ---------- 宾客（xlsx） ----------

def load_guests() -> list[dict]:
    wb = load_workbook(GUESTS_XLSX)
    ws = wb.active
    guests = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or row[0] is None:
            continue
        # 兼容手工编辑 Excel 时留下的短行
        vals = list(row) + [None] * (len(GUEST_HEADERS) - len(row))
        guests.append({
            "id": int(vals[0]),
            "name": str(vals[1] or "").strip(),
            "party_size": int(vals[2] or 1),
            "family_names": str(vals[3] or "").strip(),
            "table_no": str(vals[4] or "").strip(),
            "invite_status": str(vals[5] or "未发送").strip(),
            "confirm_status": str(vals[6] or "待确认").strip(),
            "note": str(vals[7] or "").strip(),
        })
    return guests


def save_guests(guests: list[dict]):
    wb = Workbook()
    ws = wb.active
    ws.title = "宾客名单"
    ws.append(GUEST_HEADERS)
    for g in guests:
        ws.append([
            g["id"], g["name"], g["party_size"], g["family_names"],
            g["table_no"], g["invite_status"], g["confirm_status"], g["note"],
        ])
    # 设置列宽，直接用 Excel 打开时可读性更好
    widths = [6, 14, 8, 24, 8, 12, 10, 24]
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
            cap_raw = (row.get("容纳人数") or "").strip()
            tables.append({
                "table_no": (row.get("桌号") or "").strip(),
                "label": (row.get("备注名") or "").strip(),
                "capacity": int(cap_raw) if cap_raw else None,
            })
    return [t for t in tables if t["table_no"]]


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
                config[key] = int(val)
    return config


def save_config(config: dict):
    tmp = CONFIG_CSV.with_suffix(".csv.tmp")
    with open(tmp, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(CONFIG_HEADERS)
        for key, val in config.items():
            writer.writerow([key, val])
    _atomic_replace(tmp, CONFIG_CSV)
