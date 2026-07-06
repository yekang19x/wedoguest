"""婚礼宾客统计系统 — FastAPI 后端。

宾客主数据存 data/guests.xlsx，桌子属性存 data/tables.csv，全局配置存 data/config.csv。
"""

from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

import storage
from storage import CONFIRM_STATUSES, INVITE_STATUSES

app = FastAPI(title="婚礼宾客统计系统")

FRONTEND_DIR = Path(__file__).resolve().parent.parent / "frontend"

storage.ensure_data_files()


# ---------- 请求模型 ----------

class GuestIn(BaseModel):
    name: str = Field(min_length=1)
    party_size: int = Field(default=1, ge=1)
    family_names: str = ""
    table_no: str = ""
    invite_status: str = "未发送"
    confirm_status: str = "待确认"
    note: str = ""
    force: bool = False  # 目标桌超容时是否强制安排


class MoveIn(BaseModel):
    target_table: str = ""  # 空字符串 = 移到未安排区
    with_family: bool = True
    force: bool = False


class TableIn(BaseModel):
    table_no: str = Field(min_length=1)
    label: str = ""
    capacity: int | None = Field(default=None, ge=1)  # None = 使用全局默认容量


class ConfigIn(BaseModel):
    default_capacity: int = Field(ge=1)
    budget_total: int = Field(ge=0)


# ---------- 内部工具 ----------

def _validate_statuses(invite_status: str, confirm_status: str):
    if invite_status not in INVITE_STATUSES:
        raise HTTPException(400, f"邀请函状态必须是 {INVITE_STATUSES} 之一")
    if confirm_status not in CONFIRM_STATUSES:
        raise HTTPException(400, f"确认状态必须是 {CONFIRM_STATUSES} 之一")


def _table_occupied(guests: list[dict], table_no: str, exclude_ids: set[int] = frozenset()) -> int:
    """某桌当前占用人数。不参加的宾客不占座位。"""
    return sum(
        g["party_size"] for g in guests
        if g["table_no"] == table_no and g["id"] not in exclude_ids
        and g["confirm_status"] != "不参加"
    )


def _effective_capacity(table: dict, config: dict) -> int:
    return table["capacity"] if table["capacity"] is not None else config["default_capacity"]


def _check_capacity(guests, tables, config, table_no: str, incoming: int,
                    exclude_ids: set[int] = frozenset(), force: bool = False):
    """校验目标桌容量。桌号不存在报 404，超容且未 force 报 409。"""
    if not table_no:
        return
    table = next((t for t in tables if t["table_no"] == table_no), None)
    if table is None:
        raise HTTPException(404, f"桌号「{table_no}」不存在")
    if force:
        return
    cap = _effective_capacity(table, config)
    occupied = _table_occupied(guests, table_no, exclude_ids)
    if occupied + incoming > cap:
        raise HTTPException(
            409,
            f"桌「{table_no}」容量不足：容纳 {cap} 人，已坐 {occupied} 人，"
            f"再安排 {incoming} 人将超出。可强制安排。",
        )


def _find_guest(guests: list[dict], gid: int) -> dict:
    guest = next((g for g in guests if g["id"] == gid), None)
    if guest is None:
        raise HTTPException(404, f"宾客 ID {gid} 不存在")
    return guest


def _build_stats(guests: list[dict], config: dict) -> dict:
    def size_sum(pred):
        return sum(g["party_size"] for g in guests if pred(g))

    confirmed = size_sum(lambda g: g["confirm_status"] == "已确认")
    pending = size_sum(lambda g: g["confirm_status"] == "待确认")
    declined = size_sum(lambda g: g["confirm_status"] == "不参加")
    attending = lambda g: g["confirm_status"] != "不参加"
    return {
        "budget_total": config["budget_total"],
        "invited_total": size_sum(lambda g: True),
        "confirmed": confirmed,
        "pending": pending,
        "declined": declined,
        "expected": confirmed + pending,
        "seated": size_sum(lambda g: attending(g) and g["table_no"]),
        "unseated": size_sum(lambda g: attending(g) and not g["table_no"]),
        "guest_count": len(guests),
        "invite_sent": sum(1 for g in guests if g["invite_status"] == "已发送"),
        "invite_unsent": sum(1 for g in guests if g["invite_status"] == "未发送"),
    }


# ---------- 总览 ----------

@app.get("/api/overview")
def overview():
    guests = storage.load_guests()
    tables = storage.load_tables()
    config = storage.load_config()
    table_views = []
    for t in tables:
        members = [g for g in guests if g["table_no"] == t["table_no"]]
        table_views.append({
            **t,
            "effective_capacity": _effective_capacity(t, config),
            "occupied": _table_occupied(guests, t["table_no"]),
            "guests": members,
        })
    known_tables = {t["table_no"] for t in tables}
    unassigned = [g for g in guests if not g["table_no"] or g["table_no"] not in known_tables]
    return {
        "guests": guests,
        "tables": table_views,
        "unassigned_guests": unassigned,
        "config": config,
        "stats": _build_stats(guests, config),
    }


# ---------- 宾客 ----------

@app.post("/api/guests")
def create_guest(body: GuestIn):
    _validate_statuses(body.invite_status, body.confirm_status)
    guests = storage.load_guests()
    tables = storage.load_tables()
    config = storage.load_config()
    incoming = body.party_size if body.confirm_status != "不参加" else 0
    _check_capacity(guests, tables, config, body.table_no, incoming, force=body.force)
    guest = {
        "id": storage.next_guest_id(guests),
        "name": body.name.strip(),
        "party_size": body.party_size,
        "family_names": body.family_names.strip(),
        "table_no": body.table_no.strip(),
        "invite_status": body.invite_status,
        "confirm_status": body.confirm_status,
        "note": body.note.strip(),
    }
    guests.append(guest)
    storage.save_guests(guests)
    return guest


@app.put("/api/guests/{gid}")
def update_guest(gid: int, body: GuestIn):
    _validate_statuses(body.invite_status, body.confirm_status)
    guests = storage.load_guests()
    tables = storage.load_tables()
    config = storage.load_config()
    guest = _find_guest(guests, gid)
    incoming = body.party_size if body.confirm_status != "不参加" else 0
    _check_capacity(guests, tables, config, body.table_no, incoming,
                    exclude_ids={gid}, force=body.force)
    guest.update({
        "name": body.name.strip(),
        "party_size": body.party_size,
        "family_names": body.family_names.strip(),
        "table_no": body.table_no.strip(),
        "invite_status": body.invite_status,
        "confirm_status": body.confirm_status,
        "note": body.note.strip(),
    })
    storage.save_guests(guests)
    return guest


@app.delete("/api/guests/{gid}")
def delete_guest(gid: int):
    guests = storage.load_guests()
    guest = _find_guest(guests, gid)
    guests.remove(guest)
    storage.save_guests(guests)
    return {"ok": True}


@app.post("/api/guests/{gid}/move")
def move_guest(gid: int, body: MoveIn):
    """换桌。with_family=False 且随行人数 > 1 时拆分记录：本人换桌，家属留在原桌。"""
    guests = storage.load_guests()
    tables = storage.load_tables()
    config = storage.load_config()
    guest = _find_guest(guests, gid)
    target = body.target_table.strip()

    split = not body.with_family and guest["party_size"] > 1
    moving_size = 1 if split else guest["party_size"]
    if guest["confirm_status"] == "不参加":
        moving_size = 0
    _check_capacity(guests, tables, config, target, moving_size,
                    exclude_ids={gid}, force=body.force)

    result = {"moved": None, "family_left": None}
    if split:
        family = {
            "id": storage.next_guest_id(guests),
            "name": f"{guest['name']}的家属",
            "party_size": guest["party_size"] - 1,
            "family_names": guest["family_names"],
            "table_no": guest["table_no"],
            "invite_status": guest["invite_status"],
            "confirm_status": guest["confirm_status"],
            "note": f"由「{guest['name']}」换桌拆分",
        }
        guests.append(family)
        guest.update({"party_size": 1, "family_names": "", "table_no": target})
        result["family_left"] = family
    else:
        guest["table_no"] = target
    result["moved"] = guest
    storage.save_guests(guests)
    return result


# ---------- 桌子 ----------

@app.post("/api/tables")
def create_table(body: TableIn):
    tables = storage.load_tables()
    no = body.table_no.strip()
    if any(t["table_no"] == no for t in tables):
        raise HTTPException(409, f"桌号「{no}」已存在")
    table = {"table_no": no, "label": body.label.strip(), "capacity": body.capacity}
    tables.append(table)
    storage.save_tables(tables)
    return table


@app.put("/api/tables/{table_no}")
def update_table(table_no: str, body: TableIn):
    tables = storage.load_tables()
    table = next((t for t in tables if t["table_no"] == table_no), None)
    if table is None:
        raise HTTPException(404, f"桌号「{table_no}」不存在")
    new_no = body.table_no.strip()
    if new_no != table_no:
        if any(t["table_no"] == new_no for t in tables):
            raise HTTPException(409, f"桌号「{new_no}」已存在")
        # 重命名桌号时同步更新该桌所有宾客
        guests = storage.load_guests()
        for g in guests:
            if g["table_no"] == table_no:
                g["table_no"] = new_no
        storage.save_guests(guests)
    table.update({"table_no": new_no, "label": body.label.strip(), "capacity": body.capacity})
    storage.save_tables(tables)
    return table


@app.delete("/api/tables/{table_no}")
def delete_table(table_no: str):
    tables = storage.load_tables()
    table = next((t for t in tables if t["table_no"] == table_no), None)
    if table is None:
        raise HTTPException(404, f"桌号「{table_no}」不存在")
    tables.remove(table)
    storage.save_tables(tables)
    # 该桌宾客移入未安排区
    guests = storage.load_guests()
    changed = False
    for g in guests:
        if g["table_no"] == table_no:
            g["table_no"] = ""
            changed = True
    if changed:
        storage.save_guests(guests)
    return {"ok": True}


# ---------- 全局配置 ----------

@app.put("/api/config")
def update_config(body: ConfigIn):
    config = {"default_capacity": body.default_capacity, "budget_total": body.budget_total}
    storage.save_config(config)
    return config


# 静态页面托管（放最后，避免遮住 /api 路由）
app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")
