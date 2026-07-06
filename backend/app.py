"""婚礼宾客统计系统 — FastAPI 后端。

数据存 data/wedding.db（SQLite）；宾客名单支持 xlsx 导入/导出。
"""

from io import BytesIO
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook, load_workbook
from pydantic import BaseModel, Field

import storage
from storage import CONFIRM_STATUSES, INVITE_STATUSES

app = FastAPI(title="婚礼宾客统计系统")

FRONTEND_DIR = Path(__file__).resolve().parent.parent / "frontend"

storage.ensure_data_files()


# ---------- 请求模型 ----------

class GuestIn(BaseModel):
    name: str = Field(min_length=1)
    party_size: int = Field(default=1, ge=1)          # 预算人数（含本人）
    confirmed_size: int | None = Field(default=None, ge=0)  # 确认会到人数，None = 等于预算人数
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
    x: float | None = Field(default=None, ge=0)       # 桌心坐标（米），None = 未摆放
    y: float | None = Field(default=None, ge=0)


class ConfigIn(BaseModel):
    default_capacity: int = Field(ge=1)
    budget_total: int = Field(ge=0)
    venue_width: float = Field(default=18.0, gt=0, le=200)   # 会场宽（米，横向）
    venue_depth: float = Field(default=25.0, gt=0, le=200)   # 会场长（米，纵向）
    table_diameter: float = Field(default=1.8, gt=0, le=10)  # 桌子直径（米）
    table_gap: float = Field(default=1.2, ge=0, le=20)       # 自动排列桌边间距（米）


# ---------- 内部工具 ----------

def _validate_statuses(invite_status: str, confirm_status: str):
    if invite_status not in INVITE_STATUSES:
        raise HTTPException(400, f"邀请函状态必须是 {INVITE_STATUSES} 之一")
    if confirm_status not in CONFIRM_STATUSES:
        raise HTTPException(400, f"确认状态必须是 {CONFIRM_STATUSES} 之一")


def _seat_size(g: dict) -> int:
    """占座人数：已确认按确认人数，待确认按预算人数，不参加不占座。"""
    if g["confirm_status"] == "不参加":
        return 0
    if g["confirm_status"] == "已确认":
        return g["confirmed_size"]
    return g["party_size"]


def _table_occupied(guests: list[dict], table_no: str, exclude_ids: set[int] = frozenset()) -> int:
    return sum(
        _seat_size(g) for g in guests
        if g["table_no"] == table_no and g["id"] not in exclude_ids
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
    # 已确认宾客按「确认人数」计，待确认按「预算人数」计，不参加不计座
    confirmed = sum(g["confirmed_size"] for g in guests if g["confirm_status"] == "已确认")
    pending = sum(g["party_size"] for g in guests if g["confirm_status"] == "待确认")
    declined = sum(g["party_size"] for g in guests if g["confirm_status"] == "不参加")
    return {
        "budget_total": config["budget_total"],
        "invited_total": sum(g["party_size"] for g in guests),  # 预算合计
        "confirmed": confirmed,
        "pending": pending,
        "declined": declined,
        "expected": confirmed + pending,
        "seated": sum(_seat_size(g) for g in guests if g["table_no"]),
        "unseated": sum(_seat_size(g) for g in guests if not g["table_no"]),
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

def _guest_fields(body: GuestIn) -> dict:
    return {
        "name": body.name.strip(),
        "party_size": body.party_size,
        "confirmed_size": body.confirmed_size if body.confirmed_size is not None else body.party_size,
        "family_names": body.family_names.strip(),
        "table_no": body.table_no.strip(),
        "invite_status": body.invite_status,
        "confirm_status": body.confirm_status,
        "note": body.note.strip(),
    }


@app.post("/api/guests")
def create_guest(body: GuestIn):
    _validate_statuses(body.invite_status, body.confirm_status)
    guests = storage.load_guests()
    tables = storage.load_tables()
    config = storage.load_config()
    fields = _guest_fields(body)
    guest = {"id": storage.next_guest_id(guests), **fields}
    _check_capacity(guests, tables, config, guest["table_no"], _seat_size(guest), force=body.force)
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
    fields = _guest_fields(body)
    _check_capacity(guests, tables, config, fields["table_no"], _seat_size(fields),
                    exclude_ids={gid}, force=body.force)
    guest.update(fields)
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
    if split:
        # 本人 1 人换桌；确认人数同步拆分（本人至多占 1）
        person_confirmed = min(1, guest["confirmed_size"])
        moving = {**guest, "party_size": 1, "confirmed_size": person_confirmed}
    else:
        moving = guest
    _check_capacity(guests, tables, config, target, _seat_size(moving),
                    exclude_ids={gid}, force=body.force)

    result = {"moved": None, "family_left": None}
    if split:
        family = {
            "id": storage.next_guest_id(guests),
            "name": f"{guest['name']}的家属",
            "party_size": guest["party_size"] - 1,
            "confirmed_size": guest["confirmed_size"] - min(1, guest["confirmed_size"]),
            "family_names": guest["family_names"],
            "table_no": guest["table_no"],
            "invite_status": guest["invite_status"],
            "confirm_status": guest["confirm_status"],
            "note": f"由「{guest['name']}」换桌拆分",
        }
        guests.append(family)
        guest.update({"party_size": 1, "confirmed_size": min(1, guest["confirmed_size"]),
                      "family_names": "", "table_no": target})
        result["family_left"] = family
    else:
        guest["table_no"] = target
    result["moved"] = guest
    storage.save_guests(guests)
    return result


# ---------- 宾客名单导入 / 导出（xlsx） ----------
# 交换格式：姓名 | 配偶 | 子女1..N | 预计人数 | 确认人数 | 桌号
# 家属名平铺为多列（第一位视为配偶，其余为子女），按预计人数-1 补足空位。

EXPORT_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _norm_cell(v) -> str:
    """单元格转字符串；Excel 数字桌号 1.0 规整为 '1'。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _cell_int(v) -> int | None:
    s = _norm_cell(v)
    try:
        return int(float(s)) if s else None
    except ValueError:
        return None


@app.get("/api/guests/export")
def export_guests():
    guests = storage.load_guests()
    padded = []
    for g in guests:
        fam = [s.strip() for s in g["family_names"].split(",") if s.strip()]
        fam += [""] * max(0, g["party_size"] - 1 - len(fam))  # 按人数补全空位
        padded.append(fam)
    # 至少保留 配偶+2子女 列，空名单导出也可当录入模板
    max_family = max([len(f) for f in padded] + [3])
    headers = ["姓名", "配偶"] + [f"子女{i}" for i in range(1, max_family)] + \
              ["预计人数", "确认人数", "桌号"]

    wb = Workbook()
    ws = wb.active
    ws.title = "宾客名单"
    ws.append(headers)
    for g, fam in zip(guests, padded):
        fam_cells = fam + [""] * (max_family - len(fam))
        ws.append([g["name"], *fam_cells, g["party_size"], g["confirmed_size"], g["table_no"]])
    for i in range(1, len(headers) + 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 12
    ws.freeze_panes = "A2"
    buf = BytesIO()
    wb.save(buf)
    return Response(buf.getvalue(), media_type=EXPORT_MIME, headers={
        "Content-Disposition": f"attachment; filename*=UTF-8''{quote('宾客名单.xlsx')}",
    })


@app.post("/api/guests/import")
async def import_guests(request: Request, mode: str = "append"):
    """导入宾客名单。body 为 xlsx 原始字节；mode=append 追加 / replace 覆盖全部。"""
    if mode not in ("append", "replace"):
        raise HTTPException(400, "mode 必须是 append 或 replace")
    data = await request.body()
    if not data:
        raise HTTPException(400, "请上传 xlsx 文件")
    try:
        ws = load_workbook(BytesIO(data), data_only=True).active
    except Exception:
        raise HTTPException(400, "无法解析文件，请上传 .xlsx 格式的 Excel 文件")

    rows = ws.iter_rows(values_only=True)
    header = next(rows, None) or ()
    col = {}
    for i, h in enumerate(header):
        name = str(h or "").strip()
        if name:
            col[name] = i
    if "姓名" not in col:
        raise HTTPException(400, "缺少「姓名」列（表头需含：姓名、配偶、子女N、预计人数、确认人数、桌号）")
    family_cols = sorted(i for name, i in col.items() if name == "配偶" or name.startswith("子女"))

    def cell(vals, name):
        i = col.get(name)
        return vals[i] if i is not None and i < len(vals) else None

    guests = [] if mode == "replace" else storage.load_guests()
    next_id = storage.next_guest_id(guests)
    tables = storage.load_tables()
    known_tables = {t["table_no"] for t in tables}
    imported, skipped, unknown_tables = 0, 0, []

    for row in rows:
        if row is None or all(v is None or str(v).strip() == "" for v in row):
            continue  # 整行空白直接忽略
        name = _norm_cell(cell(row, "姓名"))
        if not name:
            skipped += 1
            continue
        family = [_norm_cell(row[i]) for i in family_cols
                  if i < len(row) and _norm_cell(row[i])]
        party_raw = _cell_int(cell(row, "预计人数"))
        # 人数缺省按 本人+家属数 推断；显式填写时不小于已列出的家属数
        party = max(party_raw or 0, 1 + len(family)) if (party_raw or family) else 1
        conf_raw = _cell_int(cell(row, "确认人数"))
        confirmed = min(conf_raw, party) if conf_raw is not None else party
        table_no = _norm_cell(cell(row, "桌号"))
        if table_no and table_no not in known_tables:
            unknown_tables.append(table_no)
            known_tables.add(table_no)
        guests.append({
            "id": next_id,
            "name": name,
            "party_size": party,
            "confirmed_size": max(0, confirmed),
            "family_names": ",".join(family),
            "table_no": table_no,
            "invite_status": "未发送",
            "confirm_status": "待确认",
            "note": "",
        })
        next_id += 1
        imported += 1

    if imported == 0 and skipped == 0:
        raise HTTPException(400, "文件里没有可导入的数据行")
    # 名单里出现的新桌号自动建桌（容量用全局默认，位置走自动布局）
    for no in unknown_tables:
        tables.append({"table_no": no, "label": "", "capacity": None, "x": None, "y": None})
    if unknown_tables:
        storage.save_tables(tables)
    storage.save_guests(guests)
    return {"imported": imported, "skipped": skipped,
            "created_tables": unknown_tables, "mode": mode}


# ---------- 桌子 ----------

@app.post("/api/tables")
def create_table(body: TableIn):
    tables = storage.load_tables()
    no = body.table_no.strip()
    if any(t["table_no"] == no for t in tables):
        raise HTTPException(409, f"桌号「{no}」已存在")
    table = {"table_no": no, "label": body.label.strip(), "capacity": body.capacity,
             "x": body.x, "y": body.y}
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
    table.update({"table_no": new_no, "label": body.label.strip(), "capacity": body.capacity,
                  "x": body.x, "y": body.y})
    storage.save_tables(tables)
    return table


@app.post("/api/tables/reset-positions")
def reset_table_positions():
    """清空所有桌子坐标 → 平面图恢复按桌号有序的自动排列。"""
    tables = storage.load_tables()
    for t in tables:
        t["x"] = None
        t["y"] = None
    storage.save_tables(tables)
    return {"ok": True, "count": len(tables)}


@app.put("/api/tables/{table_no}/position")
def move_table(table_no: str, body: dict):
    """仅更新桌子坐标（平面图拖动摆位）。"""
    tables = storage.load_tables()
    table = next((t for t in tables if t["table_no"] == table_no), None)
    if table is None:
        raise HTTPException(404, f"桌号「{table_no}」不存在")
    try:
        table["x"] = round(float(body["x"]), 1)
        table["y"] = round(float(body["y"]), 1)
    except (KeyError, TypeError, ValueError):
        raise HTTPException(400, "需要数字坐标 x、y（单位米）")
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
    config = {
        "default_capacity": body.default_capacity,
        "budget_total": body.budget_total,
        "venue_width": body.venue_width,
        "venue_depth": body.venue_depth,
        "table_diameter": body.table_diameter,
        "table_gap": body.table_gap,
    }
    storage.save_config(config)
    return config


# 静态页面托管（放最后，避免遮住 /api 路由）
app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")
