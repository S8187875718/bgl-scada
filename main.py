import os
import io
import sqlite3
from datetime import datetime
import openpyxl
from typing import Dict, List, Any
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel
import networkx as nx

app = FastAPI(title="BGL Gas SCADA Engine")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_HTML = os.path.join(BASE_DIR, "templates", "index.html")
DB_PATH = os.path.join(BASE_DIR, "scada.db")
MASTER_EXCEL = os.path.join(BASE_DIR, "master_data.xlsx")

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def load_excel_bytes_to_db(contents: bytes, cursor):
    wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=True)
    now = datetime.now().strftime("%Y-%m-%d")

    cursor.execute("DELETE FROM assets")
    cursor.execute("DELETE FROM pipelines")

    assets_sheet = None
    for s in wb.sheetnames:
        if s.strip().lower() == "assets":
            assets_sheet = wb[s]
            break
    if not assets_sheet:
        assets_sheet = wb.active

    asset_count = 0
    if assets_sheet:
        headers = [str(c.value or "").strip() for c in assets_sheet[1]]
        col_map = {name: idx for idx, name in enumerate(headers) if name}
        rows = []
        for row in assets_sheet.iter_rows(min_row=2, values_only=True):
            if not row or len(row) == 0:
                continue
            raw_id = row[col_map.get("Asset ID", 0)]
            if raw_id is None or str(raw_id).strip() == "":
                continue
            aid = str(raw_id).strip()
            name = str(row[col_map.get("Asset Name", 1)] or aid).strip()
            atype = str(row[col_map.get("Type", 2)] or "VALVE").strip().upper()
            tag = str(row[col_map.get("Tag", 3)] or "125 MM").strip()
            details = str(row[col_map.get("Details", 4)] or "").strip() if len(row) > 4 else ""
            status = str(row[col_map.get("Status", 7)] or "ON").strip().upper() if len(row) > 7 else "ON"
            if status not in ["ON", "OFF"]:
                status = "ON"
            loc = str(row[col_map.get("Location", 8)] or "").strip() if len(row) > 8 else ""
            if loc == "None":
                loc = ""

            press = 4.0 if atype == "DRS" else 0.0
            if atype == "DRS" and "bar" in details.lower():
                try:
                    press = float("".join(ch for ch in details if ch.isdigit() or ch == "."))
                except:
                    pass
            rows.append((aid, name, atype, tag, details, now, now, status, loc, press))
        cursor.executemany("INSERT OR REPLACE INTO assets VALUES (?,?,?,?,?,?,?,?,?,?)", rows)
        asset_count = len(rows)

    links_sheet = None
    for s in wb.sheetnames:
        if s.strip().lower() in ["links", "link", "sheet2"]:
            links_sheet = wb[s]
            break

    link_count = 0
    if links_sheet:
        headers = [str(c.value or "").strip() for c in links_sheet[1]]
        col_map = {name: idx for idx, name in enumerate(headers) if name}
        p_idx = col_map.get("Parent ID", 0)
        c_idx = col_map.get("Child ID", 1)
        links = []
        for row in links_sheet.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 2:
                continue
            src = str(row[p_idx] or "").strip() if row[p_idx] is not None else ""
            tgt = str(row[c_idx] or "").strip() if row[c_idx] is not None else ""
            if src and tgt and src != tgt and src != "None" and tgt != "None":
                links.append((src, tgt, now))
        cursor.executemany("INSERT OR IGNORE INTO pipelines VALUES (?,?,?)", links)
        link_count = len(links)

    return asset_count, link_count

def seed_default_network(cursor):
    # If master_data.xlsx exists in project folder, auto-load it permanently!
    if os.path.exists(MASTER_EXCEL):
        try:
            with open(MASTER_EXCEL, "rb") as f:
                load_excel_bytes_to_db(f.read(), cursor)
            return
        except Exception as e:
            print("Master excel load error:", e)

    now = datetime.now().strftime("%Y-%m-%d")
    default_assets = [
        ("D01", "MS DRS", "DRS", "125 MM", "4.1 Bar", now, now, "ON", "Shamirpet, Hyderabad", 4.1),
        ("D03", "Check Post DRS", "DRS", "125 MM", "3.6 Bar", now, now, "ON", "Medchal Check Post, Hyderabad", 3.6),
        ("D02", "Depo DRS", "DRS", "125 MM", "3.1 Bar", now, now, "ON", "Medchal Depo, Hyderabad", 3.1),
        ("D01-DIR-01", "Registrar, Nalsar Univ.", "CPNG", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE01", "CG-D01-PE01", "VALVE", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE02", "CG-D01-PE02", "VALVE", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("31000003", "IPE Univ. Canteen", "CPNG", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("S03-D03-PE01", "S03-D03-PE01", "VALVE", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("S02-D02-PE01", "S02-D02-PE01", "VALVE", "125 MM", "", now, now, "ON", "Medchal", 0.0)
    ]
    cursor.executemany("INSERT OR REPLACE INTO assets VALUES (?,?,?,?,?,?,?,?,?,?)", default_assets)

    default_pipelines = [
        ("D01", "CG-D01-PE01", now), ("CG-D01-PE01", "31000003", now),
        ("D01", "CG-D01-PE02", now), ("D03", "S03-D03-PE01", now),
        ("D02", "S02-D02-PE01", now), ("CG-D01-PE02", "S03-D03-PE01", now)
    ]
    cursor.executemany("INSERT OR REPLACE INTO pipelines VALUES (?,?,?)", default_pipelines)

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS assets (
            id TEXT PRIMARY KEY,
            name TEXT,
            type TEXT,
            tag TEXT,
            details TEXT,
            created_at TEXT,
            updated_at TEXT,
            status TEXT DEFAULT 'ON',
            location TEXT,
            pressure REAL DEFAULT 0.0
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pipelines (
            source_id TEXT,
            target_id TEXT,
            created_at TEXT,
            PRIMARY KEY (source_id, target_id)
        )
    """)
    cursor.execute("PRAGMA table_info(assets)")
    cols = [r[1] for r in cursor.fetchall()]
    if "created_at" not in cols or "pressure" not in cols:
        cursor.execute("DROP TABLE IF EXISTS assets")
        cursor.execute("DROP TABLE IF EXISTS pipelines")
        conn.commit()
        conn.close()
        init_db()
        return

    cursor.execute("SELECT COUNT(*) FROM assets")
    if cursor.fetchone()[0] == 0:
        seed_default_network(cursor)
        conn.commit()
    conn.close()

init_db()

class ToggleReq(BaseModel):
    asset_id: str
    status: str

class PressureReq(BaseModel):
    drs_id: str
    pressure: float

class AssetSaveReq(BaseModel):
    id: str
    name: str
    type: str
    tag: str = "125 MM"
    details: str = ""
    status: str = "ON"
    location: str = ""
    pressure: float = 0.0
    parent_id: str = ""

class AssetDeleteReq(BaseModel):
    id: str

class PipelineReq(BaseModel):
    source_id: str
    target_id: str

class BackupRestoreReq(BaseModel):
    assets: Dict[str, Any]
    pipelines: List[List[str]]

def load_network_from_db():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM assets")
    assets = {row["id"]: dict(row) for row in cursor.fetchall()}
    cursor.execute("SELECT source_id, target_id FROM pipelines")
    pipelines = [(row["source_id"], row["target_id"]) for row in cursor.fetchall()]
    conn.close()

    G = nx.Graph()
    for node_id in assets:
        G.add_node(node_id)
    for u, v in pipelines:
        G.add_edge(u, v)

    return assets, pipelines, G

def run_scada_bfs_flow():
    assets, pipelines, G = load_network_from_db()

    active_drs = [
        k for k, v in assets.items()
        if str(v.get("type", "")).upper() == "DRS" and str(v.get("status", "ON")).upper() == "ON"
    ]
    active_drs.sort(key=lambda x: float(assets[x].get("pressure") or 0.0), reverse=True)

    gas_supply = {node: False for node in G.nodes()}
    fed_by = {}

    for drs in active_drs:
        press = float(assets[drs].get("pressure") or 4.0)
        gas_supply[drs] = True
        fed_by[drs] = {"id": drs, "pressure": press}
        q = [drs]
        visited = {drs}

        while q:
            curr = q.pop(0)
            c_type = str(assets.get(curr, {}).get("type", "")).upper()
            c_stat = str(assets.get(curr, {}).get("status", "ON")).upper()
            if ("VALVE" in c_type or c_type == "VC") and c_stat == "OFF":
                continue

            for nbr in G.neighbors(curr):
                if nbr not in visited:
                    visited.add(nbr)
                    gas_supply[nbr] = True
                    if nbr not in fed_by:
                        fed_by[nbr] = fed_by[curr]
                    q.append(nbr)

    return assets, pipelines, gas_supply, fed_by

def build_response(msg=""):
    assets, pipelines, supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "message": msg,
        "assets": assets,
        "pipelines": pipelines,
        "supply": supply,
        "fed_by": fed_by
    }

@app.get("/")
def index():
    if os.path.exists(INDEX_HTML):
        return FileResponse(INDEX_HTML, media_type="text/html")
    return {"error": "index.html not found"}

@app.get("/api/snapshot")
def get_snapshot():
    return build_response()

# AUTO-RESTORE ENDPOINT (Restores browser backup when Render restarts)
@app.post("/api/restore-backup")
def restore_backup(req: BackupRestoreReq):
    if not req.assets:
        return build_response()
    now = datetime.now().strftime("%Y-%m-%d")
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM assets")
    cursor.execute("DELETE FROM pipelines")

    a_rows = []
    for aid, a in req.assets.items():
        a_rows.append((
            str(aid),
            str(a.get("name") or aid),
            str(a.get("type") or "VALVE").upper(),
            str(a.get("tag") or "125 MM"),
            str(a.get("details") or ""),
            str(a.get("created_at") or now),
            str(a.get("updated_at") or now),
            str(a.get("status") or "ON").upper(),
            str(a.get("location") or ""),
            float(a.get("pressure") or 0.0)
        ))
    cursor.executemany("INSERT OR REPLACE INTO assets VALUES (?,?,?,?,?,?,?,?,?,?)", a_rows)

    p_rows = []
    for pair in req.pipelines:
        if len(pair) >= 2 and pair[0] and pair[1]:
            p_rows.append((str(pair[0]), str(pair[1]), now))
    cursor.executemany("INSERT OR IGNORE INTO pipelines VALUES (?,?,?)", p_rows)

    conn.commit()
    conn.close()
    return build_response("Backup data auto-restored!")

@app.post("/api/asset/save")
def save_asset(req: AssetSaveReq):
    asset_id = req.id.strip()
    if not asset_id:
        return {"success": False, "message": "Asset ID ivvali!"}

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    a_type = req.type.strip().upper()
    press = req.pressure if a_type == "DRS" else 0.0
    if a_type == "DRS" and press == 0.0:
        press = 4.0

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT created_at FROM assets WHERE id = ?", (asset_id,))
    existing = cursor.fetchone()
    created_at = existing["created_at"] if existing else now

    cursor.execute("""
        INSERT OR REPLACE INTO assets 
        (id, name, type, tag, details, created_at, updated_at, status, location, pressure)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        asset_id, req.name.strip() or asset_id, a_type,
        req.tag.strip() or "125 MM", req.details.strip(),
        created_at, now, req.status.strip().upper() or "ON",
        req.location.strip(), press
    ))

    if req.parent_id and req.parent_id.strip() and req.parent_id.strip() != asset_id:
        cursor.execute("INSERT OR IGNORE INTO pipelines VALUES (?, ?, ?)", (req.parent_id.strip(), asset_id, now))

    conn.commit()
    conn.close()
    return build_response(f"Asset '{asset_id}' saved successfully!")

@app.post("/api/asset/delete")
def delete_asset(req: AssetDeleteReq):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM assets WHERE id = ?", (req.id.strip(),))
    cursor.execute("DELETE FROM pipelines WHERE source_id = ? OR target_id = ?", (req.id.strip(), req.id.strip()))
    conn.commit()
    conn.close()
    return build_response("Asset deleted successfully!")

@app.post("/api/pipeline/add")
def add_pipeline(req: PipelineReq):
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("INSERT OR IGNORE INTO pipelines VALUES (?, ?, ?)", (req.source_id.strip(), req.target_id.strip(), now))
    conn.commit()
    conn.close()
    return build_response("Link added successfully!")

@app.post("/api/pipeline/delete")
def delete_pipeline(req: PipelineReq):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM pipelines WHERE (source_id=? AND target_id=?) OR (source_id=? AND target_id=?)", 
                   (req.source_id.strip(), req.target_id.strip(), req.target_id.strip(), req.source_id.strip()))
    conn.commit()
    conn.close()
    return build_response("Link removed successfully!")

@app.post("/api/toggle")
def toggle_asset(req: ToggleReq):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE assets SET status = ? WHERE id = ?", (req.status, req.asset_id))
    conn.commit()
    conn.close()
    return build_response()

@app.post("/api/pressure")
def update_pressure(req: PressureReq):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE assets SET pressure = ?, details = ? WHERE id = ?", (req.pressure, f"{req.pressure} Bar", req.drs_id))
    conn.commit()
    conn.close()
    return build_response()

@app.post("/api/reset")
def reset_all():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE assets SET status = 'ON'")
    conn.commit()
    conn.close()
    return build_response("All valves reset to ON!")

@app.post("/api/load-default")
def load_default():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM assets")
    cursor.execute("DELETE FROM pipelines")
    seed_default_network(cursor)
    conn.commit()
    conn.close()
    return build_response("Default network loaded!")

@app.post("/api/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        # Save a copy to master_data.xlsx on server
        with open(MASTER_EXCEL, "wb") as f:
            f.write(contents)
        conn = get_db()
        cursor = conn.cursor()
        asset_count, link_count = load_excel_bytes_to_db(contents, cursor)
        conn.commit()
        conn.close()
        return build_response(f"Excel Imported & Saved Permanently: {asset_count} Assets, {link_count} Links!")
    except Exception as e:
        return {"success": False, "message": f"Excel Read Error: {str(e)}"}

@app.get("/api/export-excel")
def export_excel():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM assets")
    ar = cursor.fetchall()
    cursor.execute("SELECT * FROM pipelines")
    lr = cursor.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws_a = wb.active
    ws_a.title = "Assets"
    ws_a.append(["Asset ID", "Asset Name", "Type", "Tag", "Details", "Created At", "Updated At", "Status", "Location"])
    for r in ar:
        ws_a.append([r["id"], r["name"], r["type"], r["tag"], r["details"], r["created_at"], r["updated_at"], r["status"], r["location"]])

    ws_l = wb.create_sheet(title="Links")
    ws_l.append(["Parent ID", "Child ID", "Created At"])
    for r in lr:
        ws_l.append([r["source_id"], r["target_id"], r["created_at"]])

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return StreamingResponse(stream, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=BGL_SCADA_Data.xlsx"})