import os
import io
import sqlite3
import openpyxl
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from pydantic import BaseModel
import networkx as nx

app = FastAPI(title="BGL Gas SCADA Database Engine")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_HTML = os.path.join(BASE_DIR, "templates", "index.html")
DB_PATH = os.path.join(BASE_DIR, "scada.db")

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

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
            status TEXT DEFAULT 'ON',
            location TEXT,
            pressure REAL DEFAULT 0.0
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pipelines (
            source_id TEXT,
            target_id TEXT,
            PRIMARY KEY (source_id, target_id)
        )
    """)
    conn.commit()
    conn.close()

init_db()

def load_network_from_db():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM assets")
    assets = {row["id"]: dict(row) for row in cursor.fetchall()}
    cursor.execute("SELECT source_id, target_id FROM pipelines")
    pipelines = [(row["source_id"], row["target_id"]) for row in cursor.fetchall()]
    conn.close()

    G = nx.Graph()
    for u, v in pipelines:
        G.add_edge(u, v)

    return assets, pipelines, G

# Request Models
class ToggleReq(BaseModel):
    asset_id: str
    status: str

class PressureReq(BaseModel):
    drs_id: str
    pressure: float

class AssetCreateReq(BaseModel):
    id: str
    name: str
    type: str
    tag: str = "125 MM"
    pressure: float = 0.0

class PipelineCreateReq(BaseModel):
    source_id: str
    target_id: str

def run_scada_bfs_flow():
    assets, pipelines, G = load_network_from_db()

    active_drs = [
        k for k, v in assets.items() 
        if v.get("type") == "DRS" and v.get("status") == "ON"
    ]
    active_drs.sort(key=lambda x: assets[x].get("pressure", 0.0), reverse=True)

    gas_supply = {node: False for node in G.nodes()}
    fed_by = {}

    for drs in active_drs:
        press = assets[drs].get("pressure", 4.0)
        gas_supply[drs] = True
        fed_by[drs] = {"id": drs, "pressure": press}
        q = [drs]
        visited = {drs}

        while q:
            curr = q.pop(0)

            # ONLY CLOSED VALVES STOP FLOW
            c_type = str(assets.get(curr, {}).get("type", "")).upper()
            if ("VALVE" in c_type or c_type == "VC") and assets.get(curr, {}).get("status") == "OFF":
                continue

            for nbr in G.neighbors(curr):
                if nbr not in visited:
                    visited.add(nbr)
                    gas_supply[nbr] = True
                    if nbr not in fed_by:
                        fed_by[nbr] = fed_by[curr]
                    q.append(nbr)

    return assets, pipelines, gas_supply, fed_by

@app.get("/")
def index():
    if os.path.exists(INDEX_HTML):
        return FileResponse(INDEX_HTML, media_type="text/html")
    return {"error": "index.html not found"}

@app.get("/api/snapshot")
def get_snapshot():
    assets, pipelines, supply, fed_by = run_scada_bfs_flow()
    return {
        "assets": assets,
        "pipelines": pipelines,
        "supply": supply,
        "fed_by": fed_by
    }

# 1. EXCEL IMPORT
@app.post("/api/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    filename = file.filename.lower()
    if not (filename.endswith(".xlsx") or filename.endswith(".xlsm")):
        return {"success": False, "message": "Excel (.xlsx) ఫైల్‌ను అప్‌లోడ్ చేయండి!"}

    contents = await file.read()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=True)
    except Exception as e:
        return {"success": False, "message": f"Excel లోపం: {str(e)}"}

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("DROP TABLE IF EXISTS assets")
    cursor.execute("DROP TABLE IF EXISTS pipelines")
    cursor.execute("""
        CREATE TABLE assets (
            id TEXT PRIMARY KEY,
            name TEXT,
            type TEXT,
            tag TEXT,
            details TEXT,
            status TEXT DEFAULT 'ON',
            location TEXT,
            pressure REAL DEFAULT 0.0
        )
    """)
    cursor.execute("""
        CREATE TABLE pipelines (
            source_id TEXT,
            target_id TEXT,
            PRIMARY KEY (source_id, target_id)
        )
    """)

    # Parse Assets
    assets_sheet = None
    for s_name in wb.sheetnames:
        if s_name.strip().lower() == "assets":
            assets_sheet = wb[s_name]
            break

    asset_count = 0
    if assets_sheet:
        headers = [str(cell.value or "").strip() for cell in assets_sheet[1]]
        col_map = {name: idx for idx, name in enumerate(headers) if name}

        assets_data = []
        for row in assets_sheet.iter_rows(min_row=2, values_only=True):
            asset_id = str(row[col_map.get("Asset ID", 0)] or "").strip()
            if not asset_id:
                continue

            name = str(row[col_map.get("Asset Name", 1)] or asset_id).strip()
            a_type = str(row[col_map.get("Type", 2)] or "VALVE").strip().upper()
            tag = str(row[col_map.get("Tag", 3)] or "125 MM").strip()
            details = str(row[col_map.get("Details", 4)] or "").strip()
            status = str(row[col_map.get("Status", 7)] or "ON").strip().upper()
            location = str(row[col_map.get("Location", 8)] or "").strip()

            pressure = 4.0 if a_type == "DRS" else 0.0
            if a_type == "DRS" and "bar" in details.lower():
                try:
                    pressure = float("".join(c for c in details if c.isdigit() or c == "."))
                except:
                    pass

            assets_data.append((asset_id, name, a_type, tag, details, status, location, pressure))

        cursor.executemany("INSERT OR REPLACE INTO assets VALUES (?,?,?,?,?,?,?,?)", assets_data)
        asset_count = len(assets_data)

    # Parse Links
    links_sheet = None
    for s_name in wb.sheetnames:
        if s_name.strip().lower() in ["links", "link", "sheet2"]:
            links_sheet = wb[s_name]
            break

    link_count = 0
    if links_sheet:
        headers = [str(cell.value or "").strip() for cell in links_sheet[1]]
        col_map = {name: idx for idx, name in enumerate(headers) if name}

        parent_idx = col_map.get("Parent ID", 0)
        child_idx = col_map.get("Child ID", 1)

        links_data = set()
        for row in links_sheet.iter_rows(min_row=2, values_only=True):
            src = str(row[parent_idx] or "").strip()
            tgt = str(row[child_idx] or "").strip()
            if src and tgt and src != tgt:
                links_data.add((src, tgt))

        cursor.executemany("INSERT OR IGNORE INTO pipelines VALUES (?,?)", list(links_data))
        link_count = len(links_data)

    conn.commit()
    conn.close()

    assets, pipelines, supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "message": f"Excel ఇంపోర్ట్ పూర్తయింది! ({asset_count} Assets, {link_count} Links)",
        "assets": assets,
        "pipelines": pipelines,
        "supply": supply,
        "fed_by": fed_by
    }

# 2. MANUAL ADD ASSET
@app.post("/api/asset/add")
def add_new_asset(req: AssetCreateReq):
    asset_id = req.id.strip()
    if not asset_id:
        return {"success": False, "message": "Asset ID తప్పనిసరిగా ఇవ్వాలి!"}

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM assets WHERE id = ?", (asset_id,))
    if cursor.fetchone():
        conn.close()
        return {"success": False, "message": f"Asset '{asset_id}' ఇప్పటికే ఉంది!"}

    press = req.pressure if req.type.upper() == "DRS" else 0.0
    cursor.execute(
        "INSERT INTO assets (id, name, type, tag, details, status, location, pressure) VALUES (?, ?, ?, ?, '', 'ON', '', ?)",
        (asset_id, req.name.strip() or asset_id, req.type.upper(), req.tag.strip() or "125 MM", press)
    )
    conn.commit()
    conn.close()

    assets, pipelines, supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "message": f"Asset '{asset_id}' విజయవంతంగా యాడ్ అయింది!",
        "assets": assets,
        "pipelines": pipelines,
        "supply": supply,
        "fed_by": fed_by
    }

# 3. MANUAL LINK PIPELINE
@app.post("/api/pipeline/add")
def add_pipeline_link(req: PipelineCreateReq):
    src = req.source_id.strip()
    tgt = req.target_id.strip()

    if not src or not tgt or src == tgt:
        return {"success": False, "message": "సరైన వేర్వేరు రెండు Assets ని ఎంచుకోండి!"}

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT source_id FROM pipelines WHERE (source_id=? AND target_id=?) OR (source_id=? AND target_id=?)", (src, tgt, tgt, src))
    if cursor.fetchone():
        conn.close()
        return {"success": False, "message": "ఈ రెండు Assets మధ్య లింక్ ఇప్పటికే ఉంది!"}

    cursor.execute("INSERT INTO pipelines (source_id, target_id) VALUES (?, ?)", (src, tgt))
    conn.commit()
    conn.close()

    assets, pipelines, supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "message": f"లింక్ '{src}' ⟷ '{tgt}' విజయవంతంగా కనెక్ట్ అయింది!",
        "assets": assets,
        "pipelines": pipelines,
        "supply": supply,
        "fed_by": fed_by
    }

@app.post("/api/toggle")
def toggle_asset(req: ToggleReq):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE assets SET status = ? WHERE id = ?", (req.status, req.asset_id))
    conn.commit()
    conn.close()

    assets, pipelines, supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "assets": assets,
        "pipelines": pipelines,
        "supply": supply,
        "fed_by": fed_by
    }

@app.post("/api/pressure")
def update_pressure(req: PressureReq):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE assets SET pressure = ? WHERE id = ?", (req.pressure, req.drs_id))
    conn.commit()
    conn.close()

    assets, pipelines, supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "assets": assets,
        "pipelines": pipelines,
        "supply": supply,
        "fed_by": fed_by
    }

@app.post("/api/reset")
def reset_all():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE assets SET status = 'ON'")
    conn.commit()
    conn.close()

    assets, pipelines, supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "assets": assets,
        "pipelines": pipelines,
        "supply": supply,
        "fed_by": fed_by
    }