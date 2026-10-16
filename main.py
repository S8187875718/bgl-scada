import os
import io
import sqlite3
from datetime import datetime
import openpyxl
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel
import networkx as nx

app = FastAPI(title="BGL Gas SCADA Engine")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_HTML = os.path.join(BASE_DIR, "templates", "index.html")
DB_PATH = os.path.join(BASE_DIR, "scada.db")

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def seed_default_network(cursor):
    now = datetime.now().strftime("%Y-%m-%d")
    default_assets = [
        ("D01", "MS DRS", "DRS", "125 MM", "4.1 Bar", now, now, "ON", "Shamirpet", 4.1),
        ("D03", "Check Post DRS", "DRS", "125 MM", "3.6 Bar", now, now, "ON", "Check Post", 3.6),
        ("D02", "Depo DRS", "DRS", "125 MM", "3.1 Bar", now, now, "ON", "Depo", 3.1),
        ("D01-DIR-01", "Registrar, Nalsar Univ.", "CPNG", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("D01-DIR-02", "Shamirpet Spur 01", "SR", "25", "", now, now, "ON", "Shamirpet", 0.0),
        ("D01-DIR-03", "Shamirpet Spur 02", "SR", "25", "", now, now, "ON", "Shamirpet", 0.0),
        ("D01-DIR-04", "Shamirpet Spur 03", "SR", "25", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE01", "CG-D01-PE01", "VALVE", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE02", "CG-D01-PE02", "VALVE", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE02/01", "CG-D01-PE02/01", "VALVE", "90 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE02/02", "CG-D01-PE02/02", "VALVE", "90 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE02/03", "CG-D01-PE02/03", "VALVE", "63 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE03", "CG-D01-PE03", "VALVE", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE04", "CG-D01-PE04", "VALVE", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE04/01", "CG-D01-PE04/01", "VALVE", "90 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE04/02", "CG-D01-PE04/02", "VALVE", "90 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE04/03", "CG-D01-PE04/03", "VALVE", "90 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE05", "CG-D01-PE05", "VALVE", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE05/01", "CG-D01-PE05/01", "VALVE", "90 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE06", "CG-D01-PE06", "VALVE", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE07", "CG-D01-PE07", "VALVE", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-PE07/01", "CG-D01-PE07/01", "VALVE", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("31000003", "IPE Univ. Canteen", "CPNG", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-SR03", "Shamirpet SR03/04", "SR", "40", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-SR01", "Shamirpet SR01/02", "SR", "40", "", now, now, "ON", "Shamirpet", 0.0),
        ("CG-D01-SR07", "Shamirpet SR07", "SR", "65", "", now, now, "ON", "Shamirpet", 0.0),
        ("21300007", "Compack Containers", "IPNG", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("21300016", "Tunav / Aparna Alum.", "IPNG", "125 MM", "", now, now, "ON", "Shamirpet", 0.0),
        ("S03-D03-PE01", "S03-D03-PE01", "VALVE", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("S03-D03-PE02", "S03-D03-PE02", "VALVE", "63 MM", "", now, now, "ON", "Check Post", 0.0),
        ("S03-D03-PE03", "S03-D03-PE03", "VALVE", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("S03-D03-PE03/1", "S03-D03-PE03/1", "VALVE", "63 MM", "", now, now, "ON", "Check Post", 0.0),
        ("S03-D03-PE04", "S03-D03-PE04", "VALVE", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("S03-D03-PE06", "S03-D03-PE06", "VALVE", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("S03-D03-PE07", "S03-D03-PE07", "VALVE", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("21300015", "NED Energy Ltd", "IPNG", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("21300095", "New Delta Coats", "IPNG", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("21300071", "SRH Rubbers Industries", "IPNG", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("21100002", "Kayempee Foods Unit 1", "IPNG", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("21100007", "Kayempee Foods Unit 2", "IPNG", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("21300005", "Surya Masala Ltd.", "IPNG", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("21100001", "Sumo Biscuits", "IPNG", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("21100008", "Biking Foods", "IPNG", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("21300055", "Radhan Industries", "IPNG", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("21300038", "Maheswari Fasteners", "IPNG", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("21300052", "Prasad Seeds Ltd.", "IPNG", "125 MM", "", now, now, "ON", "Check Post", 0.0),
        ("S02-D02-PE01", "S02-D02-PE01", "VALVE", "125 MM", "", now, now, "ON", "Depo", 0.0),
        ("S02-D02-PE02", "S02-D02-PE02", "VALVE", "63 MM", "", now, now, "ON", "Depo", 0.0),
        ("S02-D02-PE04", "S02-D02-PE04", "VALVE", "63 MM", "", now, now, "ON", "Depo", 0.0),
        ("S02-D02-PE06", "S02-D02-PE06", "VALVE", "32 MM", "", now, now, "ON", "Depo", 0.0),
        ("S02-D02-PE07", "S02-D02-PE07", "VALVE", "90 MM", "", now, now, "ON", "Depo", 0.0),
        ("S02-D02-PE07/02", "S02-D02-PE07/02", "VALVE", "63 MM", "", now, now, "ON", "Depo", 0.0),
        ("S02-D02-PE08", "S02-D02-PE08", "VALVE", "125 MM", "", now, now, "ON", "Depo", 0.0),
        ("D02-SR10", "Medchal Residential", "SR", "25", "", now, now, "ON", "Depo", 0.0),
        ("D02-SR08", "Medchal East SR", "SR", "25", "", now, now, "ON", "Depo", 0.0),
        ("31000015", "Hare Krishna Satsang", "CPNG", "125 MM", "", now, now, "ON", "Depo", 0.0),
        ("31000002", "Hitech Modern School", "CPNG", "125 MM", "", now, now, "ON", "Depo", 0.0),
        ("31000008", "John Foundation", "CPNG", "125 MM", "", now, now, "ON", "Depo", 0.0),
        ("31000024", "Manasa Womens Hostel", "CPNG", "125 MM", "", now, now, "ON", "Depo", 0.0),
        ("31000045", "Seasons Bar & Rest.", "CPNG", "125 MM", "", now, now, "ON", "Depo", 0.0)
    ]
    cursor.executemany("INSERT OR REPLACE INTO assets VALUES (?,?,?,?,?,?,?,?,?,?)", default_assets)

    default_pipelines = [
        ("D01", "D01-DIR-01", now), ("D01", "D01-DIR-02", now), ("D01", "D01-DIR-03", now), ("D01", "D01-DIR-04", now),
        ("D01", "CG-D01-PE01", now), ("D01", "CG-D01-PE02", now), ("D01", "CG-D01-PE03", now),
        ("CG-D01-PE01", "31000003", now),
        ("CG-D01-PE02", "CG-D01-SR03", now), ("CG-D01-PE02", "CG-D01-PE02/01", now),
        ("CG-D01-PE02/01", "CG-D01-PE02/02", now), ("CG-D01-PE02/02", "CG-D01-SR01", now),
        ("CG-D01-PE02/02", "CG-D01-PE02/03", now), ("CG-D01-PE02/03", "CG-D01-SR07", now),
        ("CG-D01-PE03", "CG-D01-PE04", now), ("CG-D01-PE04", "CG-D01-PE04/01", now),
        ("CG-D01-PE04/01", "CG-D01-PE04/02", now), ("CG-D01-PE04/02", "CG-D01-PE04/03", now),
        ("CG-D01-PE04/03", "21300007", now),
        ("CG-D01-PE04/01", "CG-D01-PE05", now),
        ("CG-D01-PE05", "CG-D01-PE05/01", now), ("CG-D01-PE05/01", "21300016", now),
        ("CG-D01-PE05/01", "CG-D01-PE06", now), ("CG-D01-PE06", "CG-D01-PE07", now),
        ("CG-D01-PE07", "CG-D01-PE07/01", now),
        ("D03", "S03-D03-PE01", now), ("D03", "S03-D03-PE02", now), ("D03", "S03-D03-PE03", now),
        ("S03-D03-PE01", "21300015", now), ("S03-D03-PE01", "21300095", now), ("S03-D03-PE01", "21300071", now),
        ("S03-D03-PE01", "S03-D03-PE06", now),
        ("S03-D03-PE06", "21300055", now), ("S03-D03-PE06", "21300038", now),
        ("S03-D03-PE06", "S03-D03-PE07", now),
        ("S03-D03-PE07", "21300052", now),
        ("S03-D03-PE02", "21100002", now), ("S03-D03-PE02", "21100007", now),
        ("S03-D03-PE03", "21300005", now), ("S03-D03-PE03", "S03-D03-PE03/1", now),
        ("S03-D03-PE03/1", "21100001", now), ("S03-D03-PE03/1", "21100008", now),
        ("S03-D03-PE03/1", "S03-D03-PE04", now),
        ("D02", "S02-D02-PE01", now), ("D02", "S02-D02-PE02", now), ("D02", "S02-D02-PE06", now), ("D02", "S02-D02-PE08", now),
        ("S02-D02-PE01", "D02-SR10", now), ("S02-D02-PE01", "D02-SR08", now), ("S02-D02-PE01", "31000015", now),
        ("S02-D02-PE01", "S02-D02-PE04", now),
        ("S02-D02-PE04", "31000002", now), ("S02-D02-PE04", "31000008", now),
        ("S02-D02-PE02", "31000024", now),
        ("S02-D02-PE06", "31000045", now),
        ("S02-D02-PE08", "S02-D02-PE07", now),
        ("S02-D02-PE07", "S02-D02-PE07/02", now),
        ("CG-D01-PE05", "S03-D03-PE04", now),
        ("S03-D03-PE04", "S02-D02-PE08", now)
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
    # Check if old schema without created_at exists
    cursor.execute("PRAGMA table_info(assets)")
    cols = [r[1] for r in cursor.fetchall()]
    if "created_at" not in cols:
        cursor.execute("DROP TABLE IF EXISTS assets")
        cursor.execute("DROP TABLE IF EXISTS pipelines")
        conn.commit()
        init_db()
        return

    cursor.execute("SELECT COUNT(*) FROM assets")
    if cursor.fetchone()[0] == 0:
        seed_default_network(cursor)
        conn.commit()
    conn.close()

init_db()

# Request Models
class LoginReq(BaseModel):
    username: str
    password: str

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

@app.post("/api/login")
def login_user(req: LoginReq):
    u = req.username.strip().lower()
    p = req.password.strip()
    if u == "admin" and p == "admin123":
        return {"success": True, "role": "admin", "username": "Admin"}
    elif u == "user" and p == "user123":
        return {"success": True, "role": "user", "username": "User"}
    return {"success": False, "message": "Invalid Username leda Password!"}

@app.get("/api/snapshot")
def get_snapshot():
    return build_response()

# CREATE OR EDIT SINGLE ASSET
@app.post("/api/asset/save")
def save_asset(req: AssetSaveReq):
    asset_id = req.id.strip()
    if not asset_id:
        return {"success": False, "message": "Asset ID tappanisari!"}

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
        asset_id,
        req.name.strip() or asset_id,
        a_type,
        req.tag.strip() or "125 MM",
        req.details.strip(),
        created_at,
        now,
        req.status.strip().upper() or "ON",
        req.location.strip(),
        press
    ))

    if req.parent_id and req.parent_id.strip() and req.parent_id.strip() != asset_id:
        p_id = req.parent_id.strip()
        cursor.execute("INSERT OR IGNORE INTO pipelines VALUES (?, ?, ?)", (p_id, asset_id, now))

    conn.commit()
    conn.close()
    action = "update" if existing else "create"
    return build_response(f"Asset '{asset_id}' successfully {action} aindi!")

# DELETE SINGLE ASSET
@app.post("/api/asset/delete")
def delete_asset(req: AssetDeleteReq):
    asset_id = req.id.strip()
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM assets WHERE id = ?", (asset_id,))
    cursor.execute("DELETE FROM pipelines WHERE source_id = ? OR target_id = ?", (asset_id, asset_id))
    conn.commit()
    conn.close()
    return build_response(f"Asset '{asset_id}' mariyu dani links delete ayyayi!")

# ADD MANUAL LINK
@app.post("/api/pipeline/add")
def add_pipeline(req: PipelineReq):
    src, tgt = req.source_id.strip(), req.target_id.strip()
    if not src or not tgt or src == tgt:
        return {"success": False, "message": "Verveeru valid Assets select cheyandi!"}
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("INSERT OR IGNORE INTO pipelines VALUES (?, ?, ?)", (src, tgt, now))
    conn.commit()
    conn.close()
    return build_response(f"Link '{src}' -> '{tgt}' add aindi!")

# DELETE MANUAL LINK
@app.post("/api/pipeline/delete")
def delete_pipeline(req: PipelineReq):
    src, tgt = req.source_id.strip(), req.target_id.strip()
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM pipelines WHERE (source_id=? AND target_id=?) OR (source_id=? AND target_id=?)", (src, tgt, tgt, src))
    conn.commit()
    conn.close()
    return build_response(f"Link '{src}' <-> '{tgt}' delete aindi!")

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
    return build_response("Anni valves ON cheyabaddayi!")

@app.post("/api/load-default")
def load_default():
    conn = get_db()
    cursor