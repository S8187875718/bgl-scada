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

    # Database khali ga unte automatic ga default network ni seed chesthundi
    cursor.execute("SELECT COUNT(*) FROM assets")
    if cursor.fetchone()[0] == 0:
        seed_default_network(cursor)
        conn.commit()
    conn.close()

def seed_default_network(cursor):
    default_assets = [
        ("D01", "MS DRS", "DRS", "125 MM", "4.1 Bar", "ON", "Shamirpet", 4.1),
        ("D03", "Check Post DRS", "DRS", "125 MM", "3.6 Bar", "ON", "Check Post", 3.6),
        ("D02", "Depo DRS", "DRS", "125 MM", "3.1 Bar", "ON", "Depo", 3.1),
        ("D01-DIR-01", "Registrar, Nalsar Univ.", "CPNG", "125 MM", "", "ON", "", 0.0),
        ("D01-DIR-02", "Shamirpet Spur 01", "SR", "25", "", "ON", "", 0.0),
        ("D01-DIR-03", "Shamirpet Spur 02", "SR", "25", "", "ON", "", 0.0),
        ("D01-DIR-04", "Shamirpet Spur 03", "SR", "25", "", "ON", "", 0.0),
        ("CG-D01-PE01", "CG-D01-PE01", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("CG-D01-PE02", "CG-D01-PE02", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("CG-D01-PE02/01", "CG-D01-PE02/01", "VALVE", "90 MM", "", "ON", "", 0.0),
        ("CG-D01-PE02/02", "CG-D01-PE02/02", "VALVE", "90 MM", "", "ON", "", 0.0),
        ("CG-D01-PE02/03", "CG-D01-PE02/03", "VALVE", "63 MM", "", "ON", "", 0.0),
        ("CG-D01-PE03", "CG-D01-PE03", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("CG-D01-PE04", "CG-D01-PE04", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("CG-D01-PE04/01", "CG-D01-PE04/01", "VALVE", "90 MM", "", "ON", "", 0.0),
        ("CG-D01-PE04/02", "CG-D01-PE04/02", "VALVE", "90 MM", "", "ON", "", 0.0),
        ("CG-D01-PE04/03", "CG-D01-PE04/03", "VALVE", "90 MM", "", "ON", "", 0.0),
        ("CG-D01-PE05", "CG-D01-PE05", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("CG-D01-PE05/01", "CG-D01-PE05/01", "VALVE", "90 MM", "", "ON", "", 0.0),
        ("CG-D01-PE06", "CG-D01-PE06", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("CG-D01-PE07", "CG-D01-PE07", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("CG-D01-PE07/01", "CG-D01-PE07/01", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("31000003", "IPE Univ. Canteen", "CPNG", "125 MM", "", "ON", "", 0.0),
        ("CG-D01-SR03", "Shamirpet SR03/04", "SR", "40", "", "ON", "", 0.0),
        ("CG-D01-SR01", "Shamirpet SR01/02", "SR", "40", "", "ON", "", 0.0),
        ("CG-D01-SR07", "Shamirpet SR07", "SR", "65", "", "ON", "", 0.0),
        ("21300007", "Compack Containers", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("21300016", "Tunav / Aparna Alum.", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("S03-D03-PE01", "S03-D03-PE01", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("S03-D03-PE02", "S03-D03-PE02", "VALVE", "63 MM", "", "ON", "", 0.0),
        ("S03-D03-PE03", "S03-D03-PE03", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("S03-D03-PE03/1", "S03-D03-PE03/1", "VALVE", "63 MM", "", "ON", "", 0.0),
        ("S03-D03-PE04", "S03-D03-PE04", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("S03-D03-PE06", "S03-D03-PE06", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("S03-D03-PE07", "S03-D03-PE07", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("21300015", "NED Energy Ltd", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("21300095", "New Delta Coats", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("21300071", "SRH Rubbers Industries", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("21100002", "Kayempee Foods Unit 1", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("21100007", "Kayempee Foods Unit 2", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("21300005", "Surya Masala Ltd.", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("21100001", "Sumo Biscuits", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("21100008", "Biking Foods", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("21300055", "Radhan Industries", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("21300038", "Maheswari Fasteners", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("21300052", "Prasad Seeds Ltd.", "IPNG", "125 MM", "", "ON", "", 0.0),
        ("S02-D02-PE01", "S02-D02-PE01", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("S02-D02-PE02", "S02-D02-PE02", "VALVE", "63 MM", "", "ON", "", 0.0),
        ("S02-D02-PE04", "S02-D02-PE04", "VALVE", "63 MM", "", "ON", "", 0.0),
        ("S02-D02-PE06", "S02-D02-PE06", "VALVE", "32 MM", "", "ON", "", 0.0),
        ("S02-D02-PE07", "S02-D02-PE07", "VALVE", "90 MM", "", "ON", "", 0.0),
        ("S02-D02-PE07/02", "S02-D02-PE07/02", "VALVE", "63 MM", "", "ON", "", 0.0),
        ("S02-D02-PE08", "S02-D02-PE08", "VALVE", "125 MM", "", "ON", "", 0.0),
        ("D02-SR10", "Medchal Residential", "SR", "25", "", "ON", "", 0.0),
        ("D02-SR08", "Medchal East SR", "SR", "25", "", "ON", "", 0.0),
        ("31000015", "Hare Krishna Satsang", "CPNG", "125 MM", "", "ON", "", 0.0),
        ("31000002", "Hitech Modern School", "CPNG", "125 MM", "", "ON", "", 0.0),
        ("31000008", "John Foundation", "CPNG", "125 MM", "", "ON", "", 0.0),
        ("31000024", "Manasa Womens Hostel", "CPNG", "125 MM", "", "ON", "", 0.0),
        ("31000045", "Seasons Bar & Rest.", "CPNG", "125 MM", "", "ON", "", 0.0)
    ]
    cursor.executemany("INSERT OR REPLACE INTO assets VALUES (?,?,?,?,?,?,?,?)", default_assets)

    default_pipelines = [
        ("D01", "D01-DIR-01"), ("D01", "D01-DIR-02"), ("D01", "D01-DIR-03"), ("D01", "D01-DIR-04"),
        ("D01", "CG-D01-PE01"), ("D01", "CG-D01-PE02"), ("D01", "CG-D01-PE03"),
        ("CG-D01-PE01", "31000003"),
        ("CG-D01-PE02", "CG-D01-SR03"), ("CG-D01-PE02", "CG-D01-PE02/01"),
        ("CG-D01-PE02/01", "CG-D01-PE02/02"), ("CG-D01-PE02/02", "CG-D01-SR01"),
        ("CG-D01-PE02/02", "CG-D01-PE02/03"), ("CG-D01-PE02/03", "CG-D01-SR07"),
        ("CG-D01-PE03", "CG-D01-PE04"), ("CG-D01-PE04", "CG-D01-PE04/01"),
        ("CG-D01-PE04/01", "CG-D01-PE04/02"), ("CG-D01-PE04/02", "CG-D01-PE04/03"),
        ("CG-D01-PE04/03", "21300007"),
        ("CG-D01-PE04/01", "CG-D01-PE05"),
        ("CG-D01-PE05", "CG-D01-PE05/01"), ("CG-D01-PE05/01", "21300016"),
        ("CG-D01-PE05/01", "CG-D01-PE06"), ("CG-D01-PE06", "CG-D01-PE07"),
        ("CG-D01-PE07", "CG-D01-PE07/01"),
        ("D03", "S03-D03-PE01"), ("D03", "S03-D03-PE02"), ("D03", "S03-D03-PE03"),
        ("S03-D03-PE01", "21300015"), ("S03-D03-PE01", "21300095"), ("S03-D03-PE01", "21300071"),
        ("S03-D03-PE01", "S03-D03-PE06"),
        ("S03-D03-PE06", "21300055"), ("S03-D03-PE06", "21300038"),
        ("S03-D03-PE06", "S03-D03-PE07"),
        ("S03-D03-PE07", "21300052"),
        ("S03-D03-PE02", "21100002"), ("S03-D03-PE02", "21100007"),
        ("S03-D03-PE03", "21300005"), ("S03-D03-PE03", "S03-D03-PE03/1"),
        ("S03-D03-PE03/1", "21100001"), ("S03-D03-PE03/1", "21100008"),
        ("S03-D03-PE03/1", "S03-D03-PE04"),
        ("D02", "S02-D02-PE01"), ("D02", "S02-D02-PE02"), ("D02", "S02-D02-PE06"), ("D02", "S02-D02-PE08"),
        ("S02-D02-PE01", "D02-SR10"), ("S02-D02-PE01", "D02-SR08"), ("S02-D02-PE01", "31000015"),
        ("S02-D02-PE01", "S02-D02-PE04"),
        ("S02-D02-PE04", "31000002"), ("S02-D02-PE04", "31000008"),
        ("S02-D02-PE02", "31000024"),
        ("S02-D02-PE06", "31000045"),
        ("S02-D02-PE08", "S02-D02-PE07"),
        ("S02-D02-PE07", "S02-D02-PE07/02"),
        ("CG-D01-PE05", "S03-D03-PE04"),
        ("S03-D03-PE04", "S02-D02-PE08")
    ]
    cursor.executemany("INSERT OR REPLACE INTO pipelines VALUES (?,?)", default_pipelines)

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
    @app.post("/api/load-default")
def load_default_network_endpoint():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DROP TABLE IF EXISTS assets")
    cursor.execute("DROP TABLE IF EXISTS pipelines")
    cursor.execute("""
        CREATE TABLE assets (
            id TEXT PRIMARY KEY, name TEXT, type TEXT, tag TEXT,
            details TEXT, status TEXT DEFAULT 'ON', location TEXT, pressure REAL DEFAULT 0.0
        )
    """)
    cursor.execute("""
        CREATE TABLE pipelines (
            source_id TEXT, target_id TEXT, PRIMARY KEY (source_id, target_id)
        )
    """)
    seed_default_network(cursor)
    conn.commit()
    conn.close()

    assets, pipelines, supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "message": "Default BGL Network successfully loaded!",
        "assets": assets,
        "pipelines": pipelines,
        "supply": supply,
        "fed_by": fed_by
    }