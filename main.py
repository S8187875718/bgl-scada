import os
from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel
import networkx as nx

app = FastAPI(title="BGL Gas SCADA Engine")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_HTML = os.path.join(BASE_DIR, "templates", "index.html")

# In-Memory Network Assets Registry
ASSETS = {
    # DRS Stations
    "D01": {"id": "D01", "type": "DRS", "name": "MS DRS", "pressure": 4.1, "status": "ON"},
    "D03": {"id": "D03", "type": "DRS", "name": "Check Post DRS", "pressure": 3.6, "status": "ON"},
    "D02": {"id": "D02", "type": "DRS", "name": "Depo DRS", "pressure": 3.1, "status": "ON"},
    
    # D01 Direct Spurs
    "D01-DIR-01": {"id": "D01-DIR-01", "type": "CPNG", "name": "Registrar, Nalsar Univ.", "tag": "125 MM", "status": "ON"},
    "D01-DIR-02": {"id": "D01-DIR-02", "type": "SR", "name": "Shamirpet Spur 01", "tag": "25", "status": "ON"},
    "D01-DIR-03": {"id": "D01-DIR-03", "type": "SR", "name": "Shamirpet Spur 02", "tag": "25", "status": "ON"},
    "D01-DIR-04": {"id": "D01-DIR-04", "type": "SR", "name": "Shamirpet Spur 03", "tag": "25", "status": "ON"},

    # D01 Line Valves
    "CG-D01-PE01": {"id": "CG-D01-PE01", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "CG-D01-PE02": {"id": "CG-D01-PE02", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "CG-D01-PE02/01": {"id": "CG-D01-PE02/01", "type": "VALVE", "tag": "90 MM", "status": "ON"},
    "CG-D01-PE02/02": {"id": "CG-D01-PE02/02", "type": "VALVE", "tag": "90 MM", "status": "ON"},
    "CG-D01-PE02/03": {"id": "CG-D01-PE02/03", "type": "VALVE", "tag": "63 MM", "status": "ON"},
    "CG-D01-PE03": {"id": "CG-D01-PE03", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "CG-D01-PE04": {"id": "CG-D01-PE04", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "CG-D01-PE04/01": {"id": "CG-D01-PE04/01", "type": "VALVE", "tag": "90 MM", "status": "ON"},
    "CG-D01-PE04/02": {"id": "CG-D01-PE04/02", "type": "VALVE", "tag": "90 MM", "status": "ON"},
    "CG-D01-PE04/03": {"id": "CG-D01-PE04/03", "type": "VALVE", "tag": "90 MM", "status": "ON"},
    "CG-D01-PE05": {"id": "CG-D01-PE05", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "CG-D01-PE05/01": {"id": "CG-D01-PE05/01", "type": "VALVE", "tag": "90 MM", "status": "ON"},
    "CG-D01-PE06": {"id": "CG-D01-PE06", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "CG-D01-PE07": {"id": "CG-D01-PE07", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "CG-D01-PE07/01": {"id": "CG-D01-PE07/01", "type": "VALVE", "tag": "125 MM", "status": "ON"},

    # D01 Consumers
    "31000003": {"id": "31000003", "type": "CPNG", "name": "IPE Univ. Canteen", "tag": "125 MM", "status": "ON"},
    "CG-D01-SR03": {"id": "CG-D01-SR03", "type": "SR", "name": "Shamirpet SR03/04", "tag": "40", "status": "ON"},
    "CG-D01-SR01": {"id": "CG-D01-SR01", "type": "SR", "name": "Shamirpet SR01/02", "tag": "40", "status": "ON"},
    "CG-D01-SR07": {"id": "CG-D01-SR07", "type": "SR", "name": "Shamirpet SR07", "tag": "65", "status": "ON"},
    "21300007": {"id": "21300007", "type": "IPNG", "name": "Compack Containers", "tag": "125 MM", "status": "ON"},
    "21300016": {"id": "21300016", "type": "IPNG", "name": "Tunav / Aparna Alum.", "tag": "125 MM", "status": "ON"},

    # D03 Valves & Consumers
    "S03-D03-PE01": {"id": "S03-D03-PE01", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "S03-D03-PE02": {"id": "S03-D03-PE02", "type": "VALVE", "tag": "63 MM", "status": "ON"},
    "S03-D03-PE03": {"id": "S03-D03-PE03", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "S03-D03-PE03/1": {"id": "S03-D03-PE03/1", "type": "VALVE", "tag": "63 MM", "status": "ON"},
    "S03-D03-PE04": {"id": "S03-D03-PE04", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "S03-D03-PE06": {"id": "S03-D03-PE06", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "S03-D03-PE07": {"id": "S03-D03-PE07", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "21300015": {"id": "21300015", "type": "IPNG", "name": "NED Energy Ltd", "tag": "125 MM", "status": "ON"},
    "21300095": {"id": "21300095", "type": "IPNG", "name": "New Delta Coats", "tag": "125 MM", "status": "ON"},
    "21300071": {"id": "21300071", "type": "IPNG", "name": "SRH Rubbers Industries", "tag": "125 MM", "status": "ON"},
    "21100002": {"id": "21100002", "type": "IPNG", "name": "Kayempee Foods Unit 1", "tag": "125 MM", "status": "ON"},
    "21100007": {"id": "21100007", "type": "IPNG", "name": "Kayempee Foods Unit 2", "tag": "125 MM", "status": "ON"},
    "21300005": {"id": "21300005", "type": "IPNG", "name": "Surya Masala Ltd.", "tag": "125 MM", "status": "ON"},
    "21100001": {"id": "21100001", "type": "IPNG", "name": "Sumo Biscuits", "tag": "125 MM", "status": "ON"},
    "21100008": {"id": "21100008", "type": "IPNG", "name": "Biking Foods", "tag": "125 MM", "status": "ON"},
    "21300055": {"id": "21300055", "type": "IPNG", "name": "Radhan Industries", "tag": "125 MM", "status": "ON"},
    "21300038": {"id": "21300038", "type": "IPNG", "name": "Maheswari Fasteners", "tag": "125 MM", "status": "ON"},
    "21300052": {"id": "21300052", "type": "IPNG", "name": "Prasad Seeds Ltd.", "tag": "125 MM", "status": "ON"},

    # D02 Valves & Consumers
    "S02-D02-PE01": {"id": "S02-D02-PE01", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "S02-D02-PE02": {"id": "S02-D02-PE02", "type": "VALVE", "tag": "63 MM", "status": "ON"},
    "S02-D02-PE04": {"id": "S02-D02-PE04", "type": "VALVE", "tag": "63 MM", "status": "ON"},
    "S02-D02-PE06": {"id": "S02-D02-PE06", "type": "VALVE", "tag": "32 MM", "status": "ON"},
    "S02-D02-PE07": {"id": "S02-D02-PE07", "type": "VALVE", "tag": "90 MM", "status": "ON"},
    "S02-D02-PE07/02": {"id": "S02-D02-PE07/02", "type": "VALVE", "tag": "63 MM", "status": "ON"},
    "S02-D02-PE08": {"id": "S02-D02-PE08", "type": "VALVE", "tag": "125 MM", "status": "ON"},
    "D02-SR10": {"id": "D02-SR10", "type": "SR", "name": "Medchal Residential", "tag": "25", "status": "ON"},
    "D02-SR08": {"id": "D02-SR08", "type": "SR", "name": "Medchal East SR", "tag": "25", "status": "ON"},
    "31000015": {"id": "31000015", "type": "CPNG", "name": "Hare Krishna Satsang", "tag": "125 MM", "status": "ON"},
    "31000002": {"id": "31000002", "type": "CPNG", "name": "Hitech Modern School", "tag": "125 MM", "status": "ON"},
    "31000008": {"id": "31000008", "type": "CPNG", "name": "John Foundation", "tag": "125 MM", "status": "ON"},
    "31000024": {"id": "31000024", "type": "CPNG", "name": "Manasa Womens Hostel", "tag": "125 MM", "status": "ON"},
    "31000045": {"id": "31000045", "type": "CPNG", "name": "Seasons Bar & Rest.", "tag": "125 MM", "status": "ON"}
}

# Network Topology Pipelines
PIPELINES = [
    # D01 Direct Spurs
    ("D01", "D01-DIR-01"), ("D01", "D01-DIR-02"), ("D01", "D01-DIR-03"), ("D01", "D01-DIR-04"),

    # D01 Distribution
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

    # D03 Distribution
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

    # D02 Distribution
    ("D02", "S02-D02-PE01"), ("D02", "S02-D02-PE02"), ("D02", "S02-D02-PE06"), ("D02", "S02-D02-PE08"),
    ("S02-D02-PE01", "D02-SR10"), ("S02-D02-PE01", "D02-SR08"), ("S02-D02-PE01", "31000015"),
    ("S02-D02-PE01", "S02-D02-PE04"),
    ("S02-D02-PE04", "31000002"), ("S02-D02-PE04", "31000008"),
    ("S02-D02-PE02", "31000024"),
    ("S02-D02-PE06", "31000045"),
    ("S02-D02-PE08", "S02-D02-PE07"),
    ("S02-D02-PE07", "S02-D02-PE07/02"),

    # Inter-DRS Ring Loops
    ("CG-D01-PE05", "S03-D03-PE04"),
    ("S03-D03-PE04", "S02-D02-PE08")
]

G = nx.Graph()
for u, v in PIPELINES:
    G.add_edge(u, v)

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
    active_drs = [
        k for k, v in ASSETS.items() 
        if v.get("type") == "DRS" and v.get("status") == "ON"
    ]
    active_drs.sort(key=lambda x: ASSETS[x].get("pressure", 0.0), reverse=True)

    gas_supply = {node: False for node in G.nodes()}
    fed_by = {}

    for drs in active_drs:
        press = ASSETS[drs].get("pressure", 4.0)
        gas_supply[drs] = True
        fed_by[drs] = {"id": drs, "pressure": press}
        q = [drs]
        visited = {drs}

        while q:
            curr = q.pop(0)

            # ONLY CLOSED VALVES BLOCK FLOW
            if ASSETS.get(curr, {}).get("type") == "VALVE" and ASSETS.get(curr, {}).get("status") == "OFF":
                continue

            for nbr in G.neighbors(curr):
                if nbr not in visited:
                    visited.add(nbr)
                    gas_supply[nbr] = True
                    if nbr not in fed_by:
                        fed_by[nbr] = fed_by[curr]
                    q.append(nbr)

    return gas_supply, fed_by

@app.get("/")
def index():
    if os.path.exists(INDEX_HTML):
        return FileResponse(INDEX_HTML, media_type="text/html")
    return {"error": "index.html not found"}

@app.get("/api/snapshot")
def get_snapshot():
    supply, fed_by = run_scada_bfs_flow()
    return {
        "assets": ASSETS,
        "pipelines": PIPELINES,
        "supply": supply,
        "fed_by": fed_by
    }

@app.post("/api/toggle")
def toggle_asset(req: ToggleReq):
    if req.asset_id in ASSETS:
        ASSETS[req.asset_id]["status"] = req.status
    supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "assets": ASSETS,
        "pipelines": PIPELINES,
        "supply": supply,
        "fed_by": fed_by
    }

@app.post("/api/pressure")
def update_pressure(req: PressureReq):
    if req.drs_id in ASSETS:
        ASSETS[req.drs_id]["pressure"] = req.pressure
    supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "assets": ASSETS,
        "pipelines": PIPELINES,
        "supply": supply,
        "fed_by": fed_by
    }

# STEP 1: ADD ASSET ENDPOINT
@app.post("/api/asset/add")
def add_new_asset(req: AssetCreateReq):
    asset_id = req.id.strip()
    if not asset_id:
        return {"success": False, "message": "Asset ID ivvali!"}
    
    if asset_id in ASSETS:
        return {"success": False, "message": f"Asset '{asset_id}' already undi!"}

    ASSETS[asset_id] = {
        "id": asset_id,
        "name": req.name.strip() or asset_id,
        "type": req.type.upper(),
        "tag": req.tag.strip() or "125 MM",
        "pressure": req.pressure if req.type.upper() == "DRS" else 0.0,
        "status": "ON"
    }
    
    if not G.has_node(asset_id):
        G.add_node(asset_id)

    supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "message": f"Asset '{asset_id}' successfully create aindi!",
        "assets": ASSETS,
        "pipelines": PIPELINES,
        "supply": supply,
        "fed_by": fed_by
    }

# STEP 2: LINK ASSETS ENDPOINT
@app.post("/api/pipeline/add")
def add_pipeline_link(req: PipelineCreateReq):
    src = req.source_id.strip()
    tgt = req.target_id.strip()

    if not src or not tgt:
        return {"success": False, "message": "Source mariyu Target rendu select cheyali!"}

    if src == tgt:
        return {"success": False, "message": "Same asset ki connect cheyalemu!"}

    if src not in ASSETS or tgt not in ASSETS:
        return {"success": False, "message": "Edo oka asset system lo ledu!"}

    # Check already existing edge
    if (src, tgt) in PIPELINES or (tgt, src) in PIPELINES:
        return {"success": False, "message": "Ee rendu assets madhya link already undi!"}

    PIPELINES.append((src, tgt))
    G.add_edge(src, tgt)

    supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "message": f"Link '{src}' ⟷ '{tgt}' successfully connect aindi!",
        "assets": ASSETS,
        "pipelines": PIPELINES,
        "supply": supply,
        "fed_by": fed_by
    }

@app.post("/api/reset")
def reset_all():
    for k in ASSETS:
        ASSETS[k]["status"] = "ON"
    supply, fed_by = run_scada_bfs_flow()
    return {
        "success": True,
        "assets": ASSETS,
        "pipelines": PIPELINES,
        "supply": supply,
        "fed_by": fed_by
    }