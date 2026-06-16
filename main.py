from pathlib import Path
from dotenv import load_dotenv
from starlette.responses import HTMLResponse, JSONResponse
from starlette.routing import Route
from starlette.applications import Starlette
import uvicorn

load_dotenv()

WORKSPACE = Path(__file__).parent
HTML_FILE = WORKSPACE / "TaskPlanner.html"


async def homepage(request):
    return HTMLResponse(HTML_FILE.read_text(encoding="utf-8"))


async def health(request):
    return JSONResponse({"status": "Healthy"})


async def invocations(request):
    return JSONResponse({"status": "ok", "message": "Use GET / to open the Task Planner."})


app = Starlette(routes=[
    Route("/", homepage, methods=["GET"]),
    Route("/health", health, methods=["GET"]),
    Route("/invocations", invocations, methods=["POST"]),
])

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8080)
