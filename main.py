import io
from pprint import pprint

from fastapi import *
from fastapi.responses import StreamingResponse
from starlette.middleware.cors import CORSMiddleware
from starlette.responses import FileResponse

import function

app = FastAPI(title="test", version="0.0.1", description="ghksdbfrPtksrl", debug=True)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
async def main():
    return FileResponse("static/index.html")


@app.get("/style.css")
async def style():
    return FileResponse("static/style.css")


@app.post("/exchange")
async def exchange(file: UploadFile):
    file_name = file.filename
    file_data = await file.read()
    fdata = io.BytesIO(file_data)
    await file.close()
    fdata.seek(0)
    data = function.excel_exchange(fdata)
    response = StreamingResponse(data, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                 headers={"Content-Disposition": f"attachment; filename={file_name}"})
    return response
