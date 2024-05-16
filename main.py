from fastapi import *
from fastapi.responses import FileResponse
import function

app = FastAPI(title="test", version="0.0.1", description="ghksdbfrPtksrl")


@app.get("/")
async def main():
    return {"Hello": "World"}

@app.post("/exchange")
async def exchange(file: UploadFile):
    f_n = file.filename
    exl = function.excel_exchange(file)
    await file.close
    return FileResponse(exl, filename=f_n)
