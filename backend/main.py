from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from contextlib import asynccontextmanager
from database import connect_to_mongo, close_mongo_connection
from routers import upload, voc, statistics, memos, chipset, export, qdata


@asynccontextmanager
async def lifespan(app: FastAPI):
    await connect_to_mongo()
    yield
    await close_mongo_connection()


app = FastAPI(
    title="VOC 관리 시스템",
    description="Vue + FastAPI + MongoDB 기반 VOC 관리 시스템",
    version="0.2.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(upload.router)
app.include_router(voc.router)
app.include_router(statistics.router)
app.include_router(memos.router)
app.include_router(chipset.router)
app.include_router(export.router)
app.include_router(qdata.router)


@app.get("/health")
async def health_check():
    return {"status": "ok", "service": "VOC Management API"}
