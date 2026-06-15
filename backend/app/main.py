from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from .database import init_db
from .config import Config
from .api import setup, expenses, analytics, categorize, budgets, alerts, advanced_analytics, exports, saved_views

app = FastAPI(
    title="Expense Categorizer API",
    description="API para gestionar gastos con categorización automática",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

init_db()

# Include all routers
app.include_router(setup.router)
app.include_router(expenses.router)
app.include_router(analytics.router)
app.include_router(categorize.router)
app.include_router(budgets.router)
app.include_router(alerts.router)
app.include_router(advanced_analytics.router)
app.include_router(exports.router)
app.include_router(saved_views.router)

@app.get("/health")
def health_check():
    return {"status": "ok"}

@app.get("/api/config/status")
def config_status():
    """Devuelve estado de la configuración"""
    exists = Config.exists()
    config = Config.load() if exists else {}
    
    return {
        "configured": exists,
        "database_path": config.get("database_path"),
        "records_count": config.get("records_count", 0),
        "categories_count": config.get("categories_count", 0)
    }

@app.get("/")
def root():
    return {
        "message": "Expense Categorizer API v2.0",
        "docs": "/docs",
        "status": config_status()
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
