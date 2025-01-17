import logging
from fastapi import FastAPI
from app.api.endpoints import uploads
from app.api.endpoints.ypareo_endpoints import router as ypareo_router

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()  # Affiche les logs dans la console
    ]
)

app = FastAPI()

# Include routers
app.include_router(uploads.router, prefix="", tags=["uploads"])
app.include_router(ypareo_router, prefix="/ypareo", tags=["ypareo"])

@app.get("/")
def read_root():
    return {"message": "Bienvenue dans l'application génération des bulletins"}

@app.post("/process-template")
def process_template(output_path: str):
    try:
        return {"message": "Template processed successfully", "output_path": output_path}
    except Exception as e:
        logging.error(f"Error processing template: {str(e)}")
        return {"error": str(e)}
