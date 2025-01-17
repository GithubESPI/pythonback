from fastapi import APIRouter, HTTPException
from app.services.ypareo_service import YpareoService

router = APIRouter()

@router.get("/periodes/2023_2024")
async def get_periode_2023_2024():
    try:
        return YpareoService.get_periode_2023_2024()
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/apprenants/frequentes")
async def get_frequentes():
    try:
        return YpareoService.get_frequentes()
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/apprenants")
async def get_apprenants():
    try:
        return YpareoService.get_apprenants()
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/groupes")
async def get_groupes():
    try:
        return YpareoService.get_groupes()
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/absences")
async def get_absences():
    try:
        return YpareoService.get_absences()
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
