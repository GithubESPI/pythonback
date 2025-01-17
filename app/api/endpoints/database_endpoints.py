from fastapi import APIRouter, Depends
from sqlalchemy.orm import Session
from app.core.database import get_db
from app.services.database_services import create_user, get_users, get_templates

router = APIRouter()

@router.post("/users")
def add_user(name: str, email: str, db: Session = Depends(get_db)):
    user = create_user(db, name, email)
    return {"user": user}

@router.get("/users")
def list_users(db: Session = Depends(get_db)):
    users = get_users(db)
    return {"users": users}

@router.get("/templates")
def list_templates(db: Session = Depends(get_db)):
    templates = get_templates(db)
    return {"templates": templates}
