from sqlalchemy.orm import Session
from app.core.models import User, Configuration, GeneratedFile

def create_user(db: Session, name: str, email: str):
    new_user = User(name=name, email=email)
    db.add(new_user)
    db.commit()
    db.refresh(new_user)
    return new_user

def get_users(db: Session):
    return db.query(User).all()

def get_templates(db: Session, is_template: bool = True):
    return db.query(GeneratedFile).filter(GeneratedFile.isTemplate == is_template).all()
