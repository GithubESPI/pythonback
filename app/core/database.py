from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from app.core.config import settings

# Créer une connexion à la base de données PostgreSQL de Prisma
engine = create_engine(settings.DATABASE_URL, connect_args={"sslmode": "require"})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

# Dépendance pour obtenir une session de base de données
def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
