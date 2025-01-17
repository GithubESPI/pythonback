from sqlalchemy.orm import Session
from sqlalchemy import text

def get_users(db: Session):
    # Exemple de requête SQL brute pour une table Prisma appelée `User`
    result = db.execute(text("SELECT * FROM User"))
    return result.fetchall()
