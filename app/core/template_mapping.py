"""
Module contenant le mapping des templates et les fonctions associées
"""
from prisma import Prisma

TEMPLATE_MAPPING = {
    # BG-TP-S1
    "B-BG1 TP - TP Semestre 1": "BG-TP-S1.xlsx",
    "L-BG1 TP - TP Semestre 1": "BG-TP-S1.xlsx",
    "M-BG1 TP - TP Semestre 1": "BG-TP-S1.xlsx",
    "N-BG1 TP - TP Semestre 1": "BG-TP-S1.xlsx",
    "P-BG1 TP 1 - TP Semestre 1": "BG-TP-S1.xlsx",
    "P-BG1 TP 2 Rentrée décalée": "BG-TP-S1.xlsx",
    
    # BG-TP-S2
    "B-BG1 TP - TP Semestre 2": "BG-TP-S2.xlsx",
    "L-BG1 TP - TP Semestre 2": "BG-TP-S2.xlsx",
    "M-BG1 TP - TP Semestre 2": "BG-TP-S2.xlsx",
    "N-BG1 TP - TP Semestre 2": "BG-TP-S2.xlsx",
    "P-BG1 TP 1 - TP Semestre 2": "BG-TP-S2.xlsx",
    
    # BG-TP-S3
    "P-BG2 TP - TP Semestre 1": "BG-TP-S3.xlsx",
    "L-BG2 TP - TP Semestre 1": "BG-TP-S3.xlsx",
    "M-BG2 TP - TP Semestre 1": "BG-TP-S3.xlsx",
    "N-BG2 TP - TP Semestre 1": "BG-TP-S3.xlsx",
    
    # BG-TP-S4
    "P-BG2 TP - TP Semestre 2": "BG-TP-S4.xlsx",
    "L-BG2 TP - TP Semestre 2": "BG-TP-S4.xlsx",
    "M-BG2 TP - TP Semestre 2": "BG-TP-S4.xlsx",
    "N-BG2 TP - TP Semestre 2": "BG-TP-S4.xlsx",
    
    # BG-TP-S5
    "P-BG3 TP 1 - TP Semestre 1": "BG-TP-S5.xlsx",
    "P-BG3 TP 2 Rentrée décalée": "BG-TP-S5.xlsx",
    "P-BG3 TP Section Internationale": "BG-TP-S5.xlsx",
    "N-BG3 TP - TP Semestre 1": "BG-TP-S5.xlsx",
    "L-BG3 TP - TP Semestre 1": "BG-TP-S5.xlsx",
    
    # BG-TP-S6
    "P-BG3 TP 1 - TP Semestre 2": "BG-TP-S6.xlsx",
    "N-BG3 TP - TP Semestre 2": "BG-TP-S6.xlsx",
    "L-BG3 TP - TP Semestre 2": "BG-TP-S6.xlsx",
    
    # BG-ALT-S1
    "L-BG1 ALT 1 - ALT Semestre 1 - 1ère année": "BG-ALT-S1.xlsx",
    "L-BG1 ALT 2 - ALT Semestre 1 - 1ère année": "BG-ALT-S1.xlsx",
    "LI-BG1 ALT - ALT Semestre 1 - 1ère année": "BG-ALT-S1.xlsx",
    "M-BG1 ALT - ALT Semestre 1 - 1ère année": "BG-ALT-S1.xlsx",
    "MP-BG1 ALT - ALT Semestre 1 - 1ère année": "BG-ALT-S1.xlsx",
    "N-BG1 ALT - ALT Semestre 1 - 1ère année": "BG-ALT-S1.xlsx",
    "P-BG1 ALT 1 - ALT Semestre 1 - 1ère année": "BG-ALT-S1.xlsx",
    "P-BG1 ALT 2 - ALT Semestre 1 - 1ère année": "BG-ALT-S1.xlsx",
    "P-BG1 ALT 3 - ALT Semestre 1 - 1ère année": "BG-ALT-S1.xlsx",
    "L-BG1 ALT 2 - ALT Semestre 1": "BG-ALT-S1.xlsx",
    
    # BG-ALT-S2
    "L-BG1 ALT 1 - ALT Semestre 2 - 1ère année": "BG-ALT-S2.xlsx",
    
    # BG-ALT-S3
    "L-BG2 ALT - ALT Semestre 1": "BG-ALT-S3.xlsx",
    
    # BG-ALT-S4
    "B-BG2 ALT - ALT Semestre 2": "BG-ALT-S4.xlsx",
    
    # BG-ALT-S5
    "L-BG3 ALT 1 - ALT Semestre 1": "BG-ALT-S5.xlsx",
    
    # BG-ALT-S6
    "L-BG3 ALT 1 - ALT Semestre 2": "BG-ALT-S6.xlsx",
}

async def get_template_id_from_group_name(db: Prisma, group_name: str) -> int:
    """
    Détermine l'ID du template Excel en fonction du nom du groupe.
    """
    template_filename = TEMPLATE_MAPPING.get(group_name)
    if not template_filename:
        raise ValueError(f"Aucun template trouvé pour le groupe : {group_name}")
        
    # Récupérer l'ID du template depuis la base de données
    template = await db.generatedfile.find_first(
        where={
            "filename": template_filename,
            "isTemplate": True
        }
    )
    
    if not template:
        raise ValueError(f"Template {template_filename} non trouvé dans la base de données")
        
    return template.id