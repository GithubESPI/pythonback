# absences_service.py
from typing import List, Dict

def process_absences(absences_data: List[Dict]):
    absences_summary = {}

    for absence in absences_data:
        apprenant_id = absence.get("codeApprenant")
        if not apprenant_id:
            continue

        duration = int(absence.get("duree", 0))
        is_justifie = absence.get("isJustifie", False)
        is_retard = absence.get("isRetard", False)

        if apprenant_id not in absences_summary:
            absences_summary[apprenant_id] = {
                "justified": [],
                "unjustified": [],
                "delays": []
            }

        if is_retard:
            absences_summary[apprenant_id]["delays"].append(duration)
        elif is_justifie:
            absences_summary[apprenant_id]["justified"].append(duration)
        else:
            absences_summary[apprenant_id]["unjustified"].append(duration)

    return absences_summary
