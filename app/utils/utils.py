def convert_minutes_to_hours_and_minutes(minutes):
    """Convertit un nombre de minutes en format 'XhYm'"""
    if not minutes:
        return "0h0m"
    hours = minutes // 60
    remaining_minutes = minutes % 60
    return f"{hours}h{remaining_minutes}m"
