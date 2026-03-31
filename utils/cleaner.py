def clean_company_name(name):
    if not isinstance(name, str):
        return "-"
    return name.split("-")[0].strip()


def clean_project_name(name):
    if not isinstance(name, str):
        return "-"
    
    parts = name.split(" - ", 1)
    return parts[1].strip() if len(parts) > 1 else name.strip()