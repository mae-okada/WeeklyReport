from utils.cleaner import clean_company_name, clean_project_name
from utils.helpers import to_juta
from config.stage_map import stage_map
from services.translator import translate_project
import pandas as pd

def format_row(row, translator):
    company = clean_company_name(row.get("Company", "-"))
    size = to_juta(row.get("Size", 0))
    stage = str(row.get("Stage", ""))
    deal_id = str(row.get("ID", ""))

    project = clean_project_name(row.get("Name", "-"))
    project_jp = translate_project(project, translator)

    text = f"・{company} ： {size} / {project_jp}"

    # if not stage.startswith("1-"):
    #     text += " / <元野記入>"
        
    if stage.startswith("3"):
        text += " / 注文書待ち"
    elif stage.startswith("4") or stage == "5.  Sales (Invoice)":
        text += " / 請求書発行"
    
    text += f" / {deal_id}"

    stage_old = row.get("Stage_old")

    if stage_old is None:
        owner = row.get("Owner Fullname", "").split()[-2]
        text += f" / {owner}"

    elif pd.isna(stage_old):
        text += " / 【初】"

    else:
        jp_old_stage = stage_map.get(str(stage_old).strip(), stage_old)
        jp_new_stage = stage_map.get(str(stage).strip(), stage)
        text += f" / 【{jp_old_stage} → {jp_new_stage}】"

    return text