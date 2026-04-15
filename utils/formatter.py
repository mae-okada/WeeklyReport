from utils.cleaner import clean_company_name, clean_project_name
from utils.helpers import to_juta
from config.stage_map import stage_map
import pandas as pd

def format_row(row, translator, use_name=False, use_stage=True):
    company = clean_company_name(row.get("Company", "-"))
    
    size_val = row.get("Size", 0)
    size_old_val = row.get("Size_old")
    
    size = to_juta(size_val)
    
    if pd.notna(size_old_val):
        try:
            diff = size_val - size_old_val
            if diff != 0:
                diff_text = to_juta(abs(diff))
                sign = "+" if diff > 0 else "-"
                size = f"{size} ({sign}{diff_text})"
        except:
            pass 
    
    stage = str(row.get("Stage", ""))
    deal_id = str(row.get("ID", ""))

    project = clean_project_name(row.get("Name", "-"))
    project_jp = translator.translate_project(project) if translator else project

    text = f"・{company} ： {size}"

    if stage.startswith("3"):
        goal_grade = row.get("3-3. Goal Grade to S/O（〃）*", "")
        goal_grade = str(goal_grade).split(" ")[0] if pd.notna(goal_grade) else ""    
        text += f" / ランク{goal_grade}"
        
    text += f" / {project_jp}"

    # if not stage.startswith("1-"):
    #     text += " / <元野記入>"
        
    if stage.startswith("3"):
        goal_grade = row.get("3-3. Goal Grade to S/O（〃）*", "")
        goal_grade = str(goal_grade).split(" ")[0] if pd.notna(goal_grade) else ""    
        text += f" / 注文書待ち"
    elif stage.startswith("4"):
        text += " / 請求書発行"
    elif stage == "1-2. Potential (Renewal)":
        effective_date = row.get("1-2. Effective Date (if 1-1 YES) *", "")
        if pd.notna(effective_date):
            effective_date = pd.to_datetime(
                    effective_date,
                    format="%d/%m/%Y",
                    errors="coerce"
                )
            if pd.notna(effective_date):
                text += f" / 更新日： {effective_date.date()}"
    elif stage == "5.  Sales (Invoice)":
        text += " / 請求書送付済み"
    
    text += f" / {deal_id}"

    if use_name:
        owner = row.get("Owner Fullname", "")
        text += f" / {owner}"

    if use_stage:
        stage_old = row.get("Stage_old")
        if pd.isna(stage_old):
            if project.lower().__contains__("renewal"):    
                jp_new_stage = stage_map.get(str(stage).strip(), stage)            
                text += f" / 【新規案件 → {jp_new_stage}】"
            else:
                text += " / 【初】"
        else:
            jp_old_stage = stage_map.get(str(stage_old).strip(), stage_old)
            jp_new_stage = stage_map.get(str(stage).strip(), stage)
            text += f" / 【{jp_old_stage} → {jp_new_stage}】"

    return text