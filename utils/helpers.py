import pandas as pd

def to_juta(value):
    if pd.isna(value):
        return "-"
    
    if value < 1_000_000:
        return f"{int(value):,}".replace(",", ".")
        
    return f"{int(value / 1_000_000)} Juta"