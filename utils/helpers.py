import pandas as pd

def to_juta(value):
    if pd.isna(value):
        return "-"

    try:
        val = float(value)
    except Exception:
        return str(value)

    # For values smaller than 1,000,000 show as decimal millions rounded (四捨五入)
    # e.g. 250000 -> 0.3jt, 950000 -> 1 Juta
    if abs(val) < 1_000_000:
        million = round(val / 1_000_000, 1)  # rounded to one decimal place
        if abs(million) >= 1.0:
            return f"{int(round(val / 1_000_000))} Juta"

        sign = "-" if million < 0 else ""
        million_abs = abs(million)
        # show as 0.<digit>jt (e.g. 0.3jt)
        digit = int(round(million_abs * 10))
        return f"{sign}0.{digit}jt"

    # For 1,000,000 and above, keep original integer million representation
    return f"{int(val / 1_000_000)} Juta"