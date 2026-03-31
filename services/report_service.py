from config.stage_map import stage_map
from utils.formatter import format_row

def build_report(df, translator):
    grouped = {}
    for stage, jp in stage_map.items():
        grouped.setdefault(jp, []).append(stage)

    lines = []

    for jp_title, stages in grouped.items():
        subset = df[df["Stage"].isin(stages)]

        if subset.empty:
            continue

        lines.append(f"■ {jp_title}")

        for _, row in subset.iterrows():
            lines.append(format_row(row, translator))

        lines.append("")

    return lines


def save_report(lines, path):
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))