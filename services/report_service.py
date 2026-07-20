from config.stage_map import stage_map
from utils.formatter import format_row


class ReportService:
    """Service for building and saving text report output."""

    def __init__(self, translator=None):
        """Initialize the report service with an optional translator."""
        self.translator = translator

    def build_report(self, df, use_name=False, use_stage=False):
        """Build a list of report lines grouped by Japanese stage labels."""
        grouped = {}
        for stage, jp in stage_map.items():
            grouped.setdefault(jp, []).append(stage)

        lines = []
        for jp_title, stages in grouped.items():
            subset = df[df["Stage"].isin(stages)]
            # Remove renewal projects from stages 1-1 and 1-2 entirely from the report
            try:
                mask_stage_1 = subset["Stage"].astype(str).str.startswith("1-2")
                mask_renewal = subset["Name"].astype(str).str.lower().str.contains("renewal")
                subset = subset[~(mask_stage_1 & mask_renewal)]
            except Exception:
                # If any unexpected issues occur (missing columns/types), skip filtering
                pass
            lines.append(f"■ {jp_title}")

            if subset.empty:
                lines.append("<なし>")
                lines.append("")
                continue

            for _, row in subset.iterrows():
                lines.append(format_row(row, self.translator, use_name, use_stage))

            lines.append("")

        return lines

    def save_report(self, lines, path):
        """Write the prepared report lines to a UTF-8 text file."""
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))