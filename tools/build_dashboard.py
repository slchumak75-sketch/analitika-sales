"""Render the self-contained HTML dashboard from reports/summary.json."""
from __future__ import annotations

import json
import sys
from pathlib import Path


def main(summary_path: str, template_path: str, out_path: str) -> None:
    summary = json.loads(Path(summary_path).read_text(encoding="utf-8"))
    template = Path(template_path).read_text(encoding="utf-8")
    data = json.dumps(summary, ensure_ascii=False, separators=(",", ":"))
    html = template.replace("/*__DATA_JSON__*/", data)
    Path(out_path).write_text(html, encoding="utf-8")
    print(f"dashboard: {out_path} ({len(html) / 1024:.0f} KiB, {len(summary['items'])} items)")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
