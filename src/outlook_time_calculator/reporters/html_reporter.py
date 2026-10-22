from __future__ import annotations

from datetime import date
from pathlib import Path

from outlook_time_calculator.models import CategoryRow


def write_html(
    rows: list[CategoryRow],
    output_dir: Path,
    start: date,
    end: date,
) -> Path:
    """Write aggregated category rows to a self-contained HTML file and return the output path."""
    output_dir.mkdir(parents=True, exist_ok=True)
    filename = f"meeting_report_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.html"
    output_path = output_dir / filename

    grand_total = round(sum(r.total_hours for r in rows), 2)

    row_html = "\n".join(
        f"""        <tr>
          <td>{_esc(row.category)}</td>
          <td class="num">{row.accepted_hours}</td>
          <td class="num">{row.tentative_hours}</td>
          <td class="num">{row.declined_hours}</td>
          <td class="num total">{row.total_hours}</td>
        </tr>"""
        for row in rows
    )

    date_range = f"{start.strftime('%d %b %Y')} – {end.strftime('%d %b %Y')}"

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Meeting Report {date_range}</title>
  <style>
    body {{
      font-family: -apple-system, "Segoe UI", system-ui, sans-serif;
      font-size: 14px;
      color: #1f2328;
      background: #ffffff;
      margin: 2rem auto;
      max-width: 760px;
      padding: 0 1rem;
    }}
    h1 {{ font-size: 1.25rem; margin-bottom: 0.25rem; }}
    p.subtitle {{ color: #57606a; margin-top: 0; }}
    table {{
      width: 100%;
      border-collapse: collapse;
      margin-top: 1rem;
    }}
    th, td {{
      padding: 0.5rem 0.75rem;
      border: 1px solid #e5e7eb;
      text-align: left;
    }}
    th {{
      background: #f7f8fa;
      font-weight: 600;
    }}
    tr:nth-child(even) td {{ background: #f7f8fa; }}
    .num {{ text-align: right; }}
    .total {{ font-weight: 600; }}
    tfoot td {{
      background: #f0f4ff;
      font-weight: 700;
      border-top: 2px solid #3b82d4;
    }}
  </style>
</head>
<body>
  <h1>Meeting Time by Category</h1>
  <p class="subtitle">{_esc(date_range)}</p>
  <table>
    <thead>
      <tr>
        <th>Category</th>
        <th class="num">Accepted Hours</th>
        <th class="num">Tentative Hours</th>
        <th class="num">Declined Hours</th>
        <th class="num">Total Hours</th>
      </tr>
    </thead>
    <tbody>
{row_html}
    </tbody>
    <tfoot>
      <tr>
        <td>Grand Total</td>
        <td class="num"></td>
        <td class="num"></td>
        <td class="num"></td>
        <td class="num">{grand_total}</td>
      </tr>
    </tfoot>
  </table>
</body>
</html>
"""

    output_path.write_text(html, encoding="utf-8")
    return output_path


def _esc(text: str) -> str:
    """Minimal HTML escaping for safe text insertion."""
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
    )
