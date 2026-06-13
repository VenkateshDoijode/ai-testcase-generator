"""
ReportGenerator — generates an HTML report for Jira Zephyr Scale bulk updates.
"""

import os
from datetime import datetime


class ReportGenerator:

    DEFAULT_FILENAME = "Jira_TC_Report.html"

    def __init__(self, output_path: str = "", action_label: str = "Updated", report_title: str = ""):
        self.action_label  = action_label
        self.report_title  = report_title or "Zephyr Scale Report"
        if not output_path:
            reports_dir = os.path.join(os.path.dirname(__file__), "..", "..", "reports")
            os.makedirs(reports_dir, exist_ok=True)
            output_path = os.path.join(reports_dir, self.DEFAULT_FILENAME)
        self.output_path = output_path

    def generate(
        self,
        test_cases: list,
        ok_list: list,
        fail_list: list,
        project_key: str,
        folder_path: str = "",
        **kwargs,
    ):
        """Build and write the HTML report to output_path."""
        self._write(test_cases, ok_list, fail_list, project_key, folder_path)

    def _write(self, test_cases, ok_list, fail_list, project_key, folder_path):
        output_path = self.output_path
        ok_set   = set(ok_list)
        fail_set = set(fail_list)
        total    = len(test_cases)
        run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        rows_html = ""
        for i, tc in enumerate(test_cases, 1):
            key  = tc["key"]
            name = tc["name"]
            if key in ok_set:
                badge     = f'<span class="badge ok">✅ {self.action_label}</span>'
                row_class = "row-ok"
            elif key in fail_set:
                badge     = '<span class="badge fail">❌ Failed</span>'
                row_class = "row-fail"
            else:
                badge     = '<span class="badge skip">— Skipped</span>'
                row_class = ""

            rows_html += f"""
            <tr class="{row_class}">
                <td>{i}</td>
                <td class="key">{key}</td>
                <td>{name}</td>
                <td>{badge}</td>
            </tr>"""

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Zephyr Scale Update Report</title>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: #f0f2f5;
      color: #1a1a2e;
      padding: 32px 24px;
    }}

    .container {{
      max-width: 1100px;
      margin: 0 auto;
    }}

    header {{
      background: linear-gradient(135deg, #0052cc 0%, #0747a6 100%);
      color: #fff;
      border-radius: 12px;
      padding: 28px 32px;
      margin-bottom: 24px;
    }}

    header h1 {{
      font-size: 1.6rem;
      font-weight: 700;
      letter-spacing: -0.3px;
    }}

    header p {{
      margin-top: 6px;
      font-size: 0.875rem;
      opacity: 0.85;
    }}

    .meta-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 12px;
      margin-bottom: 24px;
    }}

    .meta-card {{
      background: #fff;
      border-radius: 10px;
      padding: 16px 20px;
      box-shadow: 0 1px 4px rgba(0,0,0,.08);
    }}

    .meta-card .label {{
      font-size: 0.72rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.6px;
      color: #6b7280;
      margin-bottom: 4px;
    }}

    .meta-card .value {{
      font-size: 1rem;
      font-weight: 600;
      color: #111827;
      word-break: break-all;
    }}

    .stats {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 12px;
      margin-bottom: 28px;
    }}

    .stat {{
      background: #fff;
      border-radius: 10px;
      padding: 20px 24px;
      text-align: center;
      box-shadow: 0 1px 4px rgba(0,0,0,.08);
    }}

    .stat .number {{
      font-size: 2.2rem;
      font-weight: 700;
      line-height: 1;
    }}

    .stat .desc {{
      font-size: 0.8rem;
      color: #6b7280;
      margin-top: 6px;
      font-weight: 500;
    }}

    .stat.total   .number {{ color: #0052cc; }}
    .stat.success .number {{ color: #057a55; }}
    .stat.failed  .number {{ color: #e02424; }}

    .card {{
      background: #fff;
      border-radius: 10px;
      box-shadow: 0 1px 4px rgba(0,0,0,.08);
      overflow: hidden;
    }}

    .card-header {{
      padding: 16px 20px;
      font-weight: 600;
      font-size: 0.95rem;
      border-bottom: 1px solid #f3f4f6;
      display: flex;
      align-items: center;
      gap: 8px;
    }}

    table {{
      width: 100%;
      border-collapse: collapse;
    }}

    thead th {{
      background: #f9fafb;
      padding: 10px 16px;
      text-align: left;
      font-size: 0.78rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: #6b7280;
      border-bottom: 1px solid #e5e7eb;
    }}

    tbody td {{
      padding: 11px 16px;
      font-size: 0.88rem;
      border-bottom: 1px solid #f3f4f6;
      vertical-align: middle;
    }}

    tbody tr:last-child td {{ border-bottom: none; }}

    tbody tr:hover {{ background: #f9fafb; }}

    .row-ok   td:first-child {{ border-left: 3px solid #057a55; }}
    .row-fail td:first-child {{ border-left: 3px solid #e02424; }}

    .key {{
      font-family: SFMono-Regular, Consolas, monospace;
      font-size: 0.82rem;
      color: #0052cc;
      font-weight: 600;
    }}

    .badge {{
      display: inline-block;
      padding: 3px 10px;
      border-radius: 999px;
      font-size: 0.78rem;
      font-weight: 600;
    }}

    .badge.ok   {{ background: #def7ec; color: #03543f; }}
    .badge.fail {{ background: #fde8e8; color: #9b1c1c; }}
    .badge.skip {{ background: #f3f4f6; color: #6b7280; }}

    footer {{
      text-align: center;
      margin-top: 28px;
      font-size: 0.78rem;
      color: #9ca3af;
    }}
  </style>
</head>
<body>
  <div class="container">
    <header>
      <h1>{self.report_title}</h1>
      <p>Generated: {run_time}</p>
    </header>

    <div class="meta-grid">
      <div class="meta-card">
        <div class="label">Project</div>
        <div class="value">{project_key}</div>
      </div>
      <div class="meta-card">
        <div class="label">Folder</div>
        <div class="value">{folder_path}</div>
      </div>
    </div>

    <div class="stats">
      <div class="stat total">
        <div class="number">{total}</div>
        <div class="desc">Total Found</div>
      </div>
      <div class="stat success">
        <div class="number">{len(ok_list)}</div>
        <div class="desc">{self.action_label}</div>
      </div>
      <div class="stat failed">
        <div class="number">{len(fail_list)}</div>
        <div class="desc">Failed</div>
      </div>
    </div>

    <div class="card">
      <div class="card-header">&#x1F4CB; Test Case Results</div>
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Key</th>
            <th>Name</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          {rows_html}
        </tbody>
      </table>
    </div>

    <footer>Jira Testcase Updater &nbsp;&nbsp; {run_time}</footer>
  </div>
</body>
</html>"""

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)

        print(f"\n  📄  HTML report saved → {os.path.abspath(output_path)}")
