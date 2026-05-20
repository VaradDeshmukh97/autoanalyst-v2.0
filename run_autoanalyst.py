import subprocess
import sys

cmd = [
    sys.executable,
    "autoanalyst_v4.py",

    "--company_name", "Eagle Nuclear Energy Corp. (NUCL)",
    "--short_name", "Eagle",
    "--ticker", "NUCL",

    "--links_xlsx", "links/NUCL.xlsx",

    "--report_model", "gpt-5-nano",
    "--source_model", "gpt-5-nano",

    "--output_dir", "outputs/NUCL",

    "--auto_accept_highlights"
]

subprocess.run(cmd, check=True)