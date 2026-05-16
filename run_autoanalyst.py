# run_nvda_goldman.py

import subprocess
import sys

cmd = [
    sys.executable,
    "autoanalyst_v3.py",

    "--company_name", "CoinShares PLC (CSHR)",
    "--short_name", "CoinShares",
    "--ticker", "CSHR",

    "--links_xlsx", "links/CSHR.xlsx",

    "--report_model", "gpt-5-nano",
    "--source_model", "gpt-5-nano",

    "--output_dir", "outputs/CSHR",

    # "--auto_accept_highlights"
]

subprocess.run(cmd, check=True)