# Install the aqarati-theme skill into Claude Code.
$ErrorActionPreference = "Stop"
claude plugin marketplace add Moosakalam/aqarati-plugins
claude plugin install aqarati-theme@aqarati-plugins
python -m pip install --user python-pptx reportlab
Write-Host "Done. Restart Claude Code to use the aqarati-theme skill."
