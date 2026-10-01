#!/usr/bin/env sh
# Install the aqarati-theme skill into Claude Code.
set -e
claude plugin marketplace add Moosakalam/aqarati-plugins
claude plugin install aqarati-theme@aqarati-plugins
python3 -m pip install --user python-pptx reportlab || pip3 install --user python-pptx reportlab
echo "Done. Restart Claude Code to use the aqarati-theme skill."
