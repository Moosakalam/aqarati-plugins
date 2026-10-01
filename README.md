# aqarati-plugins

Claude Code plugin marketplace with the Aqarati deck theme skill.

## Install (one command)

macOS / Linux:

```
curl -fsSL https://raw.githubusercontent.com/Moosakalam/aqarati-plugins/main/install.sh | sh
```

Windows (PowerShell):

```
irm https://raw.githubusercontent.com/Moosakalam/aqarati-plugins/main/install.ps1 | iex
```

Requires the `claude` CLI and Python 3 on PATH. Restart Claude Code afterwards.

## Manual install

```
claude plugin marketplace add Moosakalam/aqarati-plugins
claude plugin install aqarati-theme@aqarati-plugins
pip3 install python-pptx reportlab
```

Or inside Claude Code: `/plugin marketplace add Moosakalam/aqarati-plugins`, then `/plugin install aqarati-theme@aqarati-plugins`.
