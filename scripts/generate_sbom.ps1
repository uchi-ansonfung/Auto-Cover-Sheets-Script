$ErrorActionPreference = "Stop"
$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $RepoRoot
python -m pip install -q "cyclonedx-bom>=4.0" "pip-licenses>=5.0"
python scripts/generate_sbom.py
Write-Host "Done. See licenses\README.md"