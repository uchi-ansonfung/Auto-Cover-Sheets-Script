#!/usr/bin/env python3
"""Generate CycloneDX SBOM + third-party license reports under licenses/.

Intended for local regeneration and the Windows full-installer CI job.
Does not need to run on every PR; committed copies under licenses/ are a
baseline for offline reading.
"""

from __future__ import annotations

import csv
import shutil
import subprocess
import sys
import time
from importlib.metadata import distribution
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "licenses"
PYPDFIUM_OUT = OUT / "pypdfium2"

# Omitted from the "shipped" human table (dev / SBOM tooling only).
DEV_ONLY = {
    "pytest",
    "iniconfig",
    "pluggy",
    "pygments",
    "colorama",
    "pyinstaller",
    "pyinstaller-hooks-contrib",
    "altgraph",
    "pefile",
    "pywin32-ctypes",
    "cyclonedx-bom",
    "cyclonedx-python-lib",
    "packageurl-python",
    "pip-licenses",
    "license-expression",
    "boolean.py",
    "boolean-py",
    "py-serializable",
    "sortedcontainers",
    "coversheets",
}


def run(cmd: list[str]) -> None:
    print("+", " ".join(cmd))
    subprocess.check_call(cmd, cwd=ROOT)


def ensure_tools() -> None:
    try:
        import cyclonedx_py  # noqa: F401
    except ImportError:
        run([sys.executable, "-m", "pip", "install", "cyclonedx-bom", "pip-licenses"])


def generate_cyclonedx() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    out = str(OUT / "sbom.cdx.json")
    try:
        run(
            [
                sys.executable,
                "-m",
                "cyclonedx_py",
                "environment",
                "-o",
                out,
                "--of",
                "JSON",
            ]
        )
    except subprocess.CalledProcessError:
        run(["cyclonedx-py", "environment", "-o", out, "--of", "JSON"])


def generate_pip_licenses() -> None:
    common = ["pip-licenses"]
    run(
        common
        + [
            "--format=markdown",
            "--with-urls",
            "--with-authors",
            "--order=name",
            f"--output-file={OUT / 'THIRD_PARTY_NOTICES.md'}",
        ]
    )
    run(
        common
        + [
            "--format=markdown",
            "--with-urls",
            "--order=license",
            f"--output-file={OUT / 'THIRD_PARTY_SUMMARY.md'}",
        ]
    )
    run(
        common
        + [
            "--format=csv",
            "--with-urls",
            f"--output-file={OUT / 'third_party.csv'}",
        ]
    )


def write_shipped_table() -> None:
    rows: list[dict[str, str]] = []
    with (OUT / "third_party.csv").open(encoding="utf-8") as fh:
        for row in csv.DictReader(fh):
            name = (row.get("Name") or "").strip()
            key = name.lower().replace("_", "-")
            if key in {d.lower() for d in DEV_ONLY}:
                continue
            if key.startswith("pip-") or "cyclonedx" in key:
                continue
            rows.append(row)

    lines = [
        "# Third-party components (product / full installer path)",
        "",
        "Generated from an environment with `coversheets[full,dev]` installed.",
        "Dev-only tooling (pytest, PyInstaller, SBOM generators) is omitted.",
        "",
        "Machine-readable SBOMs: [`sbom.cdx.json`](sbom.cdx.json) (CycloneDX 1.6).",
        "Full table including authors: [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md).",
        "PDFium nested licenses: [`pypdfium2/`](pypdfium2/).",
        "",
        "| Name | Version | License | URL |",
        "| --- | --- | --- | --- |",
    ]
    for row in sorted(rows, key=lambda r: (r.get("Name") or "").lower()):
        name = row.get("Name") or ""
        ver = row.get("Version") or ""
        lic = (row.get("License") or "").replace("|", r"\|")
        url = row.get("URL") or ""
        lines.append(f"| {name} | {ver} | {lic} | {url} |")
    (OUT / "THIRD_PARTY_SHIPPED.md").write_text(
        "\n".join(lines) + "\n", encoding="utf-8"
    )
    print(f"shipped components: {len(rows)}")


def _clear_dir(path: Path) -> None:
    """Remove contents of path; tolerate Windows/OneDrive locks."""
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)
        return
    for child in path.iterdir():
        for attempt in range(3):
            try:
                if child.is_dir():
                    shutil.rmtree(child)
                else:
                    child.unlink(missing_ok=True)
                break
            except OSError:
                if attempt == 2:
                    raise
                time.sleep(0.2)


def copy_pypdfium_licenses() -> None:
    _clear_dir(PYPDFIUM_OUT)
    try:
        dist = distribution("pypdfium2")
    except Exception as exc:
        print(f"skip pypdfium2 licenses: {exc}")
        return

    copied = 0
    for f in dist.files or []:
        s = str(f).replace("\\", "/")
        if "LICENSES" not in s and "LicenseRef" not in s and not s.endswith("LICENSE"):
            continue
        src = Path(dist.locate_file(f))
        if not src.is_file():
            continue
        dest = PYPDFIUM_OUT / src.name
        if dest.exists():
            dest = PYPDFIUM_OUT / f"{src.parent.name}_{src.name}"
        shutil.copy2(src, dest)
        copied += 1
    print(f"copied {copied} pypdfium2/PDFium license files")


def write_freeze() -> None:
    text = subprocess.check_output(
        [sys.executable, "-m", "pip", "freeze"],
        cwd=ROOT,
        text=True,
    )
    (OUT / "pip-freeze-full.txt").write_text(text, encoding="utf-8")
    print(f"pip freeze: {len(text.splitlines())} lines")


def main() -> int:
    ensure_tools()
    OUT.mkdir(parents=True, exist_ok=True)
    generate_cyclonedx()
    generate_pip_licenses()
    write_shipped_table()
    copy_pypdfium_licenses()
    write_freeze()
    print(f"Wrote reports under {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
