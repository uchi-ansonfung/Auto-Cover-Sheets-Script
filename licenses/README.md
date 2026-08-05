# Third-party licenses and SBOM

Machine-assisted inventory for **Automatic Exhibit Cover Sheets** when built with
the full OCR + optimize extras (`pip install -e ".[full,dev]"`).

## Files

| File | Purpose |
|------|---------|
| [`sbom.cdx.json`](sbom.cdx.json) | **CycloneDX 1.6** Software Bill of Materials (JSON) |
| [`THIRD_PARTY_SHIPPED.md`](THIRD_PARTY_SHIPPED.md) | Human-readable licenses for product-relevant packages |
| [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md) | Full pip-licenses dump (name, version, license, author, URL) |
| [`THIRD_PARTY_SUMMARY.md`](THIRD_PARTY_SUMMARY.md) | Same list ordered by license |
| [`third_party.csv`](third_party.csv) | CSV for spreadsheets / tooling |
| [`pip-freeze-full.txt`](pip-freeze-full.txt) | Exact `pip freeze` of the generation environment |
| [`pypdfium2/`](pypdfium2/) | License texts for pypdfium2 and **bundled PDFium** third parties |

## App license

Cover Sheets itself is **GPL-3.0-or-later** (see [`../LICENSE`](../LICENSE)).

## Bundled native tool (not in the Python SBOM)

The Windows full installer also ships **Tesseract OCR** (Apache-2.0) next to the
exe. Tesseract is not a pip package, so it does not appear in `sbom.cdx.json`.
When redistributing that installer, include Apache-2.0 notices for Tesseract
and the English traineddata files.

## Regenerate

```powershell
pip install -e ".[full,dev]" cyclonedx-bom pip-licenses
python scripts/generate_sbom.py
```

## Tools used

- [cyclonedx-bom](https://pypi.org/project/cyclonedx-bom/) — environment SBOM
- [pip-licenses](https://pypi.org/project/pip-licenses/) — license tables