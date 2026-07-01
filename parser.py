"""SPDX JSON parser — extracts structured data from an SPDX document."""

import json
import pathlib
from typing import Any

from utils import flatten_checksums, flatten_list, logger, safe_get


class SPDXParseError(Exception):
    """Raised when the file cannot be parsed as a valid SPDX document."""


def load_spdx(path: pathlib.Path) -> dict:
    """Load and validate an SPDX JSON file.

    Args:
        path: Path to the .json file.

    Returns:
        Parsed SPDX data as a dict.

    Raises:
        SPDXParseError: On any parsing or validation failure.
    """
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise SPDXParseError(f"Cannot read file: {exc}") from exc

    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise SPDXParseError(f"Invalid JSON: {exc}") from exc

    if not isinstance(data, dict):
        raise SPDXParseError("SPDX document must be a JSON object.")

    if "spdxVersion" not in data and "SPDXVersion" not in data:
        raise SPDXParseError(
            "File does not appear to be an SPDX document "
            "(missing 'spdxVersion' / 'SPDXVersion' field)."
        )

    logger.info("Loaded SPDX file: %s", path)
    return data


# ---------------------------------------------------------------------------
# Section extractors
# ---------------------------------------------------------------------------

def extract_document_info(data: dict) -> list[dict]:
    """Extract top-level document metadata."""
    creation = data.get("creationInfo", {})
    creators = creation.get("creators", [])
    return [
        {
            "SPDXID": safe_get(data, "SPDXID"),
            "SPDXVersion": safe_get(data, "spdxVersion"),
            "DataLicense": safe_get(data, "dataLicense"),
            "Document Name": safe_get(data, "name"),
            "Document Namespace": safe_get(data, "documentNamespace"),
            "Creators": flatten_list(creators),
            "Created Date": safe_get(creation, "created"),
            "LicenseListVersion": safe_get(creation, "licenseListVersion"),
            "Comment": safe_get(data, "comment"),
        }
    ]


def extract_packages(data: dict) -> list[dict]:
    """Extract package entries."""
    rows: list[dict] = []
    for pkg in data.get("packages", []):
        if not isinstance(pkg, dict):
            continue
        rows.append(
            {
                "Package Name": safe_get(pkg, "name"),
                "SPDXID": safe_get(pkg, "SPDXID"),
                "Version": safe_get(pkg, "versionInfo"),
                "Supplier": safe_get(pkg, "supplier"),
                "Download Location": safe_get(pkg, "downloadLocation"),
                "Homepage": safe_get(pkg, "homepage"),
                "License Concluded": safe_get(pkg, "licenseConcluded"),
                "License Declared": safe_get(pkg, "licenseDeclared"),
                "Copyright": safe_get(pkg, "copyrightText"),
                "Checksums": flatten_checksums(pkg.get("checksums")),
                "Files Analyzed": str(pkg.get("filesAnalyzed", "")),
                "Summary": safe_get(pkg, "summary"),
                "Description": safe_get(pkg, "description"),
                "Comment": safe_get(pkg, "comment"),
            }
        )
    logger.info("Extracted %d packages", len(rows))
    return rows


def extract_files(data: dict) -> list[dict]:
    """Extract file entries."""
    rows: list[dict] = []
    for f in data.get("files", []):
        if not isinstance(f, dict):
            continue
        rows.append(
            {
                "File Name": safe_get(f, "fileName"),
                "SPDXID": safe_get(f, "SPDXID"),
                "License Concluded": safe_get(f, "licenseConcluded"),
                "License Info In File": flatten_list(f.get("licenseInfoInFiles")),
                "Copyright": safe_get(f, "copyrightText"),
                "Checksums": flatten_checksums(f.get("checksums")),
                "Comment": safe_get(f, "comment"),
            }
        )
    logger.info("Extracted %d files", len(rows))
    return rows


def extract_relationships(data: dict) -> list[dict]:
    """Extract relationship entries."""
    rows: list[dict] = []
    for rel in data.get("relationships", []):
        if not isinstance(rel, dict):
            continue
        rows.append(
            {
                "SPDX Element ID": safe_get(rel, "spdxElementId"),
                "Relationship Type": safe_get(rel, "relationshipType"),
                "Related SPDX Element": safe_get(rel, "relatedSpdxElement"),
                "Comment": safe_get(rel, "comment"),
            }
        )
    logger.info("Extracted %d relationships", len(rows))
    return rows


def extract_extracted_licensing_info(data: dict) -> list[dict]:
    """Extract extracted (custom) license info."""
    rows: list[dict] = []
    for lic in data.get("hasExtractedLicensingInfos", []):
        if not isinstance(lic, dict):
            continue
        rows.append(
            {
                "License ID": safe_get(lic, "licenseId"),
                "Name": safe_get(lic, "name"),
                "Extracted Text": safe_get(lic, "extractedText"),
                "Comment": safe_get(lic, "comment"),
            }
        )
    logger.info("Extracted %d licensing info entries", len(rows))
    return rows


def parse_spdx(data: dict) -> dict[str, list[dict]]:
    """Run all extractors and return a mapping of sheet name → rows.

    Sections with zero rows are omitted so no empty sheets appear.
    """
    extractors: dict[str, Any] = {
        "Document Info": extract_document_info,
        "Packages": extract_packages,
        "Files": extract_files,
        "Relationships": extract_relationships,
        "Extracted Licensing Info": extract_extracted_licensing_info,
    }
    result: dict[str, list[dict]] = {}
    for sheet_name, fn in extractors.items():
        try:
            rows = fn(data)
            if rows:
                result[sheet_name] = rows
        except Exception as exc:  # pragma: no cover
            logger.warning("Error extracting '%s': %s", sheet_name, exc)
    return result
