#!/usr/bin/env python3
"""
FHIR Resource Analyzer — QI Core 6.0.0

Recursively scans a directory for .ndjson files (one FHIR resource per line)
and evaluates field presence for specified data elements.

Usage:
    python fhir_analyzer.py <root_directory> [-o output.xlsx]

Requirements:
    pip install openpyxl
"""

import json
import sys
import argparse
from pathlib import Path
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl is required.  Run:  pip install openpyxl")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Field-presence helpers
# ---------------------------------------------------------------------------

def field_present(obj, path: str) -> bool:
    """
    Return True if `path` (dot-separated) resolves to a non-null, non-empty
    value within `obj`.  Arrays at any level are traversed — True if at least
    one element satisfies the remaining path.
    """
    if obj is None:
        return False
    if not path:
        return obj not in (None, [], "")

    key, _, rest = path.partition(".")

    if isinstance(obj, list):
        return any(field_present(item, path) for item in obj)

    if not isinstance(obj, dict):
        return False

    val = obj.get(key)
    if val is None or val == [] or val == "":
        return False

    if not rest:
        return True

    return field_present(val, rest)


def poly_present(obj, *field_names: str) -> bool:
    """True if any of the named polymorphic variants exist in `obj`."""
    if not isinstance(obj, dict):
        return False
    return any(
        obj.get(f) not in (None, [], "")
        for f in field_names
    )


def extension_present(resource: dict, url: str) -> bool:
    """True if the resource has a top-level extension with the given URL."""
    return any(
        isinstance(ext, dict) and ext.get("url") == url
        for ext in resource.get("extension", [])
    )


def category_lab_present(resource: dict) -> bool:
    """True if Observation.category contains a 'laboratory' coding."""
    for cat in resource.get("category", []):
        if not isinstance(cat, dict):
            continue
        for coding in cat.get("coding", []):
            if isinstance(coding, dict) and coding.get("code") == "laboratory":
                return True
    return False


def component_value_present(resource: dict) -> bool:
    """True if any Observation.component element carries a value[x] field."""
    VALUE_X = (
        "valueQuantity", "valueCodeableConcept", "valueString", "valueBoolean",
        "valueInteger", "valueRange", "valueRatio", "valueSampledData",
        "valueTime", "valueDateTime", "valuePeriod",
    )
    components = resource.get("component")
    if not components:
        return False
    return any(
        poly_present(comp, *VALUE_X)
        for comp in components
        if isinstance(comp, dict)
    )


def collected_x_present(resource: dict) -> bool:
    """True if Specimen.collection.collected[x] is present (any variant)."""
    collection = resource.get("collection")
    if not isinstance(collection, dict):
        return False
    return poly_present(collection, "collectedDateTime", "collectedPeriod")


# ---------------------------------------------------------------------------
# Check definitions
# (resource_type, display_element, value_set_url, check_fn)
# ---------------------------------------------------------------------------

CHECKS = [
    # ── Coverage ─────────────────────────────────────────────────────────────
    ("Coverage",
     "Coverage.period",
     "",
     lambda r: field_present(r, "period")),

    ("Coverage",
     "Coverage.status",
     "http://hl7.org/fhir/ValueSet/fm-status|4.0.1",
     lambda r: field_present(r, "status")),

    ("Coverage",
     "Coverage.type",
     "http://cts.nlm.nih.gov/fhir/ValueSet/2.16.840.1.114222.4.11.3591",
     lambda r: field_present(r, "type")),

    # ── Encounter ─────────────────────────────────────────────────────────────
    ("Encounter",
     "Encounter.hospitalization.admitSource",
     "https://hl7.org/fhir/R4/valueset-encounter-admit-source.html",
     lambda r: field_present(r, "hospitalization.admitSource")),

    ("Encounter",
     "Encounter.hospitalization.dischargeDisposition",
     "https://terminology.hl7.org/6.1.0/ValueSet-clinical-discharge-disposition.html",
     lambda r: field_present(r, "hospitalization.dischargeDisposition")),

    ("Encounter",
     "Encounter.period",
     "",
     lambda r: field_present(r, "period")),

    ("Encounter",
     "Encounter.period.end",
     "",
     lambda r: field_present(r, "period.end")),

    ("Encounter",
     "Encounter.period.start",
     "",
     lambda r: field_present(r, "period.start")),

    # ── Medication ────────────────────────────────────────────────────────────
    ("Medication",
     "Medication.code",
     "http://cts.nlm.nih.gov/fhir/ValueSet/2.16.840.1.113762.1.4.1010.4",
     lambda r: field_present(r, "code")),

    ("Medication",
     "Medication.status",
     "http://hl7.org/fhir/R4/valueset-medication-status.html",
     lambda r: field_present(r, "status")),

    # ── MedicationAdministration ──────────────────────────────────────────────
    ("MedicationAdministration",
     "MedicationAdministration.dosage.route",
     "http://hl7.org/fhir/ValueSet/route-codes",
     lambda r: field_present(r, "dosage.route")),

    ("MedicationAdministration",
     "MedicationAdministration.effective[x]",
     "",
     lambda r: poly_present(r, "effectiveDateTime", "effectivePeriod")),

    ("MedicationAdministration",
     "MedicationAdministration.effectivePeriod.end",
     "",
     lambda r: field_present(r, "effectivePeriod.end")),

    ("MedicationAdministration",
     "MedicationAdministration.effectivePeriod.start",
     "",
     lambda r: field_present(r, "effectivePeriod.start")),

    ("MedicationAdministration",
     "MedicationAdministration.medication[x]",
     "https://vsac.nlm.nih.gov/valueset/2.16.840.1.113762.1.4.1010.4/expansion",
     lambda r: poly_present(r, "medicationCodeableConcept", "medicationReference")),

    # ── MedicationRequest ─────────────────────────────────────────────────────
    ("MedicationRequest",
     "MedicationRequest.authoredOn",
     "",
     lambda r: field_present(r, "authoredOn")),

    ("MedicationRequest",
     "MedicationRequest.category",
     "http://hl7.org/fhir/ValueSet/medicationrequest-category",
     lambda r: field_present(r, "category")),

    ("MedicationRequest",
     "MedicationRequest.dosageInstruction.route",
     "https://hl7.org/fhir/valueset-route-codes.html",
     lambda r: field_present(r, "dosageInstruction.route")),

    ("MedicationRequest",
     "MedicationRequest.dosageInstruction.Timing",
     "",
     lambda r: field_present(r, "dosageInstruction.timing")),

    ("MedicationRequest",
     "MedicationRequest.dosageInstruction.Timing.boundsPeriod.end",
     "",
     lambda r: field_present(r, "dosageInstruction.timing.repeat.boundsPeriod.end")),

    ("MedicationRequest",
     "MedicationRequest.dosageInstruction.Timing.boundsPeriod.start",
     "",
     lambda r: field_present(r, "dosageInstruction.timing.repeat.boundsPeriod.start")),

    ("MedicationRequest",
     "MedicationRequest.medication[x]",
     "http://cts.nlm.nih.gov/fhir/ValueSet/2.16.840.1.113762.1.4.1010.4",
     lambda r: poly_present(r, "medicationCodeableConcept", "medicationReference")),

    ("MedicationRequest",
     "MedicationRequest.requester",
     "",
     lambda r: field_present(r, "requester")),

    # ── Observation (Laboratory Result) ───────────────────────────────────────
    ("Observation",
     "Observation.category",
     "http://hl7.org/fhir/valueset-observation-category.html",
     lambda r: field_present(r, "category")),

    ("Observation",
     "Observation.category:Laboratory",
     "http://hl7.org/fhir/us/core/ValueSet-us-core-clinical-result-observation-category.html",
     category_lab_present),

    ("Observation",
     "Observation.code",
     "http://hl7.org/fhir/us/core/ValueSet/us-core-laboratory-test-codes",
     lambda r: field_present(r, "code")),

    ("Observation",
     "Observation.component",
     "",
     lambda r: field_present(r, "component")),

    ("Observation",
     "Observation.component.code",
     "http://hl7.org/fhir/ValueSet/observation-codes",
     lambda r: field_present(r, "component.code")),

    ("Observation",
     "Observation.component.value[x]",
     "",
     component_value_present),

    ("Observation",
     "Observation.effective[x]",
     "",
     lambda r: poly_present(r, "effectiveDateTime", "effectivePeriod",
                            "effectiveInstant", "effectiveTiming")),

    ("Observation",
     "Observation.status",
     "http://hl7.org/fhir/us/qicore/ValueSet-qicore-non-negative-observation-status.html",
     lambda r: field_present(r, "status")),

    ("Observation",
     "Observation.subject",
     "",
     lambda r: field_present(r, "subject")),

    ("Observation",
     "Observation.value[x]",
     "",
     lambda r: poly_present(
         r,
         "valueQuantity", "valueCodeableConcept", "valueString", "valueBoolean",
         "valueInteger", "valueRange", "valueRatio", "valueSampledData",
         "valueTime", "valueDateTime", "valuePeriod",
     )),

    # ── Patient ───────────────────────────────────────────────────────────────
    ("Patient",
     "Patient.birthDate",
     "",
     lambda r: field_present(r, "birthDate")),

    ("Patient",
     "Patient.extension (race)",
     "http://hl7.org/fhir/us/core/StructureDefinition/us-core-race",
     lambda r: extension_present(r, "http://hl7.org/fhir/us/core/StructureDefinition/us-core-race")),

    ("Patient",
     "Patient.extension (sex at birth)",
     "http://hl7.org/fhir/us/core/StructureDefinition/us-core-birthsex",
     lambda r: extension_present(r, "http://hl7.org/fhir/us/core/StructureDefinition/us-core-birthsex")),

    ("Patient",
     "Patient.extension:ethnicity",
     "https://hl7.org/fhir/us/core/STU6.1/ValueSet-omb-ethnicity-category.html",
     lambda r: extension_present(r, "http://hl7.org/fhir/us/core/StructureDefinition/us-core-ethnicity")),

    ("Patient",
     "Patient.identifier",
     "",
     lambda r: field_present(r, "identifier")),

    ("Patient",
     "Patient.name",
     "",
     lambda r: field_present(r, "name")),

    ("Patient",
     "Patient.name.family",
     "",
     lambda r: field_present(r, "name.family")),

    ("Patient",
     "Patient.name.given",
     "",
     lambda r: field_present(r, "name.given")),

    # ── Specimen ──────────────────────────────────────────────────────────────
    ("Specimen",
     "Specimen.collection",
     "",
     lambda r: field_present(r, "collection")),

    ("Specimen",
     "Specimen.collection.bodySite",
     "http://hl7.org/fhir/valueset-body-site.html",
     lambda r: field_present(r, "collection.bodySite")),

    ("Specimen",
     "Specimen.collection.collected[x]",
     "",
     collected_x_present),

    ("Specimen",
     "Specimen.type",
     "https://vsac.nlm.nih.gov/valueset/2.16.840.1.113762.1.4.1099.54/expansion",
     lambda r: field_present(r, "type")),
]


# ---------------------------------------------------------------------------
# File loading
# ---------------------------------------------------------------------------

def find_ndjson_files(root: Path) -> list[Path]:
    return sorted(root.rglob("*.ndjson"))


def load_resources(files: list[Path]) -> tuple[dict, int, int]:
    """
    Parse every .ndjson file and group resources by resourceType.
    Returns (resources_by_type, total_lines, parse_error_count).
    """
    resources: dict[str, list[dict]] = defaultdict(list)
    total_lines = 0
    parse_errors = 0

    for path in files:
        try:
            with open(path, "r", encoding="utf-8") as fh:
                for line in fh:
                    line = line.strip()
                    if not line:
                        continue
                    total_lines += 1
                    try:
                        resource = json.loads(line)
                        rtype = resource.get("resourceType")
                        if rtype:
                            resources[rtype].append(resource)
                    except json.JSONDecodeError:
                        parse_errors += 1
        except OSError as exc:
            print(f"  WARNING: cannot read {path}: {exc}")

    return dict(resources), total_lines, parse_errors


# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

def analyze(resources: dict) -> list[dict]:
    rows = []
    for resource_type, element, value_set, check_fn in CHECKS:
        resource_list = resources.get(resource_type, [])
        total = len(resource_list)

        if total == 0:
            rows.append({
                "resource_type": resource_type,
                "element": element,
                "value_set": value_set,
                "total": 0,
                "missing": 0,
                "present": "N/A",
            })
            continue

        missing = sum(1 for r in resource_list if not check_fn(r))
        rows.append({
            "resource_type": resource_type,
            "element": element,
            "value_set": value_set,
            "total": total,
            "missing": missing,
            "present": "Yes" if missing == 0 else "No",
        })
    return rows


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

# Colour palette
_DARK_BLUE   = "1F4E79"
_MID_BLUE    = "2E75B6"
_LIGHT_BLUE  = "D6E4F0"
_WHITE       = "FFFFFF"
_GREEN_FILL  = "C6EFCE"
_GREEN_FONT  = "276221"
_RED_FILL    = "FFC7CE"
_RED_FONT    = "9C0006"
_AMBER_FILL  = "FFEB9C"
_AMBER_FONT  = "7D5A00"
_GREY_FILL   = "F2F2F2"

HEADERS = [
    "FHIR Resource\n(QI Core 6.0.0)",
    "FHIR Data Element",
    "FHIR Value Set",
    "Total\nResources",
    "Count\nMissing",
    "Present",
]

COL_WIDTHS = [32, 58, 70, 14, 14, 11]


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _font(bold=False, color=None, size=11) -> Font:
    kwargs = {"bold": bold, "size": size}
    if color:
        kwargs["color"] = color
    return Font(**kwargs)


def write_excel(rows: list[dict], output_path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "FHIR Element Analysis"

    center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    center_mid  = Alignment(horizontal="center", vertical="center")

    # ── Header row ──────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 36
    for col_idx, header_text in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header_text)
        cell.font      = _font(bold=True, color=_WHITE)
        cell.fill      = _fill(_DARK_BLUE)
        cell.alignment = center_wrap

    # ── Data rows ────────────────────────────────────────────────────────────
    prev_rtype = None
    rtype_color = _LIGHT_BLUE

    for row_idx, row in enumerate(rows, 2):
        ws.row_dimensions[row_idx].height = 18

        # Alternate the resource-type stripe colour when the type changes
        if row["resource_type"] != prev_rtype:
            rtype_color = _LIGHT_BLUE if rtype_color == _WHITE else _WHITE
            prev_rtype = row["resource_type"]

        stripe = _fill(rtype_color)

        # Columns 1-3: text
        for col_idx, value in enumerate(
            [row["resource_type"], row["element"], row["value_set"]], 1
        ):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.fill      = stripe
            cell.alignment = left_wrap
            cell.font      = _font()

        # Columns 4-5: numeric
        for col_idx, value in enumerate([row["total"], row["missing"]], 4):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.fill      = stripe
            cell.alignment = center_mid
            cell.font      = _font()

        # Column 6: Present — colour-coded
        present_cell = ws.cell(row=row_idx, column=6, value=row["present"])
        present_cell.alignment = center_mid
        if row["present"] == "Yes":
            present_cell.fill = _fill(_GREEN_FILL)
            present_cell.font = _font(bold=True, color=_GREEN_FONT)
        elif row["present"] == "No":
            present_cell.fill = _fill(_RED_FILL)
            present_cell.font = _font(bold=True, color=_RED_FONT)
        else:  # N/A
            present_cell.fill = _fill(_AMBER_FILL)
            present_cell.font = _font(bold=True, color=_AMBER_FONT)

    # ── Column widths ────────────────────────────────────────────────────────
    for col_idx, width in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "A2"
    wb.save(output_path)


# ---------------------------------------------------------------------------
# Summary sheet
# ---------------------------------------------------------------------------

def add_summary_sheet(wb: openpyxl.Workbook, rows: list[dict]) -> None:
    ws = wb.create_sheet("Summary", 0)
    ws.title = "Summary"

    center = Alignment(horizontal="center", vertical="center")
    left   = Alignment(horizontal="left",   vertical="center")

    ws.row_dimensions[1].height = 28
    for col_idx, header in enumerate(
        ["Resource Type", "Total Resources", "Elements Checked",
         "Fully Present", "Has Gaps", "N/A (no data)"], 1
    ):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font      = _font(bold=True, color=_WHITE)
        cell.fill      = _fill(_DARK_BLUE)
        cell.alignment = center

    # Aggregate by resource type
    from collections import Counter
    type_rows: dict[str, list[dict]] = defaultdict(list)
    for row in rows:
        type_rows[row["resource_type"]].append(row)

    for row_idx, (rtype, rrows) in enumerate(sorted(type_rows.items()), 2):
        totals  = [r["total"] for r in rrows]
        t_total = totals[0] if totals else 0  # same for all rows of a type
        n_yes   = sum(1 for r in rrows if r["present"] == "Yes")
        n_no    = sum(1 for r in rrows if r["present"] == "No")
        n_na    = sum(1 for r in rrows if r["present"] == "N/A")

        ws.cell(row=row_idx, column=1, value=rtype).alignment = left
        ws.cell(row=row_idx, column=2, value=t_total).alignment = center
        ws.cell(row=row_idx, column=3, value=len(rrows)).alignment = center
        ws.cell(row=row_idx, column=4, value=n_yes).alignment = center
        ws.cell(row=row_idx, column=5, value=n_no).alignment = center
        ws.cell(row=row_idx, column=6, value=n_na).alignment = center

        fill = _fill(_GREY_FILL) if row_idx % 2 == 0 else _fill(_WHITE)
        for c in range(1, 7):
            ws.cell(row=row_idx, column=c).fill = fill

    for col_idx, width in enumerate([28, 18, 18, 16, 12, 16], 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="FHIR Resource Analyzer — QI Core 6.0.0",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python fhir_analyzer.py /data/fhir_export\n"
            "  python fhir_analyzer.py /data/fhir_export -o results/analysis.xlsx\n"
        ),
    )
    parser.add_argument("directory",
                        help="Root directory to scan recursively for .ndjson files")
    parser.add_argument("-o", "--output", default="fhir_analysis.xlsx",
                        help="Output Excel file path (default: fhir_analysis.xlsx)")
    args = parser.parse_args()

    root = Path(args.directory)
    if not root.exists():
        print(f"ERROR: Directory not found: {root}")
        sys.exit(1)
    if not root.is_dir():
        print(f"ERROR: Not a directory: {root}")
        sys.exit(1)

    print(f"Scanning:  {root.resolve()}")
    files = find_ndjson_files(root)
    print(f"Found {len(files)} .ndjson file(s)")

    if not files:
        print("No .ndjson files found. Nothing to do.")
        sys.exit(0)

    print("Loading resources…")
    resources, total_lines, parse_errors = load_resources(files)

    print(f"  {total_lines:,} lines parsed")
    if parse_errors:
        print(f"  {parse_errors:,} lines skipped (JSON parse errors)")

    type_counts = {k: len(v) for k, v in resources.items()}
    relevant_types = {rtype for rtype, *_ in CHECKS}
    print(f"\nRelevant resource type counts:")
    for rtype in sorted(relevant_types):
        count = type_counts.get(rtype, 0)
        print(f"  {rtype:<30} {count:>8,}")

    other_types = sorted(set(type_counts) - relevant_types)
    if other_types:
        print(f"\nOther types found (not analysed): {', '.join(other_types)}")

    print("\nRunning checks…")
    rows = analyze(resources)

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws_detail = wb.active
    ws_detail.title = "FHIR Element Analysis"

    # Re-use the write logic but write into the existing workbook
    write_excel(rows, output_path)               # writes and saves once
    wb = openpyxl.load_workbook(output_path)     # re-open to add summary
    add_summary_sheet(wb, rows)
    wb.save(output_path)

    n_yes = sum(1 for r in rows if r["present"] == "Yes")
    n_no  = sum(1 for r in rows if r["present"] == "No")
    n_na  = sum(1 for r in rows if r["present"] == "N/A")
    print(
        f"\nResults: {n_yes} fully present | {n_no} have gaps | {n_na} N/A (no data)\n"
        f"Report:  {output_path.resolve()}"
    )


if __name__ == "__main__":
    main()
