#!/usr/bin/env python3
"""
FHIR Resource Analyzer — QI Core 6.0.0

Recursively scans a directory for .ndjson files (one FHIR resource per line)
and evaluates field presence, code system usage, status distribution, and
profile adoption for specified QI Core data elements.

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


# ── Constants ────────────────────────────────────────────────────────────────

STANDARD_SYSTEMS = {
    "http://www.nlm.nih.gov/research/umls/rxnorm",   # RxNorm
    "http://snomed.info/sct",                          # SNOMED CT
    "http://loinc.org",                                # LOINC
}


# ── Field-presence helpers ───────────────────────────────────────────────────

def field_present(obj, path: str) -> bool:
    """
    True if the dot-separated path resolves to a non-null, non-empty value.
    Arrays at any level are traversed — returns True if any element satisfies
    the remaining path.
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


def poly_present(obj, *fields: str) -> bool:
    """True if any of the named polymorphic variants are present."""
    if not isinstance(obj, dict):
        return False
    return any(obj.get(f) not in (None, [], "") for f in fields)


def extension_present(resource: dict, url: str) -> bool:
    """True if the resource carries a top-level extension with the given URL."""
    return any(
        isinstance(e, dict) and e.get("url") == url
        for e in resource.get("extension", [])
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
    comps = resource.get("component")
    if not comps:
        return False
    return any(poly_present(c, *VALUE_X) for c in comps if isinstance(c, dict))


def collected_x_present(resource: dict) -> bool:
    """True if Specimen.collection.collected[x] is present in any variant."""
    col = resource.get("collection")
    if not isinstance(col, dict):
        return False
    return poly_present(col, "collectedDateTime", "collectedPeriod")


# ── Code-system extraction helpers ───────────────────────────────────────────

def navigate(obj, path: str):
    """
    Navigate a dot-separated path, flattening arrays at every level.
    Returns the terminal value(s) or None.
    """
    if not path or obj is None:
        return obj
    key, _, rest = path.partition(".")
    if isinstance(obj, list):
        results = []
        for item in obj:
            r = navigate(item, path)
            if r is not None:
                results.extend(r if isinstance(r, list) else [r])
        return results if results else None
    if not isinstance(obj, dict):
        return None
    val = obj.get(key)
    if val is None:
        return None
    if not rest:
        return val
    return navigate(val, rest)


def systems_from_cc(val) -> list:
    """
    Extract coding.system values from a CodeableConcept or a list of them.
    This is the core extraction — every standardised code (RxNorm, SNOMED,
    LOINC, etc.) lives in coding[].system inside a CodeableConcept.
    """
    if val is None:
        return []
    items = val if isinstance(val, list) else [val]
    out = []
    for item in items:
        if isinstance(item, dict):
            for coding in item.get("coding", []):
                if isinstance(coding, dict) and coding.get("system"):
                    out.append(coding["system"])
    return out


def cc_systems(resource: dict, path: str) -> list:
    """Extract code systems from a CodeableConcept reached via dot-path."""
    return systems_from_cc(navigate(resource, path))


def poly_cc_systems(resource: dict, *cc_field_names: str) -> list:
    """Extract code systems from polymorphic CodeableConcept fields."""
    out = []
    for f in cc_field_names:
        val = resource.get(f)
        if isinstance(val, dict):
            out.extend(systems_from_cc(val))
    return out


def extract_race_systems(resource: dict) -> list:
    """
    Extract coding systems from the us-core-race extension.
    Race codes live in nested extensions with url 'ombCategory' or 'detailed',
    carrying a valueCoding with a system URI.
    """
    out = []
    for ext in resource.get("extension", []):
        if not isinstance(ext, dict):
            continue
        if ext.get("url") == "http://hl7.org/fhir/us/core/StructureDefinition/us-core-race":
            for sub in ext.get("extension", []):
                if not isinstance(sub, dict):
                    continue
                if sub.get("url") in ("ombCategory", "detailed"):
                    vc = sub.get("valueCoding", {})
                    if isinstance(vc, dict) and vc.get("system"):
                        out.append(vc["system"])
    return out


def extract_ethnicity_systems(resource: dict) -> list:
    """
    Extract coding systems from the us-core-ethnicity extension.
    Same nested structure as us-core-race.
    """
    out = []
    for ext in resource.get("extension", []):
        if not isinstance(ext, dict):
            continue
        if ext.get("url") == "http://hl7.org/fhir/us/core/StructureDefinition/us-core-ethnicity":
            for sub in ext.get("extension", []):
                if not isinstance(sub, dict):
                    continue
                if sub.get("url") in ("ombCategory", "detailed"):
                    vc = sub.get("valueCoding", {})
                    if isinstance(vc, dict) and vc.get("system"):
                        out.append(vc["system"])
    return out


# ── Check definitions ────────────────────────────────────────────────────────

def chk(resource_type, element, value_set, check_fn,
        extract_systems_fn=None, is_status=False, status_path=None):
    """Build a check definition dict."""
    return dict(
        resource_type=resource_type,
        element=element,
        value_set=value_set,
        check_fn=check_fn,
        extract_systems_fn=extract_systems_fn,
        is_status=is_status,
        status_path=status_path,
    )


CHECKS = [

    # ── Coverage ─────────────────────────────────────────────────────────────
    chk("Coverage", "Coverage.period", "",
        lambda r: field_present(r, "period")),

    chk("Coverage", "Coverage.status",
        "http://hl7.org/fhir/ValueSet/fm-status|4.0.1",
        lambda r: field_present(r, "status"),
        is_status=True, status_path="status"),

    chk("Coverage", "Coverage.type",
        "http://cts.nlm.nih.gov/fhir/ValueSet/2.16.840.1.114222.4.11.3591",
        lambda r: field_present(r, "type"),
        extract_systems_fn=lambda r: cc_systems(r, "type")),

    # ── Encounter ─────────────────────────────────────────────────────────────
    chk("Encounter", "Encounter.hospitalization.admitSource",
        "https://hl7.org/fhir/R4/valueset-encounter-admit-source.html",
        lambda r: field_present(r, "hospitalization.admitSource"),
        extract_systems_fn=lambda r: cc_systems(r, "hospitalization.admitSource")),

    chk("Encounter", "Encounter.hospitalization.dischargeDisposition",
        "https://terminology.hl7.org/6.1.0/ValueSet-clinical-discharge-disposition.html",
        lambda r: field_present(r, "hospitalization.dischargeDisposition"),
        extract_systems_fn=lambda r: cc_systems(r, "hospitalization.dischargeDisposition")),

    chk("Encounter", "Encounter.period", "",
        lambda r: field_present(r, "period")),

    chk("Encounter", "Encounter.period.end", "",
        lambda r: field_present(r, "period.end")),

    chk("Encounter", "Encounter.period.start", "",
        lambda r: field_present(r, "period.start")),

    # ── Medication ────────────────────────────────────────────────────────────
    chk("Medication", "Medication.code",
        "http://cts.nlm.nih.gov/fhir/ValueSet/2.16.840.1.113762.1.4.1010.4",
        lambda r: field_present(r, "code"),
        extract_systems_fn=lambda r: cc_systems(r, "code")),

    chk("Medication", "Medication.status",
        "http://hl7.org/fhir/R4/valueset-medication-status.html",
        lambda r: field_present(r, "status"),
        is_status=True, status_path="status"),

    # ── MedicationAdministration ──────────────────────────────────────────────
    chk("MedicationAdministration", "MedicationAdministration.dosage.route",
        "http://hl7.org/fhir/ValueSet/route-codes",
        lambda r: field_present(r, "dosage.route"),
        extract_systems_fn=lambda r: cc_systems(r, "dosage.route")),

    chk("MedicationAdministration", "MedicationAdministration.effective[x]", "",
        lambda r: poly_present(r, "effectiveDateTime", "effectivePeriod")),

    chk("MedicationAdministration", "MedicationAdministration.effectivePeriod.end", "",
        lambda r: field_present(r, "effectivePeriod.end")),

    chk("MedicationAdministration", "MedicationAdministration.effectivePeriod.start", "",
        lambda r: field_present(r, "effectivePeriod.start")),

    chk("MedicationAdministration", "MedicationAdministration.medication[x]",
        "https://vsac.nlm.nih.gov/valueset/2.16.840.1.113762.1.4.1010.4/expansion",
        lambda r: poly_present(r, "medicationCodeableConcept", "medicationReference"),
        # RxNorm codes live inside medicationCodeableConcept.coding[].system
        extract_systems_fn=lambda r: poly_cc_systems(r, "medicationCodeableConcept")),

    # ── MedicationRequest ─────────────────────────────────────────────────────
    chk("MedicationRequest", "MedicationRequest.authoredOn", "",
        lambda r: field_present(r, "authoredOn")),

    chk("MedicationRequest", "MedicationRequest.category",
        "http://hl7.org/fhir/ValueSet/medicationrequest-category",
        lambda r: field_present(r, "category"),
        extract_systems_fn=lambda r: cc_systems(r, "category")),

    chk("MedicationRequest", "MedicationRequest.dosageInstruction.route",
        "https://hl7.org/fhir/valueset-route-codes.html",
        lambda r: field_present(r, "dosageInstruction.route"),
        extract_systems_fn=lambda r: cc_systems(r, "dosageInstruction.route")),

    chk("MedicationRequest", "MedicationRequest.dosageInstruction.Timing", "",
        lambda r: field_present(r, "dosageInstruction.timing")),

    chk("MedicationRequest", "MedicationRequest.dosageInstruction.Timing.boundsPeriod.end", "",
        lambda r: field_present(r, "dosageInstruction.timing.repeat.boundsPeriod.end")),

    chk("MedicationRequest", "MedicationRequest.dosageInstruction.Timing.boundsPeriod.start", "",
        lambda r: field_present(r, "dosageInstruction.timing.repeat.boundsPeriod.start")),

    chk("MedicationRequest", "MedicationRequest.medication[x]",
        "http://cts.nlm.nih.gov/fhir/ValueSet/2.16.840.1.113762.1.4.1010.4",
        lambda r: poly_present(r, "medicationCodeableConcept", "medicationReference"),
        # RxNorm codes live inside medicationCodeableConcept.coding[].system
        extract_systems_fn=lambda r: poly_cc_systems(r, "medicationCodeableConcept")),

    chk("MedicationRequest", "MedicationRequest.requester", "",
        lambda r: field_present(r, "requester")),

    # ── Observation (Laboratory Result) ───────────────────────────────────────
    chk("Observation", "Observation.category",
        "http://hl7.org/fhir/valueset-observation-category.html",
        lambda r: field_present(r, "category"),
        extract_systems_fn=lambda r: cc_systems(r, "category")),

    chk("Observation", "Observation.category:Laboratory",
        "http://hl7.org/fhir/us/core/ValueSet-us-core-clinical-result-observation-category.html",
        category_lab_present,
        extract_systems_fn=lambda r: cc_systems(r, "category")),

    chk("Observation", "Observation.code",
        "http://hl7.org/fhir/us/core/ValueSet/us-core-laboratory-test-codes",
        lambda r: field_present(r, "code"),
        # LOINC codes live in code.coding[].system
        extract_systems_fn=lambda r: cc_systems(r, "code")),

    chk("Observation", "Observation.component", "",
        lambda r: field_present(r, "component")),

    chk("Observation", "Observation.component.code",
        "http://hl7.org/fhir/ValueSet/observation-codes",
        lambda r: field_present(r, "component.code"),
        extract_systems_fn=lambda r: cc_systems(r, "component.code")),

    chk("Observation", "Observation.component.value[x]", "",
        component_value_present),

    chk("Observation", "Observation.effective[x]", "",
        lambda r: poly_present(r, "effectiveDateTime", "effectivePeriod",
                               "effectiveInstant", "effectiveTiming")),

    chk("Observation", "Observation.status",
        "http://hl7.org/fhir/us/qicore/ValueSet-qicore-non-negative-observation-status.html",
        lambda r: field_present(r, "status"),
        is_status=True, status_path="status"),

    chk("Observation", "Observation.subject", "",
        lambda r: field_present(r, "subject")),

    chk("Observation", "Observation.value[x]", "",
        lambda r: poly_present(
            r, "valueQuantity", "valueCodeableConcept", "valueString",
            "valueBoolean", "valueInteger", "valueRange", "valueRatio",
            "valueSampledData", "valueTime", "valueDateTime", "valuePeriod",
        )),

    # ── Patient ───────────────────────────────────────────────────────────────
    chk("Patient", "Patient.birthDate", "",
        lambda r: field_present(r, "birthDate")),

    chk("Patient", "Patient.extension (race)",
        "http://hl7.org/fhir/us/core/StructureDefinition/us-core-race",
        lambda r: extension_present(
            r, "http://hl7.org/fhir/us/core/StructureDefinition/us-core-race"),
        extract_systems_fn=extract_race_systems),

    chk("Patient", "Patient.extension (sex at birth)",
        "http://hl7.org/fhir/us/core/StructureDefinition/us-core-birthsex",
        lambda r: extension_present(
            r, "http://hl7.org/fhir/us/core/StructureDefinition/us-core-birthsex")),
    # sex at birth is a valueCode (plain string), no coding system to extract

    chk("Patient", "Patient.extension:ethnicity",
        "https://hl7.org/fhir/us/core/STU6.1/ValueSet-omb-ethnicity-category.html",
        lambda r: extension_present(
            r, "http://hl7.org/fhir/us/core/StructureDefinition/us-core-ethnicity"),
        extract_systems_fn=extract_ethnicity_systems),

    chk("Patient", "Patient.identifier", "",
        lambda r: field_present(r, "identifier")),

    chk("Patient", "Patient.name", "",
        lambda r: field_present(r, "name")),

    chk("Patient", "Patient.name.family", "",
        lambda r: field_present(r, "name.family")),

    chk("Patient", "Patient.name.given", "",
        lambda r: field_present(r, "name.given")),

    # ── Specimen ──────────────────────────────────────────────────────────────
    chk("Specimen", "Specimen.collection", "",
        lambda r: field_present(r, "collection")),

    chk("Specimen", "Specimen.collection.bodySite",
        "http://hl7.org/fhir/valueset-body-site.html",
        lambda r: field_present(r, "collection.bodySite"),
        extract_systems_fn=lambda r: cc_systems(r, "collection.bodySite")),

    chk("Specimen", "Specimen.collection.collected[x]", "",
        collected_x_present),

    chk("Specimen", "Specimen.type",
        "https://vsac.nlm.nih.gov/valueset/2.16.840.1.113762.1.4.1099.54/expansion",
        lambda r: field_present(r, "type"),
        extract_systems_fn=lambda r: cc_systems(r, "type")),
]


# ── File loading ─────────────────────────────────────────────────────────────

def find_ndjson_files(root: Path) -> list:
    return sorted(root.rglob("*.ndjson"))


def load_resources(files: list) -> tuple:
    """
    Parse every .ndjson file and group resources by resourceType.
    Returns (resources_by_type, total_lines, parse_error_count).
    """
    resources = defaultdict(list)
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


# ── Profile collection ────────────────────────────────────────────────────────

def collect_profile_counts(resources: dict) -> dict:
    """
    For each resource type, count how many resources declare each
    meta.profile URL.  Returns {resource_type: {profile_url: count}}.
    """
    out = {}
    for rtype, rlist in resources.items():
        counts = defaultdict(int)
        for r in rlist:
            meta = r.get("meta")
            if isinstance(meta, dict):
                for p in meta.get("profile", []):
                    if isinstance(p, str):
                        counts[p] += 1
        out[rtype] = dict(counts)
    return out


def profile_strings(profile_counts: dict) -> tuple:
    """
    Format profile counts as two pipe-delimited strings:
      ("ProfileURL1 | ProfileURL2 ...", "count1 | count2 ...")
    Sorted by count descending.
    """
    if not profile_counts:
        return "", ""
    pairs = sorted(profile_counts.items(), key=lambda x: -x[1])
    return (
        " | ".join(p for p, _ in pairs),
        " | ".join(str(c) for _, c in pairs),
    )


# ── Analysis ──────────────────────────────────────────────────────────────────

def analyze(resources: dict) -> list:
    """
    Run every check against the loaded resources.
    Returns a list of result dicts, one per check.
    """
    results = []
    for check in CHECKS:
        rtype = check["resource_type"]
        rlist = resources.get(rtype, [])
        total = len(rlist)

        if total == 0:
            results.append({
                **check,
                "total": 0, "missing": 0, "present": "N/A",
                "system_counts": {}, "status_counts": {}, "n_standard": 0,
            })
            continue

        missing    = 0
        sys_counts = defaultdict(int)
        sta_counts = defaultdict(int)
        n_standard = 0

        for r in rlist:
            if not check["check_fn"](r):
                missing += 1

            # Extract code systems from the coded field (CodeableConcept.coding[].system)
            if check["extract_systems_fn"] and check["value_set"]:
                systems = check["extract_systems_fn"](r)
                for s in systems:
                    sys_counts[s] += 1
                # Count this resource toward n_standard if it uses at least
                # one of RxNorm / SNOMED CT / LOINC in this field
                if any(s in STANDARD_SYSTEMS for s in systems):
                    n_standard += 1

            # Collect raw status string values for status fields
            if check["is_status"] and check["status_path"]:
                val = navigate(r, check["status_path"])
                if isinstance(val, str) and val:
                    sta_counts[val] += 1

        results.append({
            **check,
            "total":         total,
            "missing":       missing,
            "present":       "Yes" if missing == 0 else "No",
            "system_counts": dict(sys_counts),
            "status_counts": dict(sta_counts),
            "n_standard":    n_standard,
        })

    return results


# ── Row expansion ─────────────────────────────────────────────────────────────

def expand_result(result: dict, prof_used: str, prof_counts: str) -> list:
    """
    Expand one analysis result into one or more Excel data rows:
      • Coded (CC) elements  → one row per distinct code system found
      • Status elements      → one row per distinct status value found
      • Everything else      → single row
    All metrics (total, missing, present, N_std, proportions) are repeated
    on every expanded row so the reader can filter/sort freely.
    """
    total   = result["total"]
    has_vs  = bool(result["value_set"])
    is_stat = result["is_status"]
    has_ext = result["extract_systems_fn"] is not None

    base = {
        "resource_type":    result["resource_type"],
        "element":          result["element"],
        "value_set":        result["value_set"],
        "total":            total,
        "missing":          result["missing"],
        "present":          result["present"],
        "profiles_used":    prof_used,
        "profiles_count":   prof_counts,
    }

    # ── Standard-system metrics ──────────────────────────────────────────────
    if has_vs and not is_stat and has_ext:
        n   = result["n_standard"]
        p   = round(n / total, 4) if total else 0.0
        std = {"n_standard": n, "proportion": p, "opposite": round(1.0 - p, 4)}
    elif has_vs and is_stat:
        # Status fields are plain FHIR codes, never use RxNorm/SNOMED/LOINC
        std = {"n_standard": "N/A", "proportion": "N/A", "opposite": "N/A"}
    else:
        std = {"n_standard": None, "proportion": None, "opposite": None}

    rows = []

    if is_stat and result["status_counts"]:
        # One row per distinct status value, sorted alphabetically
        for sv, sc in sorted(result["status_counts"].items()):
            rows.append({**base, **std,
                         "code_system": None, "cs_count": None,
                         "status_value": sv,   "status_count": sc})

    elif has_vs and has_ext and result["system_counts"]:
        # One row per distinct code system, sorted by count descending
        for sys, cnt in sorted(result["system_counts"].items(), key=lambda x: -x[1]):
            rows.append({**base, **std,
                         "code_system": sys,  "cs_count": cnt,
                         "status_value": None, "status_count": None})

    else:
        # Single row — no code-system or status detail available
        rows.append({**base, **std,
                     "code_system": None, "cs_count": None,
                     "status_value": None, "status_count": None})

    return rows


# ── Colour / style constants ──────────────────────────────────────────────────

_C = {
    "hdr_bg":     "1F4E79",
    "hdr_fg":     "FFFFFF",
    "stripe_a":   "D6E4F0",
    "stripe_b":   "FFFFFF",
    "yes_bg":     "C6EFCE",  "yes_fg":  "276221",
    "no_bg":      "FFC7CE",  "no_fg":   "9C0006",
    "na_bg":      "FFEB9C",  "na_fg":   "7D5A00",
    "std_bg":     "E2EFDA",  # highlight for standard-system cells
    "summary_hdr":"2E75B6",
}

DETAIL_HEADERS = [
    "FHIR Resource\n(QI Core 6.0.0)",       # A
    "FHIR Data Element",                     # B
    "FHIR Value Set",                        # C
    "Total\nResources",                      # D
    "Count\nMissing",                        # E
    "Present",                               # F
    "N Using\nRxNorm / SNOMED\n/ LOINC",     # G
    "Proportion Using\nRxNorm / SNOMED\n/ LOINC", # H
    "Opposite\nProportion",                  # I
    "Code System Used",                      # J
    "Use Count",                             # K
    "Status Used",                           # L
    "Status Use Count",                      # M
    "Profiles Used",                         # N
    "Profiles Use Count",                    # O
]

DETAIL_COL_WIDTHS = [28, 50, 62, 12, 12, 10, 14, 14, 12, 55, 10, 18, 14, 62, 14]


# ── Excel writer ──────────────────────────────────────────────────────────────

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color=None, size=11) -> Font:
    kw = {"bold": bold, "size": size}
    if color:
        kw["color"] = color
    return Font(**kw)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
PCTFMT = "0.00%"
NUMFMT = "#,##0"


def write_detail_sheet(ws, all_rows: list) -> None:
    """Write the main per-element analysis sheet."""
    ws.row_dimensions[1].height = 48

    # Header
    for ci, h in enumerate(DETAIL_HEADERS, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font      = _font(bold=True, color=_C["hdr_fg"])
        cell.fill      = _fill(_C["hdr_bg"])
        cell.alignment = CENTER

    prev_rtype   = None
    stripe_color = _C["stripe_a"]

    for ri, row in enumerate(all_rows, 2):
        ws.row_dimensions[ri].height = 16

        # Alternate stripe per resource-type block
        if row["resource_type"] != prev_rtype:
            stripe_color = _C["stripe_a"] if stripe_color == _C["stripe_b"] else _C["stripe_b"]
            prev_rtype = row["resource_type"]

        stripe = _fill(stripe_color)

        def put(col, val, align=LEFT, fmt=None):
            cell = ws.cell(row=ri, column=col, value=val)
            cell.fill      = stripe
            cell.alignment = align
            cell.font      = _font()
            if fmt:
                cell.number_format = fmt
            return cell

        put(1,  row["resource_type"])
        put(2,  row["element"])
        put(3,  row["value_set"])
        put(4,  row["total"],   CENTER, NUMFMT)
        put(5,  row["missing"], CENTER, NUMFMT)

        # Present — colour-coded
        pval = row["present"]
        pc = ws.cell(row=ri, column=6, value=pval)
        pc.alignment = CENTER
        if pval == "Yes":
            pc.fill = _fill(_C["yes_bg"]); pc.font = _font(bold=True, color=_C["yes_fg"])
        elif pval == "No":
            pc.fill = _fill(_C["no_bg"]);  pc.font = _font(bold=True, color=_C["no_fg"])
        else:
            pc.fill = _fill(_C["na_bg"]);  pc.font = _font(bold=True, color=_C["na_fg"])

        # Standard-system columns (G, H, I)
        for col, key, fmt in [
            (7,  "n_standard",  NUMFMT),
            (8,  "proportion",  PCTFMT),
            (9,  "opposite",    PCTFMT),
        ]:
            val = row.get(key)
            cell = ws.cell(row=ri, column=col, value=val)
            cell.alignment = CENTER
            if val == "N/A":
                cell.fill = _fill(_C["na_bg"])
                cell.font = _font(color=_C["na_fg"])
            elif val is not None:
                cell.fill = _fill(_C["std_bg"])
                cell.font = _font()
                if fmt:
                    cell.number_format = fmt
            else:
                cell.fill = stripe
                cell.font = _font()

        # Code system (J, K)
        put(10, row.get("code_system"), LEFT)
        put(11, row.get("cs_count"),    CENTER, NUMFMT)

        # Status (L, M)
        put(12, row.get("status_value"), LEFT)
        put(13, row.get("status_count"), CENTER, NUMFMT)

        # Profiles (N, O)
        put(14, row.get("profiles_used"),  LEFT)
        put(15, row.get("profiles_count"), LEFT)

    # Column widths & freeze
    for ci, w in enumerate(DETAIL_COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A2"


def write_summary_sheet(ws, results: list, all_resource_counts: dict) -> None:
    """Write the per-resource-type summary sheet."""
    SUMM_HDR = [
        "Resource Type", "Total Resources", "Elements Checked",
        "Fully Present (Yes)", "Has Gaps (No)", "No Data (N/A)",
        "Distinct Code Systems", "Distinct Profiles",
    ]
    ws.row_dimensions[1].height = 28
    for ci, h in enumerate(SUMM_HDR, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font      = _font(bold=True, color=_C["hdr_fg"])
        cell.fill      = _fill(_C["summary_hdr"])
        cell.alignment = CENTER

    # Group results by resource type
    by_type = defaultdict(list)
    for r in results:
        by_type[r["resource_type"]].append(r)

    for ri, rtype in enumerate(sorted(by_type), 2):
        rrows = by_type[rtype]
        n_yes = sum(1 for r in rrows if r["present"] == "Yes")
        n_no  = sum(1 for r in rrows if r["present"] == "No")
        n_na  = sum(1 for r in rrows if r["present"] == "N/A")

        all_systems = set()
        for r in rrows:
            all_systems.update(r.get("system_counts", {}).keys())

        stripe = _fill(_C["stripe_a"]) if ri % 2 == 0 else _fill(_C["stripe_b"])
        vals = [
            rtype,
            all_resource_counts.get(rtype, 0),
            len(rrows),
            n_yes, n_no, n_na,
            len(all_systems),
            len(all_resource_counts.get("__profiles__" + rtype, {})),
        ]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.fill      = stripe
            cell.alignment = CENTER if ci > 1 else LEFT
            cell.font      = _font()

    for ci, w in enumerate([28, 16, 18, 18, 14, 14, 22, 18], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A2"


# ── Entry point ───────────────────────────────────────────────────────────────

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
    if not root.exists() or not root.is_dir():
        print(f"ERROR: Not a valid directory: {root}")
        sys.exit(1)

    # ── Load ──────────────────────────────────────────────────────────────────
    print(f"Scanning:  {root.resolve()}")
    files = find_ndjson_files(root)
    print(f"Found {len(files)} .ndjson file(s)")
    if not files:
        print("No .ndjson files found. Nothing to do.")
        sys.exit(0)

    print("Loading resources…")
    resources, total_lines, parse_errors = load_resources(files)
    print(f"  {total_lines:,} lines parsed  |  {parse_errors:,} parse errors")

    relevant_types = {c["resource_type"] for c in CHECKS}
    print("\nRelevant resource counts:")
    for rtype in sorted(relevant_types):
        print(f"  {rtype:<35} {len(resources.get(rtype, [])):>8,}")
    other = sorted(set(resources) - relevant_types)
    if other:
        print(f"\nOther types found (not analysed): {', '.join(other)}")

    # ── Analyse ───────────────────────────────────────────────────────────────
    print("\nCollecting profile metadata…")
    profile_counts = collect_profile_counts(resources)

    print("Running element checks…")
    results = analyze(resources)

    # ── Expand to Excel rows ──────────────────────────────────────────────────
    all_rows = []
    for result in results:
        rtype           = result["resource_type"]
        prof_used, prof_cnt = profile_strings(profile_counts.get(rtype, {}))
        all_rows.extend(expand_result(result, prof_used, prof_cnt))

    # ── Write workbook ────────────────────────────────────────────────────────
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()

    # Summary sheet first (index 0)
    ws_summary = wb.active
    ws_summary.title = "Summary"
    # Pass profile counts so summary can report distinct profile count
    resource_counts_for_summary = {k: len(v) for k, v in resources.items()}
    for rtype, pc in profile_counts.items():
        resource_counts_for_summary["__profiles__" + rtype] = pc
    write_summary_sheet(ws_summary, results, resource_counts_for_summary)

    # Detail sheet
    ws_detail = wb.create_sheet("FHIR Element Analysis")
    write_detail_sheet(ws_detail, all_rows)

    wb.save(output_path)

    # ── Console summary ───────────────────────────────────────────────────────
    n_yes = sum(1 for r in results if r["present"] == "Yes")
    n_no  = sum(1 for r in results if r["present"] == "No")
    n_na  = sum(1 for r in results if r["present"] == "N/A")
    total_sys = sum(len(r["system_counts"]) for r in results)

    print(
        f"\nResults:  {n_yes} fully present  |  {n_no} have gaps  |  {n_na} N/A\n"
        f"Code systems found across all elements: {total_sys} distinct entries\n"
        f"Report:   {output_path.resolve()}"
    )


if __name__ == "__main__":
    main()
