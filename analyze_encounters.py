import json
import sys
from collections import Counter
from pathlib import Path

SNOMED_SYSTEMS = {
    "http://snomed.info/sct",
    "https://snomed.info/sct",
    "urn:oid:2.16.840.1.113883.6.96",
}


def is_snomed(system: str) -> bool:
    return (system or "").lower().rstrip("/") in {s.rstrip("/") for s in SNOMED_SYSTEMS}


def extract_encounter_data(bundle: dict) -> tuple[list[str], list[dict]]:
    """Returns (type_labels, snomed_codings) for all Encounters in the bundle."""
    type_labels = []
    snomed_codings = []

    for entry in bundle.get("entry", []):
        resource = entry.get("resource", {})
        if resource.get("resourceType") != "Encounter":
            continue

        type_list = resource.get("type", [])
        if not type_list:
            type_labels.append("(no type)")
            continue

        for t in type_list:
            codings = t.get("coding", [])
            if codings:
                for coding in codings:
                    display = coding.get("display") or coding.get("code") or "(unknown)"
                    type_labels.append(display)
                    if is_snomed(coding.get("system", "")):
                        snomed_codings.append({
                            "code": coding.get("code", "(no code)"),
                            "display": coding.get("display", "(no display)"),
                            "system": coding.get("system", ""),
                        })
            elif t.get("text"):
                type_labels.append(t["text"])
            else:
                type_labels.append("(no type)")

    return type_labels, snomed_codings


def main(folder: str):
    root = Path(folder)
    if not root.exists():
        print(f"Error: folder '{folder}' does not exist.")
        sys.exit(1)

    type_counts = Counter()
    snomed_counts = Counter()  # keyed by (code, display, system)
    files_scanned = 0
    encounters_found = 0
    errors = []

    for json_file in root.rglob("*.json"):
        try:
            with open(json_file, encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            errors.append(f"{json_file}: {e}")
            continue

        files_scanned += 1

        bundles = []
        if isinstance(data, list):
            bundles = [b for b in data if isinstance(b, dict) and b.get("resourceType") == "Bundle"]
        elif isinstance(data, dict) and data.get("resourceType") == "Bundle":
            bundles = [data]

        for bundle in bundles:
            labels, snomed = extract_encounter_data(bundle)
            type_counts.update(labels)
            encounters_found += len(labels)
            for s in snomed:
                snomed_counts[(s["code"], s["display"], s["system"])] += 1

    # --- Encounter type counts ---
    print(f"\nFiles scanned   : {files_scanned}")
    print(f"Encounters found: {encounters_found}")
    print(f"\n{'Encounter Type':<60} {'Count':>6}")
    print("-" * 68)
    for enc_type, count in type_counts.most_common():
        print(f"{enc_type:<60} {count:>6}")

    # --- SNOMED codes ---
    print(f"\n{'SNOMED Codes Found':=^80}")
    if snomed_counts:
        print(f"\n{'Code':<15} {'Count':>6}  {'Display':<40} System")
        print("-" * 100)
        for (code, display, system), count in sorted(snomed_counts.items(), key=lambda x: -x[1]):
            print(f"{code:<15} {count:>6}  {display:<40} {system}")
    else:
        print("\n  No SNOMED codes found in any Encounter types.")

    if errors:
        print(f"\nFailed to parse {len(errors)} file(s):")
        for e in errors:
            print(f"  {e}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python analyze_encounters.py <folder>")
        sys.exit(1)
    main(sys.argv[1])
