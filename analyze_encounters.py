import json
import os
import sys
from collections import Counter
from pathlib import Path


def extract_encounter_types(bundle: dict) -> list[str]:
    types = []
    entries = bundle.get("entry", [])
    for entry in entries:
        resource = entry.get("resource", {})
        if resource.get("resourceType") != "Encounter":
            continue
        type_list = resource.get("type", [])
        if not type_list:
            types.append("(no type)")
            continue
        for t in type_list:
            codings = t.get("coding", [])
            if codings:
                for coding in codings:
                    display = coding.get("display") or coding.get("code") or "(unknown)"
                    types.append(display)
            elif t.get("text"):
                types.append(t["text"])
            else:
                types.append("(no type)")
    return types


def main(folder: str):
    root = Path(folder)
    if not root.exists():
        print(f"Error: folder '{folder}' does not exist.")
        sys.exit(1)

    counts = Counter()
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
        elif isinstance(data, dict):
            if data.get("resourceType") == "Bundle":
                bundles = [data]

        for bundle in bundles:
            types = extract_encounter_types(bundle)
            counts.update(types)
            encounters_found += len(types)

    print(f"\nFiles scanned : {files_scanned}")
    print(f"Encounters found: {encounters_found}")
    print(f"\n{'Encounter Type':<60} {'Count':>6}")
    print("-" * 68)
    for enc_type, count in counts.most_common():
        print(f"{enc_type:<60} {count:>6}")

    if errors:
        print(f"\nFailed to parse {len(errors)} file(s):")
        for e in errors:
            print(f"  {e}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python analyze_encounters.py <folder>")
        sys.exit(1)
    main(sys.argv[1])
