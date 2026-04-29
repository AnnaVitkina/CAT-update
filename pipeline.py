"""
End-to-end **data** pipeline: transform inputs → merge (update.py) → impact report → copy Excel to ``output/`` → cleaning.

**Data root**: folder that contains ``input/``, ``processing/``, and ``output/`` (or ``…/input`` —
see ``normalize_data_root``). Set ``HARDCODED_DATA_ROOT`` below to skip ``--root``, or pass
``--root`` on the command line to override it.

Steps:
1. Choose ``input/update/*.csv`` and ``input/rate/*.xlsx`` (same prompts as ``update.py``).
2. Convert CSV + Excel → JSON under ``processing/update`` and ``processing/rate``.
3. Run ``update.main()`` merge + result Excel.
4. Run ``update_impact_report`` on the combined JSON + baseline rate JSON.
5. Copy merged result ``*.xlsx`` into ``output/`` (impact report already writes there).
6. Run ``cleaning.clean_processing`` only (``input/`` is not touched).
"""

from __future__ import annotations

import json
import shutil
import sys
from argparse import ArgumentParser
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# Hardcode your data workspace here (optional).
# Same rules as ``--root``: either the parent folder of ``input/`` **or** the path ending in
# ``…/input``. Use ``None`` to rely on ``--root`` or the script directory (default).
# Examples:
#   HARDCODED_DATA_ROOT = Path("/content/drive/MyDrive/CAT test/input")
#   HARDCODED_DATA_ROOT = Path(r"C:\Users\me\Projects\CAT-data")
# ---------------------------------------------------------------------------
HARDCODED_DATA_ROOT: Path("/content/drive/MyDrive/CAT test")


def normalize_data_root(raw: Path | str | None, *, default: Path) -> Path:
    """
    Resolve the workspace folder that contains ``input/``, ``processing/``, ``output/``.

    You may pass either that folder (e.g. ``…/CAT test``) **or** the ``input`` folder itself
    (e.g. ``/content/drive/MyDrive/CAT test/input``) — if the last segment is ``input``,
    its parent is used as the data root. Spaces in path segments are fine (quote in shell).
    """
    if raw is None:
        return default.resolve()
    p = Path(raw).expanduser()
    if p.name.lower() == "input":
        p = p.parent
    try:
        return p.resolve(strict=False)
    except TypeError:
        return p.resolve()


def _ensure_on_path() -> None:
    sd = str(SCRIPT_DIR)
    if sd not in sys.path:
        sys.path.insert(0, sd)


def apply_data_root(data_root: Path) -> None:
    """
    Point ``transform_inputs`` / ``update`` / ``update_impact_report`` at ``data_root``
    for all inputs and processing paths. Call **before** importing project modules that
    depend on paths (or import them only after this runs).
    """
    data_root = data_root.resolve()
    import transform_inputs as ti

    ti.ROOT = data_root
    ti.INPUT_RATE = data_root / "input" / "rate"
    ti.INPUT_UPDATE = data_root / "input" / "update"
    ti.OUT_RATE = data_root / "processing" / "rate"
    ti.OUT_UPDATE = data_root / "processing" / "update"


def ensure_layout(data_root: Path) -> None:
    data_root = data_root.resolve()
    for rel in (
        "input/rate",
        "input/update",
        "processing/rate",
        "processing/update",
        "processing/update_to_perform",
        "processing/result",
        "output",
    ):
        (data_root.joinpath(*rel.split("/"))).mkdir(parents=True, exist_ok=True)


def run_transform_single(csv_path: Path, template_xlsx: Path) -> tuple[Path, Path]:
    """Write processing JSON for the chosen CSV and rate workbook; return (update_json, rate_json)."""
    from transform_inputs import OUT_RATE, OUT_UPDATE, csv_to_json, safe_json_name, workbook_to_json

    csv_path = csv_path.resolve()
    template_xlsx = template_xlsx.resolve()

    upd = csv_to_json(csv_path)
    up_json = OUT_UPDATE / f"{safe_json_name(csv_path.stem)}.json"
    up_json.write_text(json.dumps(upd, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[transform] Wrote {up_json}")

    rt = workbook_to_json(template_xlsx)
    rate_json = OUT_RATE / f"{safe_json_name(template_xlsx.stem)}.json"
    rate_json.write_text(json.dumps(rt, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[transform] Wrote {rate_json}")

    return up_json, rate_json


def main() -> None:
    _ensure_on_path()

    ap = ArgumentParser(
        description=(
            "Run transform_inputs → update.py → update_impact_report; copy Excel to output/; "
            "clean processing/ only (input/ preserved)."
        )
    )
    ap.add_argument(
        "--root",
        type=Path,
        default=None,
        metavar="PATH",
        help=(
            "Workspace folder containing input/, processing/, output/, OR the path to …/input "
            "(parent is used). Example: /content/drive/MyDrive/CAT test/input "
            f"(default: {SCRIPT_DIR})"
        ),
    )
    ap.add_argument(
        "--skip-clean",
        action="store_true",
        help="Do not run cleaning at the end.",
    )
    args = ap.parse_args()

    if args.root is not None:
        raw_root: Path | str | None = args.root
    elif HARDCODED_DATA_ROOT is not None:
        raw_root = HARDCODED_DATA_ROOT
    else:
        raw_root = None

    data_root = normalize_data_root(raw_root, default=SCRIPT_DIR)
    apply_data_root(data_root)
    ensure_layout(data_root)

    import update as up
    from update import result_stem_for_merge_output
    import update_impact_report as uir

    up.ROOT = data_root
    up.INPUT_UPDATE = data_root / "input" / "update"
    up.INPUT_RATE = data_root / "input" / "rate"
    up.OUT_COMBINED_DIR = data_root / "processing" / "update_to_perform"
    up.OUT_RESULT_DIR = data_root / "processing" / "result"
    uir.OUTPUT_DIR_DEFAULT = data_root / "output"

    print(f"\nData root (workspace): {data_root}\n")
    print("Select files (same lists as update.py).\n")

    csv_path = up.prompt_pick_csv()
    if not csv_path or not csv_path.is_file():
        print("No CSV selected.")
        raise SystemExit(1)
    csv_path = csv_path.resolve()
    print(f"Using CSV: {csv_path}")

    tpl_path = up.prompt_pick_template_xlsx()
    if not tpl_path or not tpl_path.is_file():
        print("No rate template Excel selected.")
        raise SystemExit(1)
    tpl_path = tpl_path.resolve()
    print(f"Using rate template: {tpl_path}\n")

    _, rate_json_path = run_transform_single(csv_path, tpl_path)

    print("\n--- update.py (merge + Excel) ---\n")
    old_argv = sys.argv[:]
    try:
        sys.argv = [
            "update.py",
            "--csv",
            str(csv_path),
            "--rate-json",
            str(rate_json_path),
            "--template-xlsx",
            str(tpl_path),
        ]
        up.main()
    finally:
        sys.argv = old_argv

    from transform_inputs import safe_json_name

    combined_path = up.OUT_COMBINED_DIR / f"{safe_json_name(csv_path.stem)}_combined.json"
    if not combined_path.is_file():
        print(f"Warning: expected combined JSON missing: {combined_path}")

    print("\n--- update_impact_report.py ---\n")
    out_dir = data_root / "output"
    old_argv = sys.argv[:]
    try:
        sys.argv = [
            "update_impact_report.py",
            "--update",
            str(combined_path),
            "--rate",
            str(rate_json_path),
            "--out-dir",
            str(out_dir),
            "--rate-card-label",
            csv_path.stem,
        ]
        uir.main()
    finally:
        sys.argv = old_argv

    print("\n--- copy merged Excel to output/ ---\n")
    try:
        prior = json.loads(rate_json_path.read_text(encoding="utf-8"))
        out_stem = result_stem_for_merge_output(rate_json_path, tpl_path, prior)
        merged_xlsx = up.OUT_RESULT_DIR / f"{out_stem}.xlsx"
        if merged_xlsx.is_file():
            dest = out_dir / merged_xlsx.name
            shutil.copy2(merged_xlsx, dest)
            print(f"Copied {merged_xlsx.name} → {dest}")
        else:
            print(f"Warning: merged Excel not found: {merged_xlsx}")
    except Exception as ex:
        print(f"Warning: could not copy merged Excel: {ex}")

    if not args.skip_clean:
        print("\n--- cleaning (processing only; input/ unchanged) ---\n")
        from cleaning import clean_processing

        errs = clean_processing(data_root)
        for line in errs:
            print(line)
        if errs:
            print("(Cleaning reported errors; output/ was not cleaned.)")
        else:
            print("Cleaning finished.")
    else:
        print("\nSkipping cleaning (--skip-clean).")

    print(f"\nDone. Excel artifacts under: {out_dir}\n")


if __name__ == "__main__":
    main()
