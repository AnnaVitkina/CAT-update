"""
End-to-end **data** pipeline: transform inputs → merge (update.py) → impact report → copy Excel to ``output/`` → cleaning.

**Paths:** Either set **three** globals ``INPUT_STORAGE``, ``PROCESSING_STORAGE``, ``OUTPUT_STORAGE``.
``INPUT_STORAGE`` must be the folder that **directly contains** the ``update/`` and ``rate/`` subfolders
(CSVs live under ``…/update``, Excel rate templates under ``…/rate``). ``PROCESSING_STORAGE`` and
``OUTPUT_STORAGE`` are the working and final-output trees. **Alternatively**, use one workspace folder with
``--root`` / ``HARDCODED_DATA_ROOT`` (folder containing ``input/``, ``processing/``, ``output/``).

Steps:
1. Choose ``update/*.csv`` and ``rate/*.xlsx`` (same prompts as ``update.py``) — in single-tree mode these live under ``input/``; in split mode they live directly under ``INPUT_STORAGE`` in those subfolders.
2. Convert CSV + Excel → JSON under ``processing/update`` and ``processing/rate``.
3. Run ``update.main()`` merge + result Excel.
4. Run ``update_impact_report`` on the combined JSON + baseline rate JSON.
5. Copy merged result ``*.xlsx`` into ``output/`` (impact report already writes there).
6. Run ``cleaning.clean_processing`` only (``input/`` is not touched).

**Colab / Jupyter:** ``exec(open(...))`` does not define ``__file__``. Use one of:

- **Recommended:** ``os.environ["CAT_UPDATE_SCRIPT_DIR"] = "/content/CAT-update"`` (repo folder with ``transform_inputs.py``) before loading the file.
- ``os.chdir("/content/CAT-update")`` before ``exec`` so ``cwd`` resolves imports.
- Pass a globals dict: ``exec(open(..., encoding="utf-8").read(), {"__file__": "/content/CAT-update/pipeline.py", "__name__": "__main__"})``
- Or ``%run /content/CAT-update/pipeline.py``.

CLI parsing uses ``parse_known_args()`` so Jupyter kernel flags (e.g. ``-f …/kernel.json``) are ignored.
"""

from __future__ import annotations

import contextlib
import json
import os
import shutil
import sys
from argparse import ArgumentParser
from pathlib import Path


def _resolve_script_dir() -> Path:
    """Folder containing ``pipeline.py`` (for imports). Works when ``__file__`` is missing (e.g. exec)."""
    env = os.environ.get("CAT_UPDATE_SCRIPT_DIR")
    if env:
        return Path(env).expanduser().resolve()
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd().resolve()


SCRIPT_DIR = _resolve_script_dir()

# ---------------------------------------------------------------------------
# **Split storage** — set all three to use different folders (e.g. separate Google Drive dirs).
# ``--root`` / ``HARDCODED_DATA_ROOT`` are then ignored.
#
#   INPUT_STORAGE     → **parent** of ``update/`` and ``rate/`` (not inside them): CSVs in ``…/update``, xlsx in ``…/rate``
#   PROCESSING_STORAGE → directory that contains ``rate/``, ``update/``, ``update_to_perform/``, ``result/``
#   OUTPUT_STORAGE    → final Excel outputs (impact report + merged workbook copy)
#
# Examples:
#   INPUT_STORAGE = Path("/content/drive/ShareA/input")
#   PROCESSING_STORAGE = Path("/content/drive/ShareB/processing")
#   OUTPUT_STORAGE = Path("/content/drive/ShareC/output")
# ---------------------------------------------------------------------------
INPUT_STORAGE = Path("/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT_CAT_update/input")
PROCESSING_STORAGE = Path("/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT_CAT_update/processing")
OUTPUT_STORAGE = Path("/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT_CAT_update/output")

# ---------------------------------------------------------------------------
# **Single-tree mode** (only when the three paths above are all ``None``):
# workspace folder with ``input/``, ``processing/``, ``output/``, or ``…/input`` only.
# ---------------------------------------------------------------------------
HARDCODED_DATA_ROOT: Path | str | None = None

# When True, after CSV/template selection the pipeline hides merge/report/clean chatter and
# prints only ``Done`` on success. Interactive pick lists still print. Override with
# ``python pipeline.py --verbose`` or set ``PIPELINE_QUIET = False`` below.
PIPELINE_QUIET = True


@contextlib.contextmanager
def _silence_stdio():
    """Redirect stdout and stderr to os.devnull (for update / impact / transform subprocess-style noise)."""
    with open(os.devnull, "w", encoding="utf-8") as devnull:
        old_out, old_err = sys.stdout, sys.stderr
        try:
            sys.stdout = devnull
            sys.stderr = devnull
            yield
        finally:
            sys.stdout, sys.stderr = old_out, old_err


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


def _repo_root_candidates() -> list[Path]:
    """Ordered locations that might contain ``transform_inputs.py`` (``exec`` has no ``__file__``)."""
    out: list[Path] = []
    env = os.environ.get("CAT_UPDATE_SCRIPT_DIR")
    if env:
        out.append(Path(env).expanduser().resolve())
    out.append(SCRIPT_DIR)
    out.append(Path.cwd().resolve())
    # De-dupe, keep order
    seen: set[str] = set()
    unique: list[Path] = []
    for p in out:
        k = str(p)
        if k in seen:
            continue
        seen.add(k)
        unique.append(p)
    return unique


def _ensure_on_path() -> None:
    for root in _repo_root_candidates():
        if (root / "transform_inputs.py").is_file():
            s = str(root)
            if s not in sys.path:
                sys.path.insert(0, s)
            return
    raise RuntimeError(
        "Cannot find transform_inputs.py (CAT-update code folder). "
        "Before running: os.chdir('/content/CAT-update') to the repo, or set "
        "os.environ['CAT_UPDATE_SCRIPT_DIR'] to that folder (the directory that contains "
        "transform_inputs.py and pipeline.py)."
    )


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


def apply_split_paths(
    input_storage: Path,
    processing_storage: Path,
    output_storage: Path,
) -> None:
    """
    Three independent roots (input / processing / output trees).

    ``input_storage`` is the directory that contains ``update/`` and ``rate/`` subfolders
    (same layout as ``…/input`` in single-tree mode, but without requiring the folder to be named ``input``).
    """
    import transform_inputs as ti

    inp = input_storage.resolve(strict=False)
    proc = processing_storage.resolve(strict=False)

    ti.ROOT = inp
    ti.INPUT_UPDATE = inp / "update"
    ti.INPUT_RATE = inp / "rate"
    ti.OUT_RATE = proc / "rate"
    ti.OUT_UPDATE = proc / "update"


def patch_update_list_rate_jsons_for_processing(up, processing_storage: Path) -> None:
    """Route ``list_rate_jsons`` to ``processing_storage/rate`` when roots differ."""
    proc = processing_storage.resolve(strict=False)

    def list_rate_jsons_split() -> list[Path]:
        d = proc / "rate"
        if not d.is_dir():
            return []
        return sorted(p for p in d.glob("*.json") if p.is_file())

    up.list_rate_jsons = list_rate_jsons_split


def _print_no_csv_help(
    iu: Path,
    ir: Path,
    script_dir: Path,
    *,
    hint_repo_workspace: bool,
) -> None:
    same_as_repo = hint_repo_workspace
    print(
        f"No Ocean Rates *.csv found in:\n  {iu}\n",
        end="",
    )
    if same_as_repo:
        print(
            "\nThe data root is the **code repo**; by default it has no CSVs. Either:\n"
            "  • Copy your update CSV into the folder above, or\n"
            "  • Point the pipeline at your **data** folder (Colab: usually Drive), e.g. in this file set:\n"
            "      HARDCODED_DATA_ROOT = Path('/content/drive/MyDrive/CAT test/input')\n"
            "    or run:  python pipeline.py --root \"/content/drive/MyDrive/CAT test/input\"\n"
        )
    else:
        print(
            "\nAdd *.csv there, or fix --root / HARDCODED_DATA_ROOT so it points to the folder "
            "that contains your ``input/update/`` tree.\n"
        )
    print(f"Rate card Excel is read from:\n  {ir}\n")


def _print_no_rate_xlsx_help(
    ir: Path,
    script_dir: Path,
    *,
    hint_repo_workspace: bool,
) -> None:
    same_as_repo = hint_repo_workspace
    print(
        f"No rate workbook *.xlsx found in:\n  {ir}\n",
        end="",
    )
    if same_as_repo:
        print(
            "\nAdd your baseline Rate Card Excel there, or set **HARDCODED_DATA_ROOT** / ``--root`` "
            "to the folder where ``input/rate`` lives (often Drive on Colab).\n"
        )
    else:
        print("\nAdd *.xlsx there or fix your data root.\n")


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


def ensure_layout_split(
    input_storage: Path,
    processing_storage: Path,
    output_storage: Path,
) -> None:
    inp = input_storage.resolve(strict=False)
    proc = processing_storage.resolve(strict=False)
    out = output_storage.resolve(strict=False)
    for sub in ("update", "rate"):
        (inp / sub).mkdir(parents=True, exist_ok=True)
    for sub in ("rate", "update", "update_to_perform", "result"):
        (proc / sub).mkdir(parents=True, exist_ok=True)
    out.mkdir(parents=True, exist_ok=True)


def run_transform_single(
    csv_path: Path, template_xlsx: Path, *, quiet: bool = False
) -> tuple[Path, Path]:
    """Write processing JSON for the chosen CSV and rate workbook; return (update_json, rate_json)."""
    from transform_inputs import OUT_RATE, OUT_UPDATE, csv_to_json, safe_json_name, workbook_to_json

    csv_path = csv_path.resolve()
    template_xlsx = template_xlsx.resolve()

    def _run() -> tuple[Path, Path]:
        upd = csv_to_json(csv_path)
        up_json = OUT_UPDATE / f"{safe_json_name(csv_path.stem)}.json"
        up_json.write_text(json.dumps(upd, ensure_ascii=False, indent=2), encoding="utf-8")
        if not quiet:
            print(f"[transform] Wrote {up_json}")

        rt = workbook_to_json(template_xlsx)
        rate_json = OUT_RATE / f"{safe_json_name(template_xlsx.stem)}.json"
        rate_json.write_text(json.dumps(rt, ensure_ascii=False, indent=2), encoding="utf-8")
        if not quiet:
            print(f"[transform] Wrote {rate_json}")

        return up_json, rate_json

    if quiet:
        with _silence_stdio():
            return _run()
    return _run()


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
    ap.add_argument(
        "--verbose",
        action="store_true",
        help="Print full step-by-step progress (default is quiet: only prompts and final Done).",
    )
    # Jupyter/IPython inject e.g. ``-f …/kernel-….json`` — ignore unknown argv tails.
    args, _unknown_argv = ap.parse_known_args()

    quiet = PIPELINE_QUIET and not args.verbose

    split_set = sum(
        1 for x in (INPUT_STORAGE, PROCESSING_STORAGE, OUTPUT_STORAGE) if x is not None
    )
    if split_set not in (0, 3):
        print(
            "Set all three of INPUT_STORAGE, PROCESSING_STORAGE, OUTPUT_STORAGE in pipeline.py, "
            "or leave all three as None for single-tree mode (--root / HARDCODED_DATA_ROOT)."
        )
        raise SystemExit(1)

    use_split = split_set == 3
    data_root: Path | None = None
    clean_processing_target: Path | None = None

    if use_split:
        if (args.root is not None or HARDCODED_DATA_ROOT is not None) and not quiet:
            print("Note: --root / HARDCODED_DATA_ROOT ignored (split storage paths are set).\n")
        input_storage = Path(INPUT_STORAGE).expanduser()
        processing_storage = Path(PROCESSING_STORAGE).expanduser()
        output_storage = Path(OUTPUT_STORAGE).expanduser()
        apply_split_paths(input_storage, processing_storage, output_storage)
        ensure_layout_split(input_storage, processing_storage, output_storage)

        import update as up
        from update import result_stem_for_merge_output
        import update_impact_report as uir

        inp = input_storage.resolve(strict=False)
        proc = processing_storage.resolve(strict=False)
        out = output_storage.resolve(strict=False)

        up.INPUT_UPDATE = inp / "update"
        up.INPUT_RATE = inp / "rate"
        up.OUT_COMBINED_DIR = proc / "update_to_perform"
        up.OUT_RESULT_DIR = proc / "result"
        if inp.name.lower() == "input" and proc.name.lower() == "processing":
            up.ROOT = inp.parent
        else:
            up.ROOT = inp
        patch_update_list_rate_jsons_for_processing(up, proc)
        uir.OUTPUT_DIR_DEFAULT = out

        iu, ir = up.INPUT_UPDATE, up.INPUT_RATE
        out_dir = out
        clean_processing_target = proc

        if not quiet:
            print(
                f"\nSplit storage:\n"
                f"  input:       {inp}\n"
                f"  processing:  {proc}\n"
                f"  output:      {out}\n"
            )
    else:
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

        iu, ir = up.INPUT_UPDATE, up.INPUT_RATE
        out_dir = data_root / "output"
        clean_processing_target = None

        if not quiet:
            print(f"\nData root (workspace): {data_root}\n")

    hint_repo = (data_root is not None) and (data_root.resolve() == SCRIPT_DIR.resolve())

    if not up.list_update_csvs():
        _print_no_csv_help(iu, ir, SCRIPT_DIR, hint_repo_workspace=hint_repo)
        raise SystemExit(1)
    if not quiet:
        print("Select files (same lists as update.py).\n")

    csv_path = up.prompt_pick_csv()
    if not csv_path or not csv_path.is_file():
        print("No CSV selected.")
        raise SystemExit(1)
    csv_path = csv_path.resolve()
    if not quiet:
        print(f"Using CSV: {csv_path}")

    if not up.list_rate_templates():
        _print_no_rate_xlsx_help(ir, SCRIPT_DIR, hint_repo_workspace=hint_repo)
        raise SystemExit(1)

    tpl_path = up.prompt_pick_template_xlsx()
    if not tpl_path or not tpl_path.is_file():
        print("No rate template Excel selected.")
        raise SystemExit(1)
    tpl_path = tpl_path.resolve()
    if not quiet:
        print(f"Using rate template: {tpl_path}\n")

    _, rate_json_path = run_transform_single(csv_path, tpl_path, quiet=quiet)

    if not quiet:
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
        if quiet:
            with _silence_stdio():
                up.main()
        else:
            up.main()
    finally:
        sys.argv = old_argv

    from transform_inputs import safe_json_name

    combined_path = up.OUT_COMBINED_DIR / f"{safe_json_name(csv_path.stem)}_combined.json"
    if not combined_path.is_file():
        print(f"Warning: expected combined JSON missing: {combined_path}")

    if not quiet:
        print("\n--- update_impact_report.py ---\n")
    # out_dir was set in split branch (output_storage) or legacy branch (data_root/output).
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
        if quiet:
            with _silence_stdio():
                uir.main()
        else:
            uir.main()
    finally:
        sys.argv = old_argv

    if not quiet:
        print("\n--- copy merged Excel to output/ ---\n")
    try:
        prior = json.loads(rate_json_path.read_text(encoding="utf-8"))
        out_stem = result_stem_for_merge_output(rate_json_path, tpl_path, prior)
        merged_xlsx = up.OUT_RESULT_DIR / f"{out_stem}.xlsx"
        if merged_xlsx.is_file():
            dest = out_dir / merged_xlsx.name
            shutil.copy2(merged_xlsx, dest)
            if not quiet:
                print(f"Copied {merged_xlsx.name} → {dest}")
        else:
            print(f"Warning: merged Excel not found: {merged_xlsx}")
    except Exception as ex:
        print(f"Warning: could not copy merged Excel: {ex}")

    if not args.skip_clean:
        if not quiet:
            print("\n--- cleaning (processing only; input/ unchanged) ---\n")
        from cleaning import clean_processing, clean_processing_folder

        if quiet:
            with _silence_stdio():
                if clean_processing_target is not None:
                    errs = clean_processing_folder(clean_processing_target)
                else:
                    assert data_root is not None
                    errs = clean_processing(data_root)
        else:
            if clean_processing_target is not None:
                errs = clean_processing_folder(clean_processing_target)
            else:
                assert data_root is not None
                errs = clean_processing(data_root)
        if not quiet:
            for line in errs:
                print(line)
            if errs:
                print("(Cleaning reported errors; output/ was not cleaned.)")
            else:
                print("Cleaning finished.")
    elif not quiet:
        print("\nSkipping cleaning (--skip-clean).")

    if quiet:
        print("Done", flush=True)
    else:
        print(f"\nDone. Excel artifacts under: {out_dir}\n")


if __name__ == "__main__":
    main()
