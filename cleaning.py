"""
Remove all files under ``input/`` and ``processing/`` immediate subfolders.
Directory structure (subfolders) is kept; only files are deleted.
"""
from __future__ import annotations

import argparse
from pathlib import Path

from transform_inputs import ROOT


def _unlink_files_under(directory: Path) -> list[str]:
    """Delete every file under ``directory`` (recursive). Directories are not removed."""
    errors: list[str] = []
    if not directory.is_dir():
        return errors
    for p in directory.rglob("*"):
        if p.is_file():
            try:
                p.unlink()
            except OSError as exc:
                errors.append(f"{p}: {exc}")
    return errors


def _clean_category_base(base: Path) -> list[str]:
    """
    Under ``base`` (e.g. ``input`` or ``processing``): delete files directly in ``base``,
    and all files inside each child subfolder (recursively). Child subfolders stay.
    """
    errors: list[str] = []
    if not base.is_dir():
        return [f"Not a directory (skip): {base}"]
    for child in base.iterdir():
        if child.is_file():
            try:
                child.unlink()
            except OSError as exc:
                errors.append(f"{child}: {exc}")
        elif child.is_dir():
            errors.extend(_unlink_files_under(child))
    return errors


def clean_input(root: Path | None = None) -> list[str]:
    """Empty all files under ``{root}/input`` (including inside ``rate``, ``update``, etc.)."""
    base = (root or ROOT) / "input"
    return _clean_category_base(base)


def clean_processing(root: Path | None = None) -> list[str]:
    """Empty all files under ``{root}/processing`` (each subfolder’s contents only)."""
    base = (root or ROOT) / "processing"
    return _clean_category_base(base)


def clean_processing_folder(processing_root: Path) -> list[str]:
    """
    Empty all files under an explicit ``processing`` directory (each child folder’s contents).
    Use when ``input/`` and ``processing/`` live under different parents (e.g. separate Drive folders).
    """
    return _clean_category_base(processing_root.resolve())


def clean_input_and_processing(root: Path | None = None) -> list[str]:
    """Run :func:`clean_input` and :func:`clean_processing`; return combined error lines."""
    r = root or ROOT
    return clean_input(r) + clean_processing(r)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Delete all files under input/ and processing/ subfolders (keep folders)."
    )
    parser.add_argument(
        "--input-only",
        action="store_true",
        help="Only clean input/*",
    )
    parser.add_argument(
        "--processing-only",
        action="store_true",
        help="Only clean processing/*",
    )
    parser.add_argument(
        "--root",
        type=Path,
        default=None,
        help=f"Project root (default: {ROOT})",
    )
    args = parser.parse_args()
    root = args.root.resolve() if args.root else ROOT

    if args.input_only and args.processing_only:
        parser.error("Use at most one of --input-only / --processing-only")

    if args.input_only:
        errs = clean_input(root)
    elif args.processing_only:
        errs = clean_processing(root)
    else:
        errs = clean_input_and_processing(root)

    for line in errs:
        print(line)
    if errs:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
