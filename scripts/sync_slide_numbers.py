#!/usr/bin/env python3
"""Renumber slide files to reserve slots for additions or close gaps after deletions.

This script only renames/deletes slide files. It does not read the source PPTX and
does not modify file contents. Run it before `generate_slides.py` whenever the user
adds or deletes slides.

Examples:
    # Reserve slots 3 and 6 for new slides (dry run)
    uv run python sync_slide_numbers.py --project-dir my-deck --add 3,6

    # Actually apply the renames
    uv run python sync_slide_numbers.py --project-dir my-deck --add 3,6 --apply

    # Delete the slides originally at positions 2 and 5
    uv run python sync_slide_numbers.py --project-dir my-deck --delete 2,5 --apply

Workflow after source PPTX changes:
    1. uv run python sync_slide_numbers.py --project-dir my-deck --add 3,6 --apply
    2. uv run python generate_slides.py --target "<target.pptx>" \
           --project-dir my-deck --slides 3,6
    3. uv run python my-deck/build_deck.py --target "<target.pptx>"
"""

import argparse
import re
import sys
import uuid
from pathlib import Path


def parse_positions(arg: str) -> list[int]:
    """Parse '3,6' or '2-5' into a sorted list of unique 1-based positions."""
    if not arg:
        return []
    result = set()
    for part in arg.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            result.update(range(int(start), int(end) + 1))
        else:
            result.add(int(part))
    return sorted(result)


def list_existing_slides(project_dir: Path) -> list[tuple[int, Path]]:
    """Return [(index, path), ...] for existing s*.py slide files, sorted by index."""
    slides_dir = project_dir / "slides"
    if not slides_dir.exists():
        raise FileNotFoundError(f"Slides directory not found: {slides_dir}")
    files = []
    for path in slides_dir.glob("s*.py"):
        m = re.match(r"s(\d+)_.*\.py$", path.name)
        if not m:
            continue
        files.append((int(m.group(1)), path))
    return sorted(files, key=lambda x: x[0])


def compute_plan(
    existing: list[tuple[int, Path]],
    add_positions: list[int],
    delete_positions: list[int],
) -> dict:
    """Return rename/delete plan.

    Returns {
        "renames": [(old_path, new_path), ...],
        "deletes": [path, ...],
        "reserved": [int, ...],
        "total_before": int,
        "total_after": int,
    }
    """
    n = len(existing)
    add_set = set(add_positions)
    delete_set = set(delete_positions)

    if any(p <= 0 for p in add_set | delete_set):
        raise ValueError("Positions must be positive integers.")
    if len(add_set) != len(add_positions):
        raise ValueError("Duplicate positions in --add.")
    if len(delete_set) != len(delete_positions):
        raise ValueError("Duplicate positions in --delete.")
    if add_set & delete_set:
        raise ValueError("Same position cannot be in both --add and --delete.")
    if any(p > n for p in delete_set):
        raise ValueError(f"Delete position out of range (max {n}).")

    final_count = n - len(delete_set) + len(add_set)
    if any(p > final_count for p in add_set):
        raise ValueError(f"Add position out of range (max {final_count}).")

    old_indices = [idx for idx, _ in existing]
    remaining = [idx for idx in old_indices if idx not in delete_set]

    # Fill final slots 1..final_count, reserving add positions.
    reserved = sorted(add_set)
    mapping = {}
    rem_iter = iter(remaining)
    for slot in range(1, final_count + 1):
        if slot in add_set:
            continue
        try:
            old_idx = next(rem_iter)
            mapping[old_idx] = slot
        except StopIteration:
            # Should not happen if counts are correct.
            raise RuntimeError("Ran out of slides while filling slots.")

    renames = []
    deletes = []
    index_to_path = {idx: path for idx, path in existing}
    for idx, path in existing:
        if idx in delete_set:
            deletes.append(path)
        elif idx in mapping:
            new_idx = mapping[idx]
            if new_idx != idx:
                # Keep the original title stem; only the index changes.
                stem = path.stem.split("_", 1)[1] if "_" in path.stem else path.stem
                new_name = f"s{new_idx:02d}_{stem}.py"
                renames.append((path, path.parent / new_name))

    return {
        "renames": renames,
        "deletes": deletes,
        "reserved": reserved,
        "total_before": n,
        "total_after": final_count,
    }


def apply_plan(plan: dict, dry_run: bool) -> None:
    """Execute renames via temp names to avoid collisions, then delete removed slides."""
    renames = plan["renames"]
    deletes = plan["deletes"]

    if dry_run:
        print("Dry run — no files changed.")
        for old_path, new_path in renames:
            print(f"  rename {old_path.name} -> {new_path.name}")
        for path in deletes:
            print(f"  delete {path.name}")
        for slot in plan["reserved"]:
            print(f"  reserve slot {slot:02d}")
        return

    # Two-pass rename to avoid collisions.
    temp_paths = []
    for old_path, _ in renames:
        temp_path = old_path.with_suffix(f".tmp-{uuid.uuid4().hex[:8]}")
        old_path.rename(temp_path)
        temp_paths.append(temp_path)

    for temp_path, (_, new_path) in zip(temp_paths, renames):
        temp_path.rename(new_path)

    for path in deletes:
        path.unlink()

    print(f"Renamed {len(renames)} file(s), deleted {len(deletes)} file(s).")
    for slot in plan["reserved"]:
        print(f"Reserved slot {slot:02d} for generation.")


def main():
    parser = argparse.ArgumentParser(
        description="Renumber slide files after additions or deletions."
    )
    parser.add_argument("--project-dir", required=True, help="Project directory")
    parser.add_argument("--add", default="", help="Comma list of slots to reserve, e.g. 3,6")
    parser.add_argument(
        "--delete", default="", help="Comma list of original slots to remove, e.g. 2,5"
    )
    parser.add_argument(
        "--apply", action="store_true", help="Apply renames/deletes (default is dry run)"
    )
    args = parser.parse_args()

    project_dir = Path(args.project_dir)
    add_positions = parse_positions(args.add)
    delete_positions = parse_positions(args.delete)

    if not add_positions and not delete_positions:
        print("Nothing to do. Use --add and/or --delete.")
        sys.exit(0)

    try:
        existing = list_existing_slides(project_dir)
    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)

    try:
        plan = compute_plan(existing, add_positions, delete_positions)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    print(
        f"Slide count: {plan['total_before']} -> {plan['total_after']} "
        f"(+{len(add_positions)} -{len(delete_positions)})"
    )

    apply_plan(plan, dry_run=not args.apply)


if __name__ == "__main__":
    main()
