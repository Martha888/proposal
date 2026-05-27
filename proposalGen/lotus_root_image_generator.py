#!/usr/bin/env python3
"""
ProposalGen - Lotus Root Opportunity Image Generator

Generate appetizing concept images for every lotus-root opportunity and write
the generated image links back into lotus_root_opportunities.json.

This script converts opportunity rows into the manifest format consumed by
skills/ppt-master/scripts/image_gen.py, runs that existing image-generation
tool, then updates the opportunity JSON with local image paths.

Usage:
    python3 proposalGen/lotus_root_image_generator.py <lotus_root_opportunities.json>

Examples:
    python3 proposalGen/lotus_root_image_generator.py \
        projects/menu_research/费大厨_20260527_095043/lotus_root_opportunities.json

    python3 proposalGen/lotus_root_image_generator.py \
        projects/menu_research/费大厨_20260527_095043/lotus_root_opportunities.json \
        --menu-json projects/menu_research/费大厨_20260527_095043/structured_menu.json \
        --concurrency 1

Dependencies:
    Uses skills/ppt-master/scripts/image_gen.py and its configured image backend.

Environment:
    IMAGE_BACKEND plus provider-specific image keys, for example:
    IMAGE_BACKEND=openai
    OPENAI_API_KEY=...
    OPENAI_MODEL=gemini-2.5-flash-image
    OPENAI_BASE_URL=...
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Any


_REPO_ROOT = Path(__file__).resolve().parent.parent
_IMAGE_GEN = _REPO_ROOT / "skills" / "ppt-master" / "scripts" / "image_gen.py"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate lotus-root opportunity dish images.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "opportunities_json",
        help="Path to lotus_root_opportunities.json.",
    )
    parser.add_argument(
        "--menu-json",
        default="",
        help="Optional structured_menu.json for existing Fei Da Chu dish photo references.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default="",
        help="Output directory. Defaults to the opportunity JSON directory.",
    )
    parser.add_argument(
        "--images-dir",
        default="",
        help="Image output directory. Defaults to <output-dir>/generated_lotus_images.",
    )
    parser.add_argument(
        "--aspect-ratio",
        default="4:3",
        help="Image aspect ratio passed to image_gen.py. Default: 4:3.",
    )
    parser.add_argument(
        "--image-size",
        default="1K",
        help="Image size passed to image_gen.py. Default: 1K.",
    )
    parser.add_argument(
        "--model",
        default="",
        help="Optional image model override passed to image_gen.py.",
    )
    parser.add_argument(
        "--backend",
        default="",
        help="Optional IMAGE_BACKEND override passed to image_gen.py.",
    )
    parser.add_argument(
        "--concurrency",
        type=int,
        default=1,
        help="Manifest generation concurrency. Default: 1.",
    )
    parser.add_argument(
        "--max-items",
        type=int,
        default=0,
        help="Generate only the first N opportunities. Default: all.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Only write the image manifest and prompt Markdown; do not call the image model.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Regenerate even when an opportunity already has a generated image file.",
    )
    return parser


def _safe_filename(value: str, default: str) -> str:
    clean = re.sub(r"\s+", "_", value.strip())
    clean = re.sub(r"[^\u4e00-\u9fa5A-Za-z0-9_-]+", "", clean)
    clean = clean.strip("_-")
    return clean[:70] or default


def _read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _write_json(path: Path, data: dict[str, Any]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _find_menu_json(opportunity_path: Path, explicit: str) -> Path | None:
    if explicit:
        path = Path(explicit)
        return path if path.exists() else None
    candidate = opportunity_path.parent / "structured_menu.json"
    return candidate if candidate.exists() else None


def _flatten_menu_items(menu: dict[str, Any]) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    for category in menu.get("categories", []):
        for item in category.get("items", []):
            if isinstance(item, dict):
                items.append(item)
    return items


def _collect_reference_notes(menu: dict[str, Any], base_item_name: str) -> list[str]:
    notes: list[str] = []
    all_items = _flatten_menu_items(menu)
    preferred = [
        item for item in all_items
        if base_item_name and item.get("name") == base_item_name
    ]
    fallback = [
        item for item in all_items
        if item.get("category") in {"招牌菜", "热销菜"} and item.get("photos")
    ]
    for item in preferred + fallback:
        for photo in item.get("photos", [])[:2]:
            title = str(photo.get("title") or item.get("name") or "").strip()
            if not title:
                continue
            note = f"{item.get('name', '')}: {title}"
            if note not in notes:
                notes.append(note)
            if len(notes) >= 6:
                return notes
    return notes


def _serving_style(name: str, method: str) -> str:
    text = f"{name} {method}"
    if any(token in text for token in ("汤", "筒骨", "排骨汤")):
        return "served in a warm ceramic soup bowl with visible broth, lotus root chunks, and gentle steam"
    if any(token in text for token in ("干锅", "砂锅")):
        return "served in a small black cast-iron dry pot, sizzling oil sheen, chili and scallion garnish"
    if any(token in text for token in ("凉拌", "藕丁", "酸辣藕尖")):
        return "served as a crisp cold appetizer on a shallow ceramic plate, glossy dressing, fresh herbs"
    if any(token in text for token in ("桂花", "糯米藕", "甜品")):
        return "served as sliced lotus-root dessert on a clean white plate with osmanthus syrup glaze"
    if any(token in text for token in ("藕合", "香酥", "炸")):
        return "served as golden crispy fried lotus-root sandwiches on a dark ceramic plate"
    if any(token in text for token in ("小炒", "擂椒", "剁椒")):
        return "served as a wok-fried Hunan dish on a rustic ceramic plate, strong wok-hei, chili aroma"
    return "served as a polished Chinese restaurant menu dish on rustic ceramic tableware"


def _build_prompt(
    restaurant: str,
    opportunity: dict[str, Any],
    reference_notes: list[str],
) -> str:
    name = str(opportunity.get("name", "")).strip()
    base = str(opportunity.get("base_menu_item", "")).strip()
    method = str(opportunity.get("method_brief", "")).strip()
    lotus_role = str(opportunity.get("lotus_role", "")).strip()
    fit_reason = str(opportunity.get("fit_reason", "")).strip()
    serving = _serving_style(name, method)
    references = "；".join(reference_notes) if reference_notes else "参考费大厨现有湘菜菜单摄影风格：热菜有锅气、油润、近景、真实餐桌质感。"

    return (
        "Use case: ads-marketing\n"
        "Asset type: restaurant menu proposal dish photo\n"
        f"Primary request: Generate a highly appetizing photorealistic concept image for a "
        f"co-created lotus-root dish for {restaurant}: {name}.\n"
        f"Base menu inspiration: {base or 'new lotus-root dish'}.\n"
        f"Dish method: {method}\n"
        f"Lotus-root role: {lotus_role}\n"
        f"Why it fits the restaurant: {fit_reason}\n"
        f"Existing Fei Da Chu visual references to infer style from: {references}\n"
        f"Scene/backdrop: {serving}; premium casual Hunan restaurant table setting, no people.\n"
        "Style/medium: photorealistic Chinese food photography, commercial menu quality, "
        "natural but vivid color, realistic textures, crisp lotus root pores, juicy sauce, "
        "fresh garnish, visible steam when appropriate.\n"
        "Composition/framing: 45-degree close-up, dish fills most of the frame, shallow depth "
        "of field, appetizing hero angle, enough context to show serving vessel.\n"
        "Lighting/mood: warm restaurant lighting, glossy but not greasy, inviting and craveable.\n"
        "Constraints: no text, no captions, no logo, no watermark, no brand mark, no people, "
        "no hands, no packaging, no cartoon style, no deformed food, no impossible ingredients."
    )


def _existing_image_path(opportunity: dict[str, Any], output_dir: Path) -> Path | None:
    generated = opportunity.get("generated_image")
    if not isinstance(generated, dict):
        return None
    rel = generated.get("path") or generated.get("url")
    if not rel:
        return None
    path = Path(str(rel))
    if not path.is_absolute():
        path = output_dir / path
    return path if path.exists() else None


def _build_manifest(
    report: dict[str, Any],
    menu: dict[str, Any],
    images_dir: Path,
    output_dir: Path,
    *,
    aspect_ratio: str,
    image_size: str,
    max_items: int,
    force: bool,
) -> dict[str, Any]:
    restaurant = str(report.get("restaurant", "餐厅"))
    items = report.get("opportunities", [])
    if max_items > 0:
        items = items[:max_items]
    manifest_items: list[dict[str, Any]] = []
    for idx, opportunity in enumerate(items, start=1):
        if not isinstance(opportunity, dict):
            continue
        if not force and _existing_image_path(opportunity, output_dir):
            continue
        name = str(opportunity.get("name", "")).strip() or f"莲藕菜品{idx}"
        stem = f"lotus_{idx:02d}_{_safe_filename(name, f'dish_{idx:02d}')}"
        reference_notes = _collect_reference_notes(menu, str(opportunity.get("base_menu_item", "")))
        manifest_items.append({
            "filename": f"{stem}.png",
            "prompt": _build_prompt(restaurant, opportunity, reference_notes),
            "aspect_ratio": aspect_ratio,
            "image_size": image_size,
            "status": "Pending",
            "purpose": "lotus_root_opportunity_dish_image",
            "type": "ai_generated_food_photo",
            "alt_text": f"{restaurant}合作莲藕菜品概念图：{name}",
            "opportunity_index": idx - 1,
            "opportunity_name": name,
            "base_menu_item": opportunity.get("base_menu_item", ""),
        })
    return {
        "project": "lotus_root_opportunity_images",
        "restaurant": restaurant,
        "generated_at": "",
        "output_dir": str(images_dir),
        "items": manifest_items,
    }


def _run_image_gen(
    manifest_path: Path,
    images_dir: Path,
    *,
    image_size: str,
    model: str,
    backend: str,
    concurrency: int,
) -> int:
    command = [
        sys.executable,
        "-B",
        str(_IMAGE_GEN),
        "--manifest",
        str(manifest_path),
        "--image_size",
        image_size,
        "--output",
        str(images_dir),
        "--concurrency",
        str(max(1, concurrency)),
    ]
    if model:
        command.extend(["--model", model])
    if backend:
        command.extend(["--backend", backend])
    print("Running:", " ".join(command), file=sys.stderr)
    result = subprocess.run(command, cwd=str(_REPO_ROOT), check=False)
    return result.returncode


def _render_manifest_md(manifest_path: Path) -> None:
    subprocess.run(
        [
            sys.executable,
            "-B",
            str(_IMAGE_GEN),
            "--render-md",
            str(manifest_path),
        ],
        cwd=str(_REPO_ROOT),
        check=False,
    )


def _find_generated_file(images_dir: Path, filename: str) -> Path | None:
    stem = Path(filename).stem
    direct = images_dir / filename
    if direct.exists():
        return direct
    matches = sorted(images_dir.glob(stem + ".*"))
    for path in matches:
        if path.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp"}:
            return path
    return None


def _update_report_with_images(
    report: dict[str, Any],
    manifest: dict[str, Any],
    manifest_path: Path,
    images_dir: Path,
    output_dir: Path,
) -> int:
    fresh_manifest = _read_json(manifest_path)
    count = 0
    opportunities = report.get("opportunities", [])
    for item in fresh_manifest.get("items", []):
        idx = item.get("opportunity_index")
        if not isinstance(idx, int) or idx < 0 or idx >= len(opportunities):
            continue
        image_path = _find_generated_file(images_dir, str(item.get("filename", "")))
        entry: dict[str, Any] = {
            "status": item.get("status", ""),
            "filename": item.get("filename", ""),
            "prompt": item.get("prompt", ""),
            "manifest": str(manifest_path.relative_to(output_dir)),
        }
        if image_path:
            entry["path"] = str(image_path.relative_to(output_dir))
            entry["url"] = entry["path"].replace(os.sep, "/")
            count += 1
        if item.get("last_error"):
            entry["last_error"] = item.get("last_error")
        opportunities[idx]["generated_image"] = entry
    return count


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    opportunity_path = Path(args.opportunities_json)
    if not opportunity_path.exists():
        print(f"Opportunity JSON not found: {opportunity_path}", file=sys.stderr)
        return 1
    output_dir = Path(args.output_dir) if args.output_dir else opportunity_path.parent
    images_dir = Path(args.images_dir) if args.images_dir else output_dir / "generated_lotus_images"
    output_dir.mkdir(parents=True, exist_ok=True)
    images_dir.mkdir(parents=True, exist_ok=True)

    menu_path = _find_menu_json(opportunity_path, args.menu_json)
    if menu_path is None:
        print(
            "Warning: structured_menu.json not found; prompts will use generic Fei Da Chu style notes.",
            file=sys.stderr,
        )
        menu: dict[str, Any] = {}
    else:
        menu = _read_json(menu_path)

    report = _read_json(opportunity_path)
    manifest = _build_manifest(
        report,
        menu,
        images_dir,
        output_dir,
        aspect_ratio=args.aspect_ratio,
        image_size=args.image_size,
        max_items=args.max_items,
        force=args.force,
    )
    manifest_path = images_dir / "lotus_image_prompts.json"
    _write_json(manifest_path, manifest)
    _render_manifest_md(manifest_path)

    if not manifest["items"]:
        print("No images to generate; existing generated_image entries are present.", file=sys.stderr)
        print(json.dumps({
            "updated_json": str(opportunity_path),
            "manifest": str(manifest_path),
            "images_dir": str(images_dir),
        }, ensure_ascii=False, indent=2))
        return 0

    if args.dry_run:
        print(json.dumps({
            "manifest": str(manifest_path),
            "prompt_markdown": str(manifest_path.with_suffix(".md")),
            "images_dir": str(images_dir),
            "dry_run": True,
        }, ensure_ascii=False, indent=2))
        return 0

    rc = _run_image_gen(
        manifest_path,
        images_dir,
        image_size=args.image_size,
        model=args.model,
        backend=args.backend,
        concurrency=args.concurrency,
    )
    generated_count = _update_report_with_images(
        report,
        manifest,
        manifest_path,
        images_dir,
        output_dir,
    )
    _write_json(opportunity_path, report)
    print(json.dumps({
        "updated_json": str(opportunity_path),
        "manifest": str(manifest_path),
        "images_dir": str(images_dir),
        "generated_images": generated_count,
        "image_gen_exit_code": rc,
    }, ensure_ascii=False, indent=2))
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
