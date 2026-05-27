#!/usr/bin/env python3
"""
PPT Master - Restaurant Menu Analyzer

Search public web pages for a restaurant, ask an OpenAI-compatible model to
extract a structured menu, and enrich each dish with photo candidates.

Usage:
    python3 proposalGen/restaurant_menu_analyzer.py <restaurant_name>

Examples:
    python3 proposalGen/restaurant_menu_analyzer.py 费大厨
    python3 proposalGen/restaurant_menu_analyzer.py 费大厨 -o projects/menu_research/feidachu
    python3 proposalGen/restaurant_menu_analyzer.py 费大厨 --download-photos

Dependencies:
    requests, beautifulsoup4

Environment:
    OPENAI_API_KEY   required
    OPENAI_TEXT_MODEL required unless --model is provided
    OPENAI_BASE_URL  optional, defaults to https://api.openai.com/v1
"""

from __future__ import annotations

import argparse
import base64
import datetime as dt
import json
import os
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, quote_plus, unquote, urljoin, urlparse

import requests
from bs4 import BeautifulSoup

try:
    from curl_cffi import requests as curl_requests  # type: ignore
    from curl_cffi.requests import exceptions as curl_exceptions  # type: ignore

    _CURL_IMPERSONATE = "chrome120"
except ImportError:
    curl_requests = None
    curl_exceptions = None
    _CURL_IMPERSONATE = None

_REPO_ROOT = Path(__file__).resolve().parent.parent
_PPT_SCRIPTS_DIR = _REPO_ROOT / "skills" / "ppt-master" / "scripts"
if str(_PPT_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_PPT_SCRIPTS_DIR))

from config import load_prefixed_env_file  # noqa: E402
from image_backends.backend_common import download_image  # noqa: E402


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
)
DEFAULT_CATEGORIES = ("招牌菜", "热销菜", "素菜", "主食", "甜品/饮品", "其他")
TEXT_LIMIT_PER_SOURCE = 9000
MODEL_SOURCE_CHAR_BUDGET = 48000
REQUEST_TIMEOUT = 25
SEARCH_TIMEOUT = 8
if curl_exceptions is not None:
    NETWORK_ERRORS = (requests.RequestException, curl_exceptions.RequestException)
else:
    NETWORK_ERRORS = (requests.RequestException,)
DISABLED_SEARCH_PROVIDERS: set[str] = set()


@dataclass
class SearchResult:
    title: str
    url: str
    snippet: str = ""


@dataclass
class PageSource:
    source_id: str
    title: str
    url: str
    text: str
    images: list[dict[str, str]] = field(default_factory=list)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Analyze a restaurant menu from public web sources.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("restaurant", help="Restaurant name, for example: 费大厨")
    parser.add_argument(
        "-o",
        "--output-dir",
        default="",
        help="Output directory. Defaults to projects/menu_research/<restaurant>_<timestamp>.",
    )
    parser.add_argument(
        "--model",
        default="",
        help="Override OPENAI_TEXT_MODEL from .env.",
    )
    parser.add_argument(
        "--base-url",
        default="",
        help="Override OPENAI_BASE_URL from .env.",
    )
    parser.add_argument(
        "--max-sources",
        type=int,
        default=14,
        help="Maximum unique pages to fetch before model analysis. Default: 14.",
    )
    parser.add_argument(
        "--max-photos",
        type=int,
        default=3,
        help="Maximum photo candidates per dish. Default: 3.",
    )
    parser.add_argument(
        "--download-photos",
        action="store_true",
        help="Download selected photo candidates into output_dir/photos.",
    )
    parser.add_argument(
        "--include-platform-query",
        action="append",
        default=[],
        help="Extra search query to run. Can be repeated.",
    )
    return parser


def _load_env() -> None:
    load_prefixed_env_file(("OPENAI_",))


def _http_get(url: str, *, timeout: int = REQUEST_TIMEOUT) -> requests.Response:
    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    }
    if curl_requests is not None:
        response = curl_requests.get(
            url,
            headers=headers,
            timeout=timeout,
            impersonate=_CURL_IMPERSONATE,
        )
    else:
        response = requests.get(url, headers=headers, timeout=timeout)
    apparent = getattr(response, "apparent_encoding", None)
    if apparent:
        response.encoding = apparent
    return response


def _normalize_url(url: str) -> str:
    parsed = urlparse(url)
    if parsed.netloc.endswith("bing.com") and parsed.path == "/ck/a":
        qs = parse_qs(parsed.query)
        if "u" in qs:
            raw = unquote(qs["u"][0])
            return _decode_base64_url_token(raw) or raw
    if not parsed.scheme:
        decoded = _decode_base64_url_token(url)
        if decoded:
            return decoded
    return url


def _decode_base64_url_token(token: str) -> str:
    """Decode Bing's occasional bare base64 URL tokens."""
    compact = token.strip()
    candidates = [compact]
    if compact.startswith("a1") and len(compact) > 8:
        candidates.append(compact[2:])
    for candidate in candidates:
        try:
            padded = candidate + ("=" * (-len(candidate) % 4))
            decoded = base64.urlsafe_b64decode(padded.encode("ascii")).decode("utf-8")
        except (UnicodeDecodeError, ValueError, TypeError):
            continue
        parsed = urlparse(decoded)
        if parsed.scheme in {"http", "https"} and parsed.netloc:
            return decoded
    return ""


def _safe_filename(value: str, default: str = "output") -> str:
    clean = re.sub(r"\s+", "_", value.strip())
    clean = re.sub(r"[^\u4e00-\u9fa5A-Za-z0-9_-]+", "", clean)
    clean = clean.strip("_-")
    return clean[:80] or default


def _extract_visible_text(html: str) -> tuple[str, str]:
    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    for tag in soup(["script", "style", "noscript", "svg", "canvas", "iframe"]):
        tag.decompose()
    text = soup.get_text("\n", strip=True)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return title, text


def _extract_images(html: str, page_url: str) -> list[dict[str, str]]:
    soup = BeautifulSoup(html, "html.parser")
    candidates: list[dict[str, str]] = []
    seen: set[str] = set()
    for img in soup.find_all("img"):
        src = ""
        for attr in ("src", "data-src", "data-original", "data-lazy-src", "data-url"):
            raw = img.get(attr)
            if raw:
                src = str(raw)
                break
        if not src or src.startswith("data:"):
            continue
        abs_url = urljoin(page_url, src)
        if abs_url in seen:
            continue
        seen.add(abs_url)
        alt = img.get("alt") or img.get("title") or ""
        candidates.append({
            "url": abs_url,
            "alt": str(alt).strip(),
            "source_url": page_url,
        })
    return candidates


def _search_duckduckgo(query: str, limit: int) -> list[SearchResult]:
    url = f"https://duckduckgo.com/html/?q={quote_plus(query)}"
    response = _http_get(url, timeout=SEARCH_TIMEOUT)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    results: list[SearchResult] = []
    for node in soup.select("a.result__a"):
        href = node.get("href") or ""
        title = node.get_text(" ", strip=True)
        if not href or not title:
            continue
        if "uddg=" in href:
            parsed = urlparse(href)
            qs = parse_qs(parsed.query)
            href = unquote(qs.get("uddg", [href])[0])
        snippet_node = node.find_parent(class_="result")
        snippet = ""
        if snippet_node:
            body = snippet_node.select_one(".result__snippet")
            snippet = body.get_text(" ", strip=True) if body else ""
        results.append(SearchResult(title=title, url=_normalize_url(href), snippet=snippet))
        if len(results) >= limit:
            break
    return results


def _search_bing(query: str, limit: int) -> list[SearchResult]:
    url = f"https://www.bing.com/search?q={quote_plus(query)}"
    response = _http_get(url, timeout=SEARCH_TIMEOUT)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    results: list[SearchResult] = []
    for node in soup.select("li.b_algo"):
        link = node.select_one("h2 a")
        if not link:
            continue
        href = link.get("href") or ""
        title = link.get_text(" ", strip=True)
        snippet_node = node.select_one(".b_caption p")
        snippet = snippet_node.get_text(" ", strip=True) if snippet_node else ""
        if href and title:
            results.append(SearchResult(title=title, url=_normalize_url(href), snippet=snippet))
        if len(results) >= limit:
            break
    return results


def search_web(query: str, limit: int) -> list[SearchResult]:
    """Search with HTML endpoints so the script needs no search API key."""
    providers = (
        ("duckduckgo", _search_duckduckgo),
        ("bing", _search_bing),
    )
    for provider_name, provider in providers:
        if provider_name in DISABLED_SEARCH_PROVIDERS:
            continue
        try:
            results = provider(query, limit)
            if results:
                return results
        except (*NETWORK_ERRORS, ValueError) as exc:
            print(f"  search warning [{provider_name}]: {query}: {exc}", file=sys.stderr)
            if provider_name == "duckduckgo":
                DISABLED_SEARCH_PROVIDERS.add(provider_name)
                print("  DuckDuckGo unavailable; using Bing for remaining queries.", file=sys.stderr)
    return []


def build_queries(restaurant: str, extra_queries: list[str]) -> list[str]:
    base = [
        f"{restaurant} 菜单 招牌菜 热销菜",
        f"{restaurant} 官方 菜单 菜品",
        f"{restaurant} 大众点评 菜单 推荐菜",
        f"{restaurant} 小红书 菜单 推荐菜",
        f"{restaurant} 外卖 菜单 美团 饿了么",
        f"{restaurant} 菜品 推荐菜 饮品 主食",
        f"{restaurant} 点菜攻略 必点菜",
    ]
    return base + extra_queries


def collect_sources(restaurant: str, max_sources: int, extra_queries: list[str]) -> list[PageSource]:
    seen_urls: set[str] = set()
    sources: list[PageSource] = []
    queries = build_queries(restaurant, extra_queries)
    for query in queries:
        if len(sources) >= max_sources:
            break
        print(f"Searching: {query}", file=sys.stderr)
        for result in search_web(query, limit=6):
            if len(sources) >= max_sources:
                break
            if result.url in seen_urls:
                continue
            seen_urls.add(result.url)
            try:
                response = _http_get(result.url)
                response.raise_for_status()
                title, text = _extract_visible_text(response.text)
                if restaurant not in (title + text) and len(text) < 800:
                    continue
                images = _extract_images(response.text, result.url)
            except (*NETWORK_ERRORS, ValueError) as exc:
                print(f"  fetch warning: {result.url}: {exc}", file=sys.stderr)
                title = result.title
                text = result.snippet
                images = []
            if not text.strip():
                continue
            source_id = f"S{len(sources) + 1:02d}"
            sources.append(
                PageSource(
                    source_id=source_id,
                    title=title or result.title,
                    url=result.url,
                    text=text[:TEXT_LIMIT_PER_SOURCE],
                    images=images,
                )
            )
            time.sleep(0.2)
    return sources


def _make_chat_url(base_url: str) -> str:
    base = (base_url or "https://api.openai.com/v1").rstrip("/")
    if base.endswith("/chat/completions"):
        return base
    return f"{base}/chat/completions"


def _strip_json_fence(content: str) -> str:
    text = content.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        return text[start:end + 1]
    return text


def _source_payload(sources: list[PageSource]) -> list[dict[str, str]]:
    payload: list[dict[str, str]] = []
    total = 0
    for source in sources:
        remaining = MODEL_SOURCE_CHAR_BUDGET - total
        if remaining <= 0:
            break
        text = source.text[:remaining]
        total += len(text)
        payload.append({
            "source_id": source.source_id,
            "title": source.title,
            "url": source.url,
            "text": text,
        })
    return payload


def analyze_with_model(
    restaurant: str,
    sources: list[PageSource],
    *,
    model: str,
    base_url: str,
) -> dict[str, Any]:
    api_key = os.environ.get("OPENAI_API_KEY", "")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not set. Add it to .env or the current environment.")
    if not model:
        raise RuntimeError("OPENAI_TEXT_MODEL is not set. Add it to .env or pass --model.")

    schema = {
        "restaurant": restaurant,
        "analysis_scope": "short note about source limits and uncertainty",
        "categories": [
            {
                "name": "招牌菜",
                "items": [
                    {
                        "name": "菜品名",
                        "description": "80-160字，介绍做法、口味、卖点、适合场景",
                        "ingredients": ["主要食材"],
                        "flavor_profile": "口味关键词",
                        "source_evidence": [
                            {
                                "source_id": "S01",
                                "evidence": "不超过40字的来源依据，不能编造",
                            }
                        ],
                        "confidence": "high|medium|low",
                        "photo_queries": [f"{restaurant} 菜品名 图片"],
                    }
                ],
            }
        ],
    }
    user_payload = {
        "restaurant": restaurant,
        "required_categories": list(DEFAULT_CATEGORIES),
        "source_pages": _source_payload(sources),
        "output_schema": schema,
    }
    messages = [
        {
            "role": "system",
            "content": (
                "你是餐饮品牌菜单研究员。你只能依据给定公开来源文本做判断，"
                "不要编造菜名、价格、门店限定信息。若来源不足，把 confidence 设为 low。"
                "输出必须是严格 JSON，不要 Markdown，不要解释。"
            ),
        },
        {
            "role": "user",
            "content": (
                "请基于以下来源抽取尽可能全面的餐厅菜单，按指定分类归并。"
                "每道菜要有详细介绍、主要食材、口味、来源依据和置信度。"
                "同一菜品有别名时合并，并在描述中说明。\n\n"
                + json.dumps(user_payload, ensure_ascii=False)
            ),
        },
    ]
    response = requests.post(
        _make_chat_url(base_url),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        json={
            "model": model,
            "messages": messages,
            "temperature": 0.2,
            "response_format": {"type": "json_object"},
        },
        timeout=180,
    )
    if response.status_code >= 400:
        body = response.text[:1000]
        raise RuntimeError(f"Model request failed ({response.status_code}): {body}")
    data = response.json()
    content = data["choices"][0]["message"]["content"]
    return json.loads(_strip_json_fence(content))


def _iter_menu_items(menu: dict[str, Any]) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    for category in menu.get("categories", []):
        category_name = category.get("name", "")
        for item in category.get("items", []):
            if isinstance(item, dict) and item.get("name"):
                item.setdefault("category", category_name)
                items.append(item)
    return items


def _score_page_image(
    image: dict[str, str],
    restaurant: str,
    dish_name: str,
) -> int:
    tokens = _dish_photo_tokens(restaurant, dish_name)
    haystack = " ".join([image.get("alt", ""), image.get("url", "")]).lower()
    if not any(token in haystack for token in tokens):
        return 0
    score = 0
    for token in tokens:
        if token in haystack:
            score += 8
    if restaurant and restaurant.lower() in haystack:
        score += 2
    if any(word in haystack for word in ("menu", "dish", "food", "菜", "美食", "辣椒炒肉")):
        score += 1
    if any(word in haystack for word in ("logo", "icon", "avatar", "sprite", "二维码")):
        score -= 20
    return score


def _dish_photo_tokens(
    restaurant: str,
    dish_name: str,
    extra_queries: list[str] | None = None,
) -> list[str]:
    candidates = [dish_name, dish_name.replace(restaurant, "")]
    for query in extra_queries or []:
        cleaned = query
        for token in (restaurant, "菜品", "图片", "照片", "菜单", "高清", "实拍"):
            cleaned = cleaned.replace(token, " ")
        candidates.extend(re.split(r"[\s,，/|]+", cleaned))
    tokens: list[str] = []
    for candidate in candidates:
        token = candidate.strip().lower()
        if len(token) < 2 or token in {restaurant.lower(), "菜", "图"}:
            continue
        if token not in tokens:
            tokens.append(token)
    return tokens


def _score_photo_candidate(
    photo: dict[str, str],
    restaurant: str,
    dish_tokens: list[str],
) -> int:
    title = photo.get("title", "").lower()
    haystack = " ".join(
        [
            photo.get("title", ""),
            photo.get("url", ""),
            photo.get("source_url", ""),
        ]
    ).lower()
    if any(word in title for word in ("logo", "icon", "avatar", "人物", "红卫兵", "新闻")):
        return 0
    score = 0
    for token in dish_tokens:
        if token in haystack:
            score += 10
        if token in title:
            score += 8
    if score == 0:
        return 0
    if restaurant and restaurant.lower() in haystack:
        score += 2
    return score


def _search_bing_images(
    query: str,
    limit: int,
    *,
    restaurant: str = "",
    dish_tokens: list[str] | None = None,
) -> list[dict[str, str]]:
    url = f"https://www.bing.com/images/search?q={quote_plus(query)}&form=HDRSC2"
    response = _http_get(url)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    ranked: list[tuple[int, dict[str, str]]] = []
    seen: set[str] = set()
    for node in soup.select("a.iusc"):
        raw = node.get("m")
        if not raw:
            continue
        try:
            meta = json.loads(str(raw))
        except json.JSONDecodeError:
            continue
        image_url = meta.get("murl") or ""
        if not image_url or image_url in seen:
            continue
        seen.add(image_url)
        photo = {
            "url": image_url,
            "source_url": meta.get("purl") or "",
            "title": meta.get("t") or query,
            "provider": "bing_images",
        }
        score = (
            _score_photo_candidate(photo, restaurant, dish_tokens)
            if dish_tokens else 1
        )
        if score > 0:
            ranked.append((score, photo))
        if len(ranked) >= max(limit * 3, limit):
            break
    ranked.sort(key=lambda item: item[0], reverse=True)
    return [photo for _, photo in ranked[:limit]]


def _photo_candidates_from_sources(
    sources: list[PageSource],
    restaurant: str,
    dish_name: str,
    limit: int,
) -> list[dict[str, str]]:
    if limit <= 0:
        return []
    ranked: list[tuple[int, dict[str, str]]] = []
    for source in sources:
        for image in source.images:
            score = _score_page_image(image, restaurant, dish_name)
            if score > 0:
                ranked.append((score, {
                    "url": image["url"],
                    "source_url": image.get("source_url", source.url),
                    "title": image.get("alt") or f"{restaurant} {dish_name}",
                    "provider": "source_page",
                }))
    ranked.sort(key=lambda item: item[0], reverse=True)
    seen: set[str] = set()
    selected: list[dict[str, str]] = []
    for _, image in ranked:
        if image["url"] in seen:
            continue
        seen.add(image["url"])
        selected.append(image)
        if len(selected) >= limit:
            break
    return selected


def enrich_photos(
    menu: dict[str, Any],
    sources: list[PageSource],
    restaurant: str,
    max_photos: int,
) -> None:
    for item in _iter_menu_items(menu):
        dish_name = str(item.get("name", "")).strip()
        if not dish_name:
            continue
        if max_photos <= 0:
            item["photos"] = []
            continue
        photo_queries = [
            str(query)
            for query in item.get("photo_queries", [])
            if str(query).strip()
        ]
        dish_tokens = _dish_photo_tokens(restaurant, dish_name, photo_queries)
        photos = _photo_candidates_from_sources(sources, restaurant, dish_name, max_photos)
        queries = photo_queries or [f"{restaurant} {dish_name} 菜品 图片"]
        for query in queries:
            if len(photos) >= max_photos:
                break
            try:
                photos.extend(_search_bing_images(
                    query,
                    max_photos - len(photos),
                    restaurant=restaurant,
                    dish_tokens=dish_tokens,
                ))
            except (*NETWORK_ERRORS, ValueError) as exc:
                print(f"  photo search warning: {query}: {exc}", file=sys.stderr)
        item["photos"] = _dedupe_photos(photos)[:max_photos]
        time.sleep(0.15)


def _dedupe_photos(photos: list[dict[str, str]]) -> list[dict[str, str]]:
    selected: list[dict[str, str]] = []
    seen: set[str] = set()
    for photo in photos:
        key = photo.get("url") or photo.get("source_url") or photo.get("title", "")
        if not key or key in seen:
            continue
        seen.add(key)
        selected.append(photo)
    return selected


def download_photo_candidates(menu: dict[str, Any], output_dir: Path) -> None:
    photo_dir = output_dir / "photos"
    photo_dir.mkdir(parents=True, exist_ok=True)
    for item in _iter_menu_items(menu):
        dish_name = _safe_filename(str(item.get("name", "")), default="dish")
        downloaded: list[dict[str, str]] = []
        for idx, photo in enumerate(item.get("photos", []), start=1):
            url = photo.get("url", "")
            if not url:
                continue
            suffix = Path(urlparse(url).path).suffix
            if suffix.lower() not in {".jpg", ".jpeg", ".png", ".webp", ".gif"}:
                suffix = ".jpg"
            path = photo_dir / f"{dish_name}_{idx:02d}{suffix}"
            try:
                actual_path = download_image(url, str(path), headers={"User-Agent": USER_AGENT})
                saved = dict(photo)
                saved["local_path"] = str(Path(actual_path).relative_to(output_dir))
                downloaded.append(saved)
            except (*NETWORK_ERRORS, OSError, RuntimeError, ValueError) as exc:
                print(f"  download warning: {url}: {exc}", file=sys.stderr)
        if downloaded:
            item["photos"] = downloaded


def write_markdown(menu: dict[str, Any], output_path: Path) -> None:
    lines = [
        f"# {menu.get('restaurant', '餐厅')} 菜单分析",
        "",
        str(menu.get("analysis_scope", "")).strip(),
        "",
    ]
    for category in menu.get("categories", []):
        lines.append(f"## {category.get('name', '未分类')}")
        lines.append("")
        for item in category.get("items", []):
            lines.append(f"### {item.get('name', '')}")
            lines.append("")
            if item.get("description"):
                lines.append(str(item["description"]))
                lines.append("")
            if item.get("ingredients"):
                lines.append(f"- 主要食材：{'、'.join(map(str, item['ingredients']))}")
            if item.get("flavor_profile"):
                lines.append(f"- 口味：{item['flavor_profile']}")
            if item.get("confidence"):
                lines.append(f"- 置信度：{item['confidence']}")
            evidence = item.get("source_evidence") or []
            if evidence:
                compact = []
                for ev in evidence[:3]:
                    compact.append(f"{ev.get('source_id', '')}: {ev.get('evidence', '')}")
                lines.append(f"- 来源依据：{'；'.join(compact)}")
            photos = item.get("photos") or []
            if photos:
                lines.append("- 菜品照片：")
                for photo in photos:
                    label = photo.get("title") or item.get("name", "photo")
                    target = photo.get("local_path") or photo.get("url")
                    source = photo.get("source_url", "")
                    lines.append(f"  - [{label}]({target}) 来源：{source}")
            lines.append("")
    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def write_sources(sources: list[PageSource], output_path: Path) -> None:
    data = [
        {
            "source_id": source.source_id,
            "title": source.title,
            "url": source.url,
            "text_excerpt": source.text[:1200],
            "image_count": len(source.images),
        }
        for source in sources
    ]
    output_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def default_output_dir(restaurant: str) -> Path:
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    return Path("projects") / "menu_research" / f"{_safe_filename(restaurant)}_{timestamp}"


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    _load_env()

    restaurant = args.restaurant.strip()
    if not restaurant:
        print("Restaurant name cannot be empty.", file=sys.stderr)
        return 1

    output_dir = Path(args.output_dir) if args.output_dir else default_output_dir(restaurant)
    output_dir.mkdir(parents=True, exist_ok=True)

    model = args.model or os.environ.get("OPENAI_TEXT_MODEL", "")
    base_url = args.base_url or os.environ.get("OPENAI_BASE_URL", "")

    try:
        sources = collect_sources(
            restaurant,
            max_sources=max(1, args.max_sources),
            extra_queries=args.include_platform_query,
        )
        if not sources:
            print("No public source pages were fetched. Try extra --include-platform-query values.", file=sys.stderr)
            return 1
        write_sources(sources, output_dir / "sources.json")
        print(f"Fetched {len(sources)} source pages.", file=sys.stderr)

        menu = analyze_with_model(
            restaurant,
            sources,
            model=model,
            base_url=base_url,
        )
        enrich_photos(menu, sources, restaurant, max(0, args.max_photos))
        if args.download_photos:
            download_photo_candidates(menu, output_dir)

        json_path = output_dir / "structured_menu.json"
        md_path = output_dir / "structured_menu.md"
        json_path.write_text(json.dumps(menu, ensure_ascii=False, indent=2), encoding="utf-8")
        write_markdown(menu, md_path)

        print(json.dumps({
            "structured_menu_json": str(json_path),
            "structured_menu_md": str(md_path),
            "sources_json": str(output_dir / "sources.json"),
        }, ensure_ascii=False, indent=2))
        return 0
    except (RuntimeError, *NETWORK_ERRORS, json.JSONDecodeError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
