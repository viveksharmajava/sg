"""
Crawl SG (Sanspareils Greenlands) products from shop.teamsg.in into import Excel + images.

All outputs live under: catalog/templates/pp/pp/SG/<collection-key>/

Resume: each category keeps JSON state (e.g. SG/english-willow/_sg_english_willow_state.json).
Already-crawled Shopify handles are skipped; Excel is rebuilt from all saved rows.

Review crawl (default 3 new products per category):
  python sg_store_crawler.py --headless

Full crawl (only products not yet in state):
  python sg_store_crawler.py --all --headless

Re-crawl everything:
  python sg_store_crawler.py --all --headless --force
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import re
import time
from dataclasses import dataclass, field
from html import unescape
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

import openpyxl
import requests
import urllib3
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://shop.teamsg.in"
BRAND = "SG"
MAX_PRODUCT_ID_LEN = 64
IMAGES_DIR_NAME = "images"
REQUESTS_VERIFY_SSL = False
DEFAULT_PAGE_LOAD_TIMEOUT_S = 45
DEFAULT_PDP_WAIT_S = 25

SCRIPT_DIR = Path(__file__).resolve().parent
SG_ROOT = Path(r"C:\vivek\project\catalog\templates\pp\pp\SG")
TEMPLATE_XLSX = SCRIPT_DIR / "ss-bats-product-import.xlsx"

COMMON_ID_DROP = frozenset(
    {
        "SG",
        "SS",
        "THE",
        "AND",
        "FOR",
        "WITH",
        "CRICKET",
        "BAT",
        "BATS",
        "BALL",
        "BALLS",
        "WILLOW",
        "ENGLISH",
        "KASHMIR",
        "TENNIS",
        "SHOE",
        "SHOES",
        "STUD",
        "SPIKE",
        "KIT",
        "KITBAG",
        "BAG",
        "BAGS",
        "WHEELIE",
        "DUFFLE",
        "GLOVE",
        "GLOVES",
        "PAD",
        "PADS",
        "GUARD",
        "GUARDS",
        "HELMET",
        "HELMETS",
        "SOCK",
        "SOCKS",
        "SHIRT",
        "TROUSER",
        "TROUSERS",
        "COMBO",
        "WHITE",
        "WHITES",
        "LEATHER",
        "SYNTHETIC",
        "INNER",
        "KEEPING",
        "WICKET",
        "BATTING",
        "THIGH",
        "ELBOW",
        "ACCESSORY",
        "ACCESSORIES",
    }
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("sg_store_crawler")


@dataclass(frozen=True)
class CollectionProfile:
    key: str
    collection_slugs: tuple[str, ...]
    category_id: str
    product_id_prefix: str
    type_label: str
    title_strip_patterns: tuple[str, ...]
    excel_out: Path
    out_root: Path
    state_file: str

    @property
    def collection_urls(self) -> list[str]:
        return [f"{BASE_URL}/collections/{slug}" for slug in self.collection_slugs]


def _sg_profile(
    key: str,
    slugs: tuple[str, ...] | str,
    category_id: str,
    product_id_prefix: str,
    type_label: str,
    title_strip_patterns: tuple[str, ...],
) -> CollectionProfile:
    slug_tuple = slugs if isinstance(slugs, tuple) else (slugs,)
    return CollectionProfile(
        key=key,
        collection_slugs=slug_tuple,
        category_id=category_id,
        product_id_prefix=product_id_prefix,
        type_label=type_label,
        title_strip_patterns=title_strip_patterns,
        excel_out=SG_ROOT / key / f"sg-{key}-import-crawled.xlsx",
        out_root=SG_ROOT / key,
        state_file=f"_sg_{key.replace('-', '_')}_state.json",
    )


COLLECTION_PROFILES: dict[str, CollectionProfile] = {
    "english-willow": _sg_profile(
        "english-willow",
        "english-willow-bat",
        "CC_EW_BATS",
        "SG_EW_BAT_",
        "English Willow Cricket Bat",
        (r"\s*english\s+willow\s*(cricket\s*)?(bat|bats)?\s*", r"\s*cricket\s+bat\s*"),
    ),
    "kashmir-willow": _sg_profile(
        "kashmir-willow",
        "kashmir-willow-bat",
        "CC_KW_BATS",
        "SG_KW_BAT_",
        "Kashmir Willow Cricket Bat",
        (r"\s*kashmir\s+willow\s*(cricket\s*)?(bat|bats)?\s*", r"\s*cricket\s+bat\s*"),
    ),
    "tennis-bats": _sg_profile(
        "tennis-bats",
        "tennis-bat",
        "CC_TENNIS_BATS",
        "SG_TENNIS_BAT_",
        "Tennis Ball Cricket Bat",
        (r"\s*tennis\s*(ball\s*)?(cricket\s*)?(bat|bats)\s*", r"\s*plastic\s*(bat|bats)\s*"),
    ),
    "helmets": _sg_profile(
        "helmets",
        "helmets",
        "CC_HELMET",
        "SG_HELMET_",
        "Cricket Helmet",
        (r"\s*cricket\s*helmets?\s*", r"\s*helmets?\s*"),
    ),
    "caps": _sg_profile(
        "caps",
        "accessories-1",
        "CC_CAPS",
        "SG_CAP_",
        "Cricket Cap",
        (r"\s*cricket\s*caps?\s*", r"\s*caps?\s*"),
    ),
    "leather-balls": _sg_profile(
        "leather-balls",
        "leather-ball-1",
        "CC_LEATHER_BALLS",
        "SG_LEATHER_BALL_",
        "Leather Cricket Ball",
        (r"\s*leather\s*(cricket\s*)?(balls?)\s*", r"\s*cricket\s+ball\s*"),
    ),
    "tennis-balls": _sg_profile(
        "tennis-balls",
        "tennis-ball",
        "CC_TENNIS_BALLS",
        "SG_TENNIS_BALL_",
        "Tennis Cricket Ball",
        (r"\s*tennis\s*(cricket\s*)?(balls?)\s*", r"\s*cricket\s+ball\s*"),
    ),
    "duffle-kitbags": _sg_profile(
        "duffle-kitbags",
        "duffle",
        "CC_DUFFLE_KITBAGS",
        "SG_DUFFLE_KITBAG_",
        "Cricket Kitbag",
        (r"\s*duffle\s*(kit\s*)?(bags?|bag)\s*", r"\s*(kit\s*)?(bags?|bag)\s*"),
    ),
    "wheelie-kitbags": _sg_profile(
        "wheelie-kitbags",
        ("wheelie-duffle", "wheelie-kit"),
        "CC_WHEELIE_KITBAGS",
        "SG_WHEELIE_KITBAG_",
        "Wheelie Cricket Kitbag",
        (r"\s*wheelie\s*(duffle\s*)?(kit\s*)?(bags?|bag)\s*", r"\s*(kit\s*)?(bags?|bag)\s*"),
    ),
    "shoes": _sg_profile(
        "shoes",
        ("stud", "spike"),
        "CC_SHOES",
        "SG_SHOES_",
        "Cricket Shoes",
        (r"\s*cricket\s*(shoes?|spikes?|studs?)\s*", r"\s*(shoes?|spikes?|studs?)\s*"),
    ),
    "socks": _sg_profile(
        "socks",
        "socks",
        "CC_SOCKS",
        "SG_SOCKS_",
        "Cricket Socks",
        (r"\s*cricket\s*socks?\s*", r"\s*socks?\s*"),
    ),
    "combo-whites": _sg_profile(
        "combo-whites",
        "combo",
        "CC_COMBO_WHITES",
        "SG_COMBO_WHITES_",
        "Cricket Combo Whites",
        (r"\s*combo\s*(cricket\s*)?(whites?)\s*", r"\s*cricket\s+whites?\s*"),
    ),
    "whites": _sg_profile(
        "whites",
        ("cricket-shirt", "cricket-trousers"),
        "CC_WHITES",
        "SG_WHITES_",
        "Cricket Whites",
        (r"\s*cricket\s*(shirt|trousers?|pants?)\s*", r"\s*(shirt|trousers?|pants?)\s*"),
    ),
    "inner-gloves": _sg_profile(
        "inner-gloves",
        "inner-gloves",
        "CC_INNER_GLOVES",
        "SG_INNER_GLOVES_",
        "Inner Cricket Gloves",
        (r"\s*inner\s*(cricket\s*)?(gloves?)\s*",),
    ),
    "batting-gloves": _sg_profile(
        "batting-gloves",
        "batting-gloves",
        "CC_BATTING_GLOVES",
        "SG_BATTING_GLOVES_",
        "Cricket Batting Gloves",
        (r"\s*batting\s*(cricket\s*)?(gloves?)\s*",),
    ),
    "keeping-gloves": _sg_profile(
        "keeping-gloves",
        "wicket-keeping-gloves",
        "CC_KEEPING_GLOVES",
        "SG_KEEPING_GLOVES_",
        "Wicket Keeping Gloves",
        (r"\s*(wicket\s*)?keeping\s*(cricket\s*)?(gloves?)\s*",),
    ),
    "batting-pads": _sg_profile(
        "batting-pads",
        "batting-legguards",
        "CC_BATTING_PADS",
        "SG_BATTING_PAD_",
        "Cricket Batting Pads",
        (r"\s*batting\s*(cricket\s*)?(pads?|leg\s*guards?)\s*",),
    ),
    "thigh-pads": _sg_profile(
        "thigh-pads",
        "thigh-pad",
        "CC_THIGH_PADS",
        "SG_THIGH_PAD_",
        "Cricket Thigh Pad",
        (r"\s*thigh\s*(guards?|pads?)\s*",),
    ),
    "elbow-guards": _sg_profile(
        "elbow-guards",
        "elbow-guard",
        "CC_ELBOW_GUARDS",
        "SG_ELBOW_GUARD_",
        "Cricket Elbow Guard",
        (r"\s*elbow\s*(guards?|guard)\s*",),
    ),
}

DEFAULT_COLLECTION_KEYS = tuple(COLLECTION_PROFILES.keys())

# Minimal import columns (matches ss-bats-product-import / ss-product-import template)
IMPORT_HEADERS = [
    "product_id",
    "sku",
    "product_type_id",
    "status_id",
    "category_ids",
    "internal_name",
    "brand_name",
    "product_name",
    "description",
    "long_description",
    "comments",
    "keywords",
    "is_virtual",
    "is_variant",
    "returnable",
    "taxable",
    "charge_shipping",
    "require_inventory",
    "introduction_date",
    "small_image",
    "medium_image",
    "large_image",
    "original_image",
    "other_images",
    "currency",
    "tax_rate",
    "DEFAULT_PRICE",
    "LIST_PRICE",
    "MAXIMUM_PRICE",
    "features",
]


@dataclass
class CrawledProduct:
    product_id: str
    title: str
    handle: str
    pdp_url: str
    description: str
    long_description: str
    price_inr: int | None
    image_urls: list[str] = field(default_factory=list)
    features: list[dict[str, str]] = field(default_factory=list)
    category_id: str | None = None
    type_label: str | None = None


def load_profile_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"version": 2, "done": {}, "order": []}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"version": 2, "done": {}, "order": []}
    if "done" not in data:
        data["done"] = {}
    if "order" not in data:
        data["order"] = list(data["done"].keys())
    return data


def save_profile_state(path: Path, state: dict[str, Any]) -> None:
    path.write_text(json.dumps(state, indent=2), encoding="utf-8")


def crawled_product_to_record(product: CrawledProduct) -> dict[str, Any]:
    return {
        "product_id": product.product_id,
        "title": product.title,
        "handle": product.handle,
        "pdp_url": product.pdp_url,
        "description": product.description,
        "long_description": product.long_description,
        "price_inr": product.price_inr,
        "image_urls": product.image_urls,
        "features": product.features,
        "category_id": product.category_id,
        "type_label": product.type_label,
    }


def image_urls_from_disk(out_root: Path, product_id: str) -> list[str]:
    manifest = out_root / IMAGES_DIR_NAME / product_id / "images_manifest.json"
    if not manifest.exists():
        return []
    try:
        data = json.loads(manifest.read_text(encoding="utf-8"))
        urls = data.get("urls")
        return urls if isinstance(urls, list) else []
    except Exception:
        return []


def crawled_product_from_record(rec: dict[str, Any], out_root: Path) -> CrawledProduct:
    product_id = str(rec.get("product_id") or "")
    urls = rec.get("image_urls") or []
    if not urls and product_id:
        urls = image_urls_from_disk(out_root, product_id)
    return CrawledProduct(
        product_id=product_id,
        title=str(rec.get("title") or ""),
        handle=str(rec.get("handle") or ""),
        pdp_url=str(rec.get("pdp_url") or ""),
        description=str(rec.get("description") or ""),
        long_description=str(rec.get("long_description") or ""),
        price_inr=rec.get("price_inr"),
        image_urls=urls,
        features=rec.get("features") or [],
        category_id=rec.get("category_id"),
        type_label=rec.get("type_label"),
    )


def already_has_images(out_root: Path, product_id: str) -> bool:
    product_dir = out_root / IMAGES_DIR_NAME / product_id
    for folder in ("original", "detail", "large", "small"):
        d = product_dir / folder
        if d.exists() and any(d.iterdir()):
            return True
    return False


def products_from_state(state: dict[str, Any], out_root: Path) -> list[CrawledProduct]:
    done: dict[str, Any] = state.get("done") or {}
    order: list[str] = state.get("order") or list(done.keys())
    products: list[CrawledProduct] = []
    for handle in order:
        rec = done.get(handle)
        if rec:
            products.append(crawled_product_from_record(rec, out_root))
    return products


def resolve_sg_product_meta(
    profile: CollectionProfile,
    title: str,
    description: str,
    long_description: str,
    handle: str,
) -> tuple[str, str, str]:
    """Returns (category_id, product_id_prefix, type_label)."""
    blob = f"{title} {description} {long_description} {handle}".lower()
    if profile.key == "tennis-balls":
        if re.search(r"synthetic|swing\s*ball|wind\s*ball|practice\s*ball|soft\s*ball", blob):
            return ("CC_SYNTHETIC_BALLS", "SG_SYNTHETIC_BALL_", "Synthetic Cricket Ball")
        return ("CC_TENNIS_BALLS", "SG_TENNIS_BALL_", "Tennis Cricket Ball")
    return (profile.category_id, profile.product_id_prefix, profile.type_label)


def build_driver(headless: bool) -> webdriver.Chrome:
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1400,1000")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--lang=en-IN")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--allow-insecure-localhost")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    )
    # Do not wait for every asset; PDP markup is enough for parsing.
    options.page_load_strategy = "eager"
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(DEFAULT_PAGE_LOAD_TIMEOUT_S)
    driver.set_script_timeout(DEFAULT_PAGE_LOAD_TIMEOUT_S)
    return driver


def stop_page_load(driver: webdriver.Chrome) -> None:
    try:
        driver.execute_script("window.stop();")
    except Exception:
        pass


def safe_driver_get(driver: webdriver.Chrome, url: str, timeout_s: int = DEFAULT_PAGE_LOAD_TIMEOUT_S) -> bool:
    """Navigate to URL; return False on timeout but keep partial DOM when possible."""
    driver.set_page_load_timeout(timeout_s)
    try:
        driver.get(url)
        return True
    except TimeoutException:
        log.warning("Page load timeout (%ss) for %s — using partial page", timeout_s, url)
        stop_page_load(driver)
        return False


def fetch_pdp_html_http(pdp_url: str) -> str:
    try:
        resp = requests.get(
            pdp_url,
            timeout=45,
            headers={"User-Agent": "Mozilla/5.0", "Accept": "text/html"},
            verify=REQUESTS_VERIFY_SSL,
        )
        if resp.status_code == 200 and resp.text:
            return resp.text
    except Exception as ex:
        log.warning("HTTP PDP fetch failed for %s: %s", pdp_url, ex)
    return ""


def wait_for_pdp_shell(driver: webdriver.Chrome, wait_s: int = DEFAULT_PDP_WAIT_S) -> None:
    try:
        WebDriverWait(driver, wait_s).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "h1, .product__title, div.accordion-details, main, body")
            )
        )
    except TimeoutException:
        log.warning("PDP shell elements slow to appear; continuing with current DOM")


def accept_cookies_if_present(driver: webdriver.Chrome) -> None:
    for selector in (
        "button#shopify-pc__banner__btn-accept",
        "button[aria-label*='Accept']",
        "button.cookie-accept",
        "#AcceptCookies",
        "button.agree-button",
    ):
        try:
            el = driver.find_element(By.CSS_SELECTOR, selector)
            if el.is_displayed():
                el.click()
                time.sleep(0.4)
                return
        except Exception:
            continue


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", unescape(value or "")).strip()


def strip_title_marketing(title: str, profile: CollectionProfile) -> str:
    t = clean_text(title)
    for pattern in profile.title_strip_patterns:
        t = re.sub(pattern, " ", t, flags=re.I)
    t = re.sub(r"^(SG|SS)\s+", "", t, flags=re.I)
    return clean_text(t)


def seo_product_id(
    title: str,
    used: set[str],
    profile: CollectionProfile,
    *,
    id_prefix: str | None = None,
) -> str:
    core = strip_title_marketing(title, profile)
    tokens = re.findall(r"[A-Za-z0-9]+", core.upper())
    tokens = [t for t in tokens if t not in COMMON_ID_DROP]
    if not tokens:
        tokens = ["PRODUCT"]
    prefix = id_prefix or profile.product_id_prefix
    pid = prefix + "_".join(tokens)
    pid = re.sub(r"_+", "_", pid).strip("_")
    if len(pid) > MAX_PRODUCT_ID_LEN:
        pid = pid[:MAX_PRODUCT_ID_LEN].rstrip("_")
    base = pid
    n = 2
    while pid in used:
        suffix = f"_{n}"
        pid = (base[: MAX_PRODUCT_ID_LEN - len(suffix)] + suffix).rstrip("_")
        n += 1
    used.add(pid)
    return pid


def parse_inr_price(text: str) -> int | None:
    if not text:
        return None
    m = re.search(r"([\d,]+(?:\.\d+)?)", text.replace("₹", ""))
    if not m:
        return None
    try:
        return int(float(m.group(1).replace(",", "")))
    except ValueError:
        return None


def html_to_plain(html: str) -> str:
    if not html:
        return ""
    soup = BeautifulSoup(html, "lxml")
    for br in soup.find_all("br"):
        br.replace_with("\n")
    text = soup.get_text("\n", strip=True)
    return clean_text(re.sub(r"\n{3,}", "\n\n", text))


def parse_product_description_accordion(soup: BeautifulSoup) -> str:
    """
    Shopify storefront: Product Description in div.accordion-details.active when present
    (accordion__content body), not .product__description.
    """
    candidates: list[Any] = []
    candidates.extend(soup.select("div.accordion-details.active"))
    candidates.extend(soup.select("div.accordion-details"))

    for block in candidates:
        title_el = block.select_one(".accordion__title, .summary__title .accordion__title")
        title = clean_text(title_el.get_text()) if title_el else ""
        if title and not re.search(r"product\s+description", title, re.I):
            continue
        content = block.select_one(".accordion__content, [data-accordion='content']")
        if not content:
            continue
        text = html_to_plain(str(content))
        if title:
            text = re.sub(r"^product\s+description\s*", "", text, flags=re.I).strip()
        if len(text) > 30:
            return text

    # Fallback: first active accordion content if title match failed
    for block in soup.select("div.accordion-details.active"):
        content = block.select_one(".accordion__content, [data-accordion='content']")
        if content:
            text = html_to_plain(str(content))
            if len(text) > 30:
                return text
    return ""


def parse_features_from_description(text: str) -> list[dict[str, str]]:
    features: list[dict[str, str]] = []
    for line in re.split(r"[\n\r]+", text):
        line = clean_text(line)
        if not line:
            continue
        m = re.match(r"^([A-Za-z][A-Za-z0-9 /\-]{1,40})\s*:\s*(.+)$", line)
        if m:
            key = m.group(1).strip()
            val = m.group(2).strip()
            if len(key) < 50 and len(val) < 500:
                features.append({"key": key, "value": val})
    seen: set[str] = set()
    out: list[dict[str, str]] = []
    for f in features:
        k = f["key"].lower()
        if k in seen:
            continue
        seen.add(k)
        out.append(f)
    return out


def _abs_url(src: str) -> str | None:
    if not src or src.startswith("data:"):
        return None
    if src.startswith("//"):
        return "https:" + src
    if src.startswith("/"):
        return urljoin(BASE_URL, src)
    if src.startswith("http"):
        return src
    return urljoin(BASE_URL, src)


def _best_from_srcset(srcset: str) -> str | None:
    if not srcset:
        return None
    best_url = None
    best_w = -1
    for part in srcset.split(","):
        part = part.strip()
        if not part:
            continue
        bits = part.split()
        url = bits[0]
        w = 0
        if len(bits) > 1 and bits[1].endswith("w"):
            try:
                w = int(bits[1][:-1])
            except ValueError:
                w = 0
        if w >= best_w:
            best_w = w
            best_url = url
    return _abs_url(best_url) if best_url else None


def _canonical_image_key(url: str) -> str:
    parsed = urlparse(url)
    path = re.sub(r"_\d+x\d+(?=\.)", "", parsed.path)
    path = re.sub(r"_\d+x(?=\.)", "", path)
    return f"{parsed.netloc}{path}".lower()


def parse_gallery_images(html: str) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    best_by_key: dict[str, tuple[int, str]] = {}

    def score(url: str) -> int:
        m = re.search(r"width=(\d+)", url)
        return int(m.group(1)) if m else 0

    def add(url: str | None) -> None:
        if not url:
            return
        low = url.lower()
        if any(x in low for x in ("icon-", "favicon", "payment", "sprite", "logo-small")):
            return
        key = _canonical_image_key(url)
        sc = score(url)
        prev = best_by_key.get(key)
        if prev is None or sc >= prev[0]:
            best_by_key[key] = (sc, url)

    containers = []
    containers.extend(soup.select("slider-component, media-gallery, .product__media-list"))
    containers.extend(soup.select("#Slider-Gallery, #slider-Gallery, [id*='Gallery']"))
    if not containers:
        containers = [soup]

    for container in containers:
        for img in container.select("img"):
            add(_best_from_srcset(img.get("srcset") or ""))
            add(_abs_url(img.get("data-src") or ""))
            add(_abs_url(img.get("src") or ""))
        for source in container.select("source"):
            add(_best_from_srcset(source.get("srcset") or ""))

    for script in soup.select('script[type="application/ld+json"]'):
        try:
            data = json.loads(script.string or "")
        except Exception:
            continue
        items = data if isinstance(data, list) else [data]
        for item in items:
            if not isinstance(item, dict):
                continue
            images = item.get("image")
            if isinstance(images, str):
                add(_abs_url(images))
            elif isinstance(images, list):
                for im in images:
                    add(_abs_url(im if isinstance(im, str) else None))

    urls: list[str] = []
    seen: set[str] = set()
    for _k, (_sc, url) in best_by_key.items():
        key = _canonical_image_key(url)
        if key not in seen:
            seen.add(key)
            urls.append(url)
    return urls


def parse_pdp(html: str, fallback_title: str = "") -> dict[str, Any]:
    soup = BeautifulSoup(html, "lxml")
    title = fallback_title
    h1 = soup.select_one("h1.product__title, h1[class*='product'], .product-title h1, h1")
    if h1:
        title = clean_text(h1.get_text())

    price_text = ""
    for sel in (
        ".price-item--regular",
        ".price__regular .price-item",
        "span.price-item",
        ".product__price",
    ):
        el = soup.select_one(sel)
        if el and el.get_text(strip=True):
            price_text = el.get_text(" ", strip=True)
            break

    plain = parse_product_description_accordion(soup)
    if not plain:
        desc_html = ""
        for sel in (
            ".product__description.rte",
            ".product__description",
            "#tab-description",
            ".rte[class*='product']",
            "[data-product-description]",
        ):
            el = soup.select_one(sel)
            if el and el.get_text(strip=True):
                desc_html = str(el)
                break
        plain = html_to_plain(desc_html)
    if not plain:
        meta = soup.select_one('meta[name="description"]')
        if meta and meta.get("content"):
            plain = clean_text(meta["content"])

    long_desc = plain
    short_desc = plain[:255] if len(plain) > 255 else plain
    features = parse_features_from_description(plain)

    return {
        "title": title or fallback_title,
        "price_inr": parse_inr_price(price_text),
        "description": short_desc,
        "long_description": long_desc,
        "features": features,
    }


def fetch_collection_via_json(
    collection_url: str,
    limit: int | None,
) -> list[tuple[str, str, str]]:
    """Return [(handle, title, pdp_url), ...] using Shopify public JSON."""
    out: list[tuple[str, str, str]] = []
    page = 1
    while True:
        url = f"{collection_url}/products.json?limit=250&page={page}"
        resp = requests.get(
            url,
            timeout=45,
            headers={"User-Agent": "Mozilla/5.0"},
            verify=REQUESTS_VERIFY_SSL,
        )
        resp.raise_for_status()
        products = resp.json().get("products") or []
        if not products:
            break
        for p in products:
            handle = p.get("handle") or ""
            title = p.get("title") or handle
            if not handle:
                continue
            pdp = f"{BASE_URL}/products/{handle}"
            out.append((handle, title, pdp))
            if limit and len(out) >= limit:
                return out
        if len(products) < 250:
            break
        page += 1
    return out[:limit] if limit else out


def collect_collection_links_selenium(
    driver: webdriver.Chrome,
    collection_url: str,
    limit: int | None,
) -> list[tuple[str, str, str]]:
    safe_driver_get(driver, collection_url)
    accept_cookies_if_present(driver)
    wait = WebDriverWait(driver, 30)
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#product-grid, .product-grid, a[href*='/products/']")))
    except TimeoutException:
        log.warning("Collection grid slow to load; continuing")

    time.sleep(1.5)
    seen: set[str] = set()
    out: list[tuple[str, str, str]] = []

    def harvest() -> None:
        for el in driver.find_elements(By.CSS_SELECTOR, "a[href*='/products/']"):
            href = el.get_attribute("href") or ""
            if "/products/" not in href or "/collections/" in href.split("/products/")[0][-20:]:
                pass
            path = urlparse(href).path
            m = re.search(r"/products/([^/?#]+)", path)
            if not m:
                continue
            handle = m.group(1)
            key = handle.lower()
            if key in seen:
                continue
            seen.add(key)
            title = clean_text(el.get_attribute("title") or el.text or handle.replace("-", " "))
            out.append((handle, title, urljoin(BASE_URL, f"/products/{handle}")))

    harvest()
    while limit is None or len(out) < limit:
        try:
            next_btn = driver.find_element(By.CSS_SELECTOR, "a[rel='next'], .pagination__item--next a")
            if not next_btn.is_displayed():
                break
            next_btn.click()
            time.sleep(1.5)
            harvest()
        except Exception:
            break

    return out[:limit] if limit else out


def fetch_product_json(handle: str, referer: str) -> dict[str, Any] | None:
    try:
        resp = requests.get(
            f"{BASE_URL}/products/{handle}.json",
            timeout=45,
            headers={"User-Agent": "Mozilla/5.0", "Referer": referer},
            verify=REQUESTS_VERIFY_SSL,
        )
        if resp.status_code != 200:
            return None
        return resp.json().get("product")
    except Exception:
        return None


def prefer_hires(url: str) -> str:
    if "width=" in url:
        return re.sub(r"width=\d+", "width=2000", url)
    if "cdn.shopify.com" in url:
        return url + ("&width=2000" if "?" in url else "?width=2000")
    return url


def extension_from_url(url: str, content_type: str | None) -> str:
    path = urlparse(url).path.lower()
    for ext in (".jpg", ".jpeg", ".png", ".webp", ".gif"):
        if path.endswith(ext):
            return ".jpg" if ext == ".jpeg" else ext
    if content_type and "png" in content_type.lower():
        return ".png"
    if content_type and "webp" in content_type.lower():
        return ".webp"
    return ".jpg"


def download_image(session: requests.Session, url: str, dest: Path) -> bool:
    try:
        hi = prefer_hires(url)
        resp = session.get(hi, timeout=45, stream=True)
        if resp.status_code != 200:
            resp = session.get(url, timeout=45, stream=True)
        resp.raise_for_status()
        ext = extension_from_url(hi, resp.headers.get("Content-Type"))
        if dest.suffix.lower() not in (".jpg", ".jpeg", ".png", ".webp", ".gif"):
            dest = dest.with_suffix(ext)
        dest.parent.mkdir(parents=True, exist_ok=True)
        with open(dest, "wb") as f:
            for chunk in resp.iter_content(8192):
                if chunk:
                    f.write(chunk)
        if dest.stat().st_size < 400:
            dest.unlink(missing_ok=True)
            return False
        return True
    except Exception as ex:
        log.warning("Download failed %s: %s", url, ex)
        return False


def save_product_images(product_id: str, urls: list[str], out_root: Path, session: requests.Session) -> int:
    if not urls:
        return 0
    product_dir = out_root / IMAGES_DIR_NAME / product_id
    for folder in ("original", "detail", "large", "small", "other-images"):
        (product_dir / folder).mkdir(parents=True, exist_ok=True)

    saved = 0
    first = urls[0]
    digest = hashlib.md5(first.encode("utf-8")).hexdigest()[:8]
    base_name = f"{product_id}_{digest}"
    ext = extension_from_url(first, None)

    for folder in ("original", "detail", "large", "small"):
        dest = product_dir / folder / f"{base_name}{ext}"
        if download_image(session, first, dest):
            saved += 1

    for i, url in enumerate(urls[1:], start=2):
        digest = hashlib.md5(url.encode("utf-8")).hexdigest()[:8]
        dest = product_dir / "other-images" / f"{product_id}_{i:02d}_{digest}.jpg"
        if download_image(session, url, dest):
            saved += 1

    (product_dir / "images_manifest.json").write_text(
        json.dumps({"product_id": product_id, "urls": urls}, indent=2),
        encoding="utf-8",
    )
    return saved


def image_excel_paths(product_id: str, urls: list[str]) -> dict[str, str]:
    """Relative paths for import zip layout."""
    if not urls:
        return {}
    digest = hashlib.md5(urls[0].encode("utf-8")).hexdigest()[:8]
    ext = extension_from_url(urls[0], None)
    base = f"{product_id}_{digest}{ext}"
    primary = f"{product_id}/small/{base}"
    other = []
    for i, url in enumerate(urls[1:], start=2):
        d = hashlib.md5(url.encode("utf-8")).hexdigest()[:8]
        other.append(f"{product_id}/other-images/{product_id}_{i:02d}_{d}.jpg")
    return {
        "small_image": primary.replace("/small/", "/small/"),
        "medium_image": f"{product_id}/medium/{base}",
        "large_image": f"{product_id}/large/{base}",
        "original_image": f"{product_id}/original/{base}",
        "other_images": ",".join(other),
    }


def load_or_create_workbook(template: Path, output: Path) -> tuple[Any, Any, dict[str, int]]:
    if template.exists():
        wb = openpyxl.load_workbook(template)
        ws = wb["Products"] if "Products" in wb.sheetnames else wb.active
        headers = [c.value for c in ws[1]]
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Products"
        headers = list(IMPORT_HEADERS)
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)

    col = {str(h): i + 1 for i, h in enumerate(headers) if h}
    if ws.max_row > 1 and output.exists() and output.samefile(template):
        pass
    return wb, ws, col


def write_products_excel(
    products: list[CrawledProduct],
    template: Path,
    output: Path,
    profile: CollectionProfile,
    *,
    replace_rows: bool,
) -> None:
    wb, ws, col = load_or_create_workbook(template, output)
    if replace_rows and ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    start_row = 2 if replace_rows else ws.max_row + 1
    if not replace_rows and ws.max_row >= 1 and ws.max_row > 1:
        start_row = ws.max_row + 1
    elif replace_rows:
        start_row = 2

    for offset, p in enumerate(products):
        row_idx = start_row + offset
        display = strip_title_marketing(p.title, profile)
        type_label = p.type_label or profile.type_label
        category_id = p.category_id or profile.category_id
        internal = f"{BRAND} {display} {type_label}"
        product_name = internal

        img_paths = image_excel_paths(p.product_id, p.image_urls)
        sku = p.product_id.replace("SG_", "SKU-SG-", 1)
        price = p.price_inr or 0
        features_json = json.dumps(p.features, ensure_ascii=False) if p.features else ""

        values: dict[str, Any] = {
            "product_id": p.product_id,
            "sku": sku,
            "product_type_id": "FINISHED_GOOD",
            "status_id": "ACTIVE",
            "category_ids": category_id,
            "internal_name": internal,
            "brand_name": BRAND,
            "product_name": product_name,
            "description": p.description,
            "long_description": p.long_description,
            "comments": f"Crawled from {p.pdp_url}",
            "keywords": f"SG {display}, {type_label.lower()}, SG cricket, {category_id}",
            "is_virtual": "N",
            "is_variant": "N",
            "returnable": "Y",
            "taxable": "Y",
            "charge_shipping": "Y",
            "require_inventory": "Y",
            "introduction_date": "2023-01-01",
            "currency": "INR",
            "tax_rate": 18,
            "DEFAULT_PRICE": price,
            "LIST_PRICE": price,
            "MAXIMUM_PRICE": price,
            "features": features_json,
            **img_paths,
        }
        for key, val in values.items():
            if key in col:
                ws.cell(row=row_idx, column=col[key], value=val)

    for letter, width in (("A", 28), ("H", 50), ("I", 60), ("J", 60)):
        ws.column_dimensions[letter].width = width
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=8, max_col=10):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    log.info("Wrote Excel: %s (%s products)", output, len(products))


def gather_collection_links(
    driver: webdriver.Chrome,
    profile: CollectionProfile,
    limit: int | None,
) -> list[tuple[str, str, str]]:
    links: list[tuple[str, str, str]] = []
    seen_handles: set[str] = set()
    need = limit if limit is not None else 10_000

    for collection_url in profile.collection_urls:
        remaining = None if limit is None else max(0, need - len(links))
        if limit is not None and remaining == 0:
            break

        batch = collect_collection_links_selenium(driver, collection_url, limit=remaining)
        if limit is not None and (len(batch) < (remaining or 0)):
            json_batch = fetch_collection_via_json(collection_url, limit=remaining)
            batch_handles = {h for h, _, _ in batch}
            for item in json_batch:
                if item[0] not in batch_handles:
                    batch.append(item)
                if remaining is not None and len(batch) >= remaining:
                    break

        for item in batch:
            handle = item[0]
            if handle in seen_handles:
                continue
            seen_handles.add(handle)
            links.append(item)
            if limit is not None and len(links) >= limit:
                return links

    return links[:limit] if limit is not None else links


def crawl_collection(
    profile: CollectionProfile,
    driver: webdriver.Chrome,
    session: requests.Session,
    template: Path,
    limit: int | None,
    pause_s: float,
    *,
    skip_existing: bool = True,
) -> None:
    out_root = profile.out_root
    output_xlsx = profile.excel_out

    out_root.mkdir(parents=True, exist_ok=True)
    state_path = out_root / profile.state_file
    state = load_profile_state(state_path)
    done: dict[str, Any] = state.setdefault("done", {})
    order: list[str] = state.setdefault("order", [])
    used_ids: set[str] = {str(rec.get("product_id")) for rec in done.values() if rec.get("product_id")}

    log.info(
        "=== SG %s (%s) slugs=%s ===",
        profile.key,
        profile.category_id,
        ",".join(profile.collection_slugs),
    )
    all_links = gather_collection_links(driver, profile, limit=None)
    if skip_existing and done:
        pending = [link for link in all_links if link[0] not in done]
        skipped = len(all_links) - len(pending)
        if skipped:
            log.info("Skipping %s product(s) already in %s", skipped, state_path.name)
        all_links = pending
    if limit is not None:
        all_links = all_links[:limit]
    links = all_links
    log.info("Processing %s new product(s) for %s (%s total in state)", len(links), profile.key, len(done))

    new_count = 0
    failed_handles: list[str] = []
    for i, (handle, list_title, pdp_url) in enumerate(links, start=1):
        log.info("[%s/%s] %s", i, len(links), pdp_url)
        try:
            html = ""
            selenium_ok = False
            for attempt in range(1, 3):
                try:
                    safe_driver_get(driver, pdp_url)
                    accept_cookies_if_present(driver)
                    wait_for_pdp_shell(driver)
                    time.sleep(0.8)
                    html = driver.page_source or ""
                    if len(html) > 500:
                        selenium_ok = True
                        break
                except WebDriverException as ex:
                    log.warning("Selenium attempt %s/2 for %s: %s", attempt, handle, ex)
                    time.sleep(1.5)

            if not selenium_ok or len(html) < 500:
                log.info("Falling back to HTTP fetch for %s", handle)
                html = fetch_pdp_html_http(pdp_url) or html

            if len(html) < 500:
                log.error("No usable HTML for %s — skip", handle)
                failed_handles.append(handle)
                continue

            parsed = parse_pdp(html, fallback_title=list_title)

            referer = profile.collection_urls[0]
            json_product = fetch_product_json(handle, referer=referer)
            if json_product:
                if not parsed.get("long_description") and json_product.get("body_html"):
                    plain = html_to_plain(json_product["body_html"])
                    parsed["long_description"] = plain
                    parsed["description"] = plain[:255] if len(plain) > 255 else plain
                    if not parsed.get("features"):
                        parsed["features"] = parse_features_from_description(plain)
                if parsed.get("price_inr") is None and json_product.get("variants"):
                    v0 = json_product["variants"][0]
                    if v0.get("price"):
                        parsed["price_inr"] = int(float(v0["price"]) / 100)

            urls = parse_gallery_images(html)
            if not urls and json_product:
                urls = [img["src"] for img in json_product.get("images") or [] if img.get("src")]

            title = parsed.get("title") or list_title
            desc = parsed.get("description") or ""
            long_desc = parsed.get("long_description") or ""

            category_id, id_prefix, type_label = resolve_sg_product_meta(
                profile, title, desc, long_desc, handle
            )

            product_id = seo_product_id(title, used_ids, profile, id_prefix=id_prefix)

            item = CrawledProduct(
                product_id=product_id,
                title=title,
                handle=handle,
                pdp_url=pdp_url,
                description=desc,
                long_description=long_desc,
                price_inr=parsed.get("price_inr"),
                image_urls=urls,
                features=parsed.get("features") or [],
                category_id=category_id,
                type_label=type_label,
            )

            if urls:
                if skip_existing and already_has_images(out_root, product_id):
                    log.info("  images already on disk for %s; skipping download", product_id)
                else:
                    saved = save_product_images(product_id, urls, out_root, session)
                    log.info(
                        "  id=%s cat=%s images=%s saved_files=%s price=%s",
                        product_id,
                        category_id or profile.category_id,
                        len(urls),
                        saved,
                        item.price_inr,
                    )
            else:
                log.warning("  No images for %s", product_id)

            done[handle] = crawled_product_to_record(item)
            if handle not in order:
                order.append(handle)
            state["collection"] = profile.key
            save_profile_state(state_path, state)
            new_count += 1
        except WebDriverException as ex:
            log.exception("Selenium error on %s: %s", handle, ex)
            failed_handles.append(handle)
        except Exception as ex:
            log.exception("Error on %s: %s", handle, ex)
            failed_handles.append(handle)
        time.sleep(pause_s)

    all_products = products_from_state(state, out_root)
    write_products_excel(all_products, template, output_xlsx, profile, replace_rows=True)
    if failed_handles:
        fail_path = out_root / "_sg_failed_handles.json"
        fail_path.write_text(json.dumps({"failed": failed_handles}, indent=2), encoding="utf-8")
        log.warning("Failed handles (%s) logged to %s — re-run to retry (not in state)", len(failed_handles), fail_path)
    log.info(
        "Finished %s: %s new this run, %s total in state -> %s | images under %s",
        profile.key,
        new_count,
        len(done),
        output_xlsx,
        out_root / IMAGES_DIR_NAME,
    )


def crawl_collections(
    profiles: list[CollectionProfile],
    template: Path,
    headless: bool,
    limit: int | None,
    pause_s: float,
    *,
    skip_existing: bool = True,
) -> None:
    session = requests.Session()
    session.verify = REQUESTS_VERIFY_SSL
    session.headers.update({"User-Agent": "Mozilla/5.0", "Referer": BASE_URL + "/"})
    driver = build_driver(headless=headless)
    try:
        for profile in profiles:
            crawl_collection(
                profile,
                driver,
                session,
                template,
                limit,
                pause_s,
                skip_existing=skip_existing,
            )
    finally:
        driver.quit()


def parse_collection_keys(raw: str | None) -> list[str]:
    if not raw or raw.strip().lower() == "all":
        return list(COLLECTION_PROFILES.keys())
    keys = [k.strip() for k in raw.split(",") if k.strip()]
    unknown = [k for k in keys if k not in COLLECTION_PROFILES]
    if unknown:
        raise SystemExit(f"Unknown collection(s): {unknown}. Choose from: {list(COLLECTION_PROFILES)}")
    return keys


def main() -> None:
    global DEFAULT_PAGE_LOAD_TIMEOUT_S

    parser = argparse.ArgumentParser(description="SG shop.teamsg.in product crawler (multi-collection)")
    parser.add_argument("--template", type=Path, default=TEMPLATE_XLSX)
    parser.add_argument("--headless", action="store_true")
    parser.add_argument(
        "--collections",
        type=str,
        default=",".join(DEFAULT_COLLECTION_KEYS),
        help=(
            "Comma-separated collection keys or 'all'. "
            f"Keys: {', '.join(COLLECTION_PROFILES.keys())}. "
            f"Default: {','.join(DEFAULT_COLLECTION_KEYS)}"
        ),
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=3,
        help="Max products per category (default 3 for review). Ignored when --all is set.",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Crawl every product in each selected collection",
    )
    parser.add_argument("--pause", type=float, default=1.0)
    parser.add_argument(
        "--force",
        action="store_true",
        help="Re-crawl products even if already recorded in the category state JSON",
    )
    parser.add_argument(
        "--page-timeout",
        type=int,
        default=DEFAULT_PAGE_LOAD_TIMEOUT_S,
        help=f"Chrome page load timeout seconds (default {DEFAULT_PAGE_LOAD_TIMEOUT_S})",
    )
    args = parser.parse_args()
    DEFAULT_PAGE_LOAD_TIMEOUT_S = max(20, args.page_timeout)

    product_limit: int | None = None if args.all else max(1, args.limit)
    keys = parse_collection_keys(args.collections)
    profiles = [COLLECTION_PROFILES[k] for k in keys]
    SG_ROOT.mkdir(parents=True, exist_ok=True)

    crawl_collections(
        profiles=profiles,
        template=args.template,
        headless=args.headless,
        limit=product_limit,
        pause_s=args.pause,
        skip_existing=not args.force,
    )


if __name__ == "__main__":
    main()
