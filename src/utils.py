from datetime import datetime, timedelta
import os
import re
import sys
import unicodedata


HEADER_ALIASES = {
    "no": {
        "no", "no.", "번호", "순번", "연번", "num", "number",
    },
    "name": {
        "식품명", "품목명", "제품명", "상품명", "품명", "식재료명", "재료명", "물품명",
        "명칭", "item", "itemname", "product", "productname", "name",
    },
    "spec": {
        "규격", "용량", "중량", "규격용량", "용량규격", "용량중량", "포장단위",
        "사이즈", "크기", "spec", "standard", "size", "weight", "capacity",
    },
    "unit": {
        "단위", "입수", "포장", "unit", "uom",
    },
    "qty": {
        "수량", "주문수량", "발주수량", "신청수량", "청구량", "납품수량",
        "주문량", "발주량", "합계수량", "수량합계", "qty", "quantity", "count", "개수", "갯수",
    },
    "price": {
        "단가", "가격", "공급가", "납품단가", "매입단가", "판매가", "단가원",
        "price", "unitprice", "amountprice",
    },
    "category": {
        "분류", "카테고리", "구분", "유형", "대분류", "중분류", "소분류",
        "category", "type",
    },
}


CATEGORY_KEYWORDS = {
    "유제품": ["우유", "치즈", "버터", "요거트", "요구르트", "생크림", "휘핑", "연유", "크림치즈"],
    "육류": ["소고기", "쇠고기", "돼지", "돈육", "닭", "계육", "오리", "양고기", "베이컨", "햄"],
    "수산": ["생선", "새우", "오징어", "문어", "연어", "고등어", "참치", "멸치", "조개", "홍합"],
    "채소": ["양파", "대파", "마늘", "당근", "감자", "고구마", "상추", "배추", "무", "오이", "호박", "버섯"],
    "과일": ["사과", "배", "바나나", "딸기", "포도", "레몬", "라임", "오렌지", "파인애플", "토마토"],
    "양념/소스": ["소스", "간장", "고추장", "된장", "식초", "설탕", "소금", "후추", "고춧가루", "드레싱", "마요"],
    "곡류": ["쌀", "밀가루", "전분", "빵가루", "면", "파스타", "또띠아", "누룽지"],
    "음료": ["주스", "음료", "탄산", "콜라", "사이다", "커피", "차", "생수"],
}

def parse_date_from_sheet_name(sheet_name, fallback_year=None):
    """
    Parses date from sheet name (e.g., '1.1', '12.25').
    Assumes current year or handles year logic if needed.
    Returns datetime object.
    """
    try:
        text = unicodedata.normalize("NFKC", str(sheet_name))
        year = extract_year_from_text(text) or fallback_year or datetime.now().year

        full_date = re.search(r"20\d{2}\D+(\d{1,2})\D+(\d{1,2})", text)
        if full_date:
            return datetime(year, int(full_date.group(1)), int(full_date.group(2)))

        month_day = re.search(r"(?<!\d)(\d{1,2})\s*(?:[./-]|월)\s*(\d{1,2})\s*(?:일)?(?!\d)", text)
        if month_day:
            return datetime(year, int(month_day.group(1)), int(month_day.group(2)))
    except Exception as e:
        print(f"Error parsing date from {sheet_name}: {e}")
        return None
    return None


def extract_year_from_text(text):
    if not text:
        return None
    match = re.search(r"(20\d{2})", str(text))
    if not match:
        return None
    try:
        return int(match.group(1))
    except ValueError:
        return None

def get_sending_date(receiving_date):
    """
    Returns receiving_date - 1 day.
    """
    if receiving_date:
        return receiving_date - timedelta(days=1)
    return None

def get_resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def normalize_string(s):
    """
    Removes spaces, hidden characters, lowercases, and handles basic normalization.
    """
    if not s:
        return ""
    s = str(s).lower()
    s = unicodedata.normalize('NFKC', s) # Normalize chars
    # Remove all whitespace
    return "".join(s.split())


def normalize_header(value):
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).lower()
    text = re.sub(r"\([^)]*\)", "", text)
    text = re.sub(r"[\s\-_/:·.,()\[\]{}<>]+", "", text)
    text = text.replace("₩", "").replace("원", "원")
    return text


def header_field(value):
    normalized = normalize_header(value)
    if not normalized:
        return None
    for field, aliases in HEADER_ALIASES.items():
        normalized_aliases = {normalize_header(alias) for alias in aliases}
        if normalized in normalized_aliases:
            return field
    return None


def normalize_name(value):
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).lower()
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"[\[\]{}<>]", " ", text)
    text = re.sub(r"[^0-9a-z가-힣]+", " ", text)
    return " ".join(text.split())


def compact_name(value):
    return re.sub(r"\s+", "", normalize_name(value))


def tokenize_name(value):
    text = normalize_name(value)
    if not text:
        return set()

    tokens = set(text.split())
    compact = compact_name(value)
    if compact:
        tokens.add(compact)

    # Extract useful mixed tokens while keeping short Korean ingredient names.
    for token in re.findall(r"[가-힣]{2,}|[a-z]{2,}|\d+[a-z가-힣]*", text):
        tokens.add(token)
    return {token for token in tokens if token}


def parse_number(value):
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    text = unicodedata.normalize("NFKC", str(value))
    text = text.replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    number = float(match.group(0))
    if number.is_integer():
        return int(number)
    return number


def normalize_spec(value):
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).lower()
    text = text.replace("×", "x").replace("*", "x")
    text = re.sub(r"\s+", "", text)
    return text


def spec_tokens(value):
    text = normalize_spec(value)
    if not text:
        return set()
    return {token for token in re.split(r"[,/|+·;_\-()\[\]\s]+", text) if token}


def _unit_to_base(unit):
    unit = unicodedata.normalize("NFKC", unit).lower()
    if unit in {"kg", "㎏", "킬로", "키로"}:
        return "g", 1000.0
    if unit in {"g", "그램", "gr"}:
        return "g", 1.0
    if unit in {"l", "ℓ", "리터", "liter"}:
        return "ml", 1000.0
    if unit in {"ml", "㎖", "미리", "미리리터"}:
        return "ml", 1.0
    if unit in {"개", "ea", "입", "봉", "팩", "포", "박스", "box"}:
        return "count", 1.0
    return unit, 1.0


def extract_quantity(value):
    """
    Return a normalized quantity dict such as {"amount": 1000.0, "unit": "g"}.
    The parser is intentionally conservative: it compares only compatible base units.
    """
    if value is None:
        return None
    text = unicodedata.normalize("NFKC", str(value)).lower()
    text = text.replace("×", "x").replace("*", "x")
    pattern = re.compile(
        r"(\d+(?:\.\d+)?)\s*(kg|㎏|킬로|키로|g|그램|gr|l|ℓ|리터|liter|ml|㎖|미리|미리리터|개|ea|입|봉|팩|포|박스|box)"
    )
    matches = list(pattern.finditer(text))
    if not matches:
        return None

    first = matches[0]
    amount = float(first.group(1))
    base_unit, multiplier = _unit_to_base(first.group(2))
    amount *= multiplier

    tail = text[first.end(): first.end() + 12]
    multiplier_match = re.search(r"x\s*(\d+(?:\.\d+)?)", tail)
    if multiplier_match and base_unit in {"g", "ml"}:
        amount *= float(multiplier_match.group(1))

    return {"amount": amount, "unit": base_unit}


def quantities_equal(left, right, tolerance=0.001):
    if not left or not right:
        return False
    if left.get("unit") != right.get("unit"):
        return False
    return abs(float(left["amount"]) - float(right["amount"])) <= tolerance


def infer_category(name=None, spec=None, explicit=None, section=None):
    if explicit:
        return str(explicit).strip()
    section_text = normalize_name(section)
    if "유제품" in section_text or "dairy" in section_text:
        return "유제품"

    haystack = normalize_name(f"{name or ''} {spec or ''}")
    for category, keywords in CATEGORY_KEYWORDS.items():
        if any(keyword in haystack for keyword in keywords):
            return category
    return "기타"
