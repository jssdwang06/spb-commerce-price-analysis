import argparse
import json
import sys
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


ROOT = Path(__file__).resolve().parent.parent  # project root
SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import etagi_commerce_scraper as scraper  # noqa: E402


LOCAL_VALUE_FIXES = {
    "Ремонт": {
        "9": "Требует ремонта / Частичный ремонт",
        "10": "Черновая отделка",
        "11": "Улучшенная черновая отделка",
        "13": "Предчистовая отделка",
        "19": "С ремонтом / Чистовая отделка",
        "24": "Дизайн / Евро / Современный ремонт",
    },
    "Расположение": {
        "1": "БЦ / ТЦ / ТРК",
        "2": "В нежилом здании",
        "3": "Отдельностоящее здание целиком",
        "4": "В жилом здании",
        "5": "Пристрой в жилом доме",
    },
}

EXCLUDED_DERIVED_FIELDS = set()

OPTIONAL_FEATURE_FIELDS = {
    "Кондиционер",
    "Видеонаблюдение",
    "Пожарная сигнализация",
    "Мебель",
}

OBJECT_TYPE_LABELS = {
    18: "office",
    19: "retail",
    20: "free_purpose_premise",
    23: "free_purpose",
}

ACTION_LABELS = {
    "lease": "rent",
    "sell": "sale",
}

LAYOUT_LABELS = {
    1: "cabinet",
    2: "open",
    3: "mixed",
}

REPAIR_LABELS = {
    2: "cosmetic_repair",
    6: "modern_repair",
    7: "euro_repair",
    9: "needs_repair",
    10: "rough_finish",
    11: "improved_rough_finish",
    13: "pre_finish",
    15: "with_finishing",
    19: "with_repair",
    24: "design_euro_modern_repair",
}

LINE_LABELS = {
    1: "first_line",
    2: "second_line",
    3: "third_line",
}

LOCATION_LABELS = {
    1: "business_or_shopping_center",
    2: "non_residential_building",
    3: "standalone_building",
    4: "residential_building",
    5: "annex_to_residential_building",
}

HEAT_SUPPLY_LABELS = {
    1: "central",
    2: "gas",
    3: "electric",
    4: "solid_fuel",
    5: "absent",
}

SEWERAGE_LABELS = {
    1: "inside_premise",
    2: "shared_in_building",
    3: "on_floor",
}

PARKING_LABELS = {
    1: "public",
    2: "private",
    3: "inside_building",
}

WATER_SOURCE_LABELS = {
    1: "central",
    2: "autonomous",
}

MAPPED_RAW_FIELDS = {
    "action_sl": ACTION_LABELS,
    "object_type_id": OBJECT_TYPE_LABELS,
    "keep": REPAIR_LABELS,
    "line": LINE_LABELS,
    "location": LOCATION_LABELS,
    "additional_meta.location": LOCATION_LABELS,
    "additional_meta.layout": LAYOUT_LABELS,
    "additional_meta.heatSupply": HEAT_SUPPLY_LABELS,
    "additional_meta.sewerage": SEWERAGE_LABELS,
    "sewerage": SEWERAGE_LABELS,
    "additional_meta.parkings": PARKING_LABELS,
    "additional_meta.water_sources": WATER_SOURCE_LABELS,
    "waterSource": WATER_SOURCE_LABELS,
}

CARD_FIELD_GROUPS = [
    (
        "",
        [
            ("Код объекта", ("object_id",)),
            ("Этаж / Этажность", ("Этаж / Этажность",)),
            ("Общая площадь", ("Общая площадь", "Площадь")),
            ("Количество входов", ("Количество входов",)),
            ("Наличие вытяжки", ("Наличие вытяжки", "Вентиляция")),
            ("Арендатор", ("Арендатор",)),
            ("Окупаемость", ("Окупаемость",)),
            ("Планировка", ("Планировка",)),
            ("Ремонт", ("Ремонт",)),
        ],
    ),
    (
        "Коммуникации",
        [
            ("Отопление", ("Отопление",)),
        ],
    ),
    (
        "О здании",
        [
            ("Линия", ("Линия",)),
            ("Расположение", ("Расположение",)),
            ("Год постройки", ("Год постройки",)),
            ("Стены", ("Стены",)),
            ("Парковка", ("Парковка",)),
            ("Кондиционер", ("Кондиционер",)),
            ("Видеонаблюдение", ("Видеонаблюдение",)),
            ("Пожарная сигнализация", ("Пожарная сигнализация",)),
            ("Мебель", ("Мебель",)),
            ("Канализация / cанузлы", ("Канализация / cанузлы",)),
        ],
    ),
]

TABLE_BASE_COLUMNS = [
    "Код объекта",
    "url",
    "Тип",
    "Адрес",
    "Цена",
    "Цена за м²",
    "Метро",
    "Метро расстояние",
    "Метро время, мин",
]


def load_cached_data(path: Path) -> dict[str, Any]:
    return scraper.load_page_data(path.read_text(encoding="utf-8"))


def clean_number(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return scraper.clean_text(value)


def metro_distance(obj: dict[str, Any]) -> str:
    stations = obj.get("metro_stations") or []
    parts = []
    for station in stations:
        name = scraper.clean_text(station.get("stationName") or station.get("name"))
        name = name.removeprefix("метро ").strip()
        distance = clean_number(station.get("distance"))
        if name and distance:
            parts.append(f"{name}: {distance} м")
    if parts:
        return "; ".join(parts)

    distance = clean_number(obj.get("metro_distance"))
    return f"{distance} м" if distance else ""


def metro_time(obj: dict[str, Any]) -> str:
    return clean_number(obj.get("time_to_metro"))


def add_metro_details(row: dict[str, Any], obj: dict[str, Any]) -> None:
    distance = metro_distance(obj)
    if distance:
        row["Метро расстояние"] = distance
    time = metro_time(obj)
    if time:
        row["Метро время, мин"] = time


def compact_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), default=str)


def nested_get(obj: dict[str, Any], path: str) -> Any:
    current: Any = obj
    for part in path.split("."):
        if not isinstance(current, dict):
            return None
        current = current.get(part)
    return current


def as_int(value: Any) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def bool_label(value: Any) -> str:
    if isinstance(value, bool):
        return "yes" if value else "no"
    if isinstance(value, (int, float)) and value in {0, 1}:
        return "yes" if value == 1 else "no"
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"true", "1", "yes", "да"}:
            return "yes"
        if lowered in {"false", "0", "no", "нет"}:
            return "no"
    return scraper.clean_text(value)


def map_code(value: Any, mapping: dict[Any, str]) -> str:
    if isinstance(value, list):
        return "; ".join(map_code(item, mapping) for item in value if item not in (None, ""))
    int_value = as_int(value)
    key = int_value if int_value is not None else scraper.clean_text(value)
    return mapping.get(key, scraper.clean_text(value))


def rich_text_from_plate_notes(value: Any) -> str:
    if not isinstance(value, list):
        return scraper.clean_text(value)
    parts: list[str] = []
    for block in value:
        if not isinstance(block, dict):
            text = scraper.clean_text(block)
            if text:
                parts.append(text)
            continue
        for child in block.get("children") or []:
            if isinstance(child, dict):
                text = scraper.clean_text(child.get("text"))
                if text:
                    parts.append(text)
    return "\n".join(parts)


def labels_from_list(value: Any, *keys: str) -> str:
    if not isinstance(value, list):
        return scraper.clean_text(value)
    labels: list[str] = []
    for item in value:
        if not isinstance(item, dict):
            label = scraper.clean_text(item)
        else:
            label = ""
            for key in keys:
                label = scraper.clean_text(item.get(key))
                if label:
                    break
        if label and label not in labels:
            labels.append(label)
    return "; ".join(labels)


def normalize_nested_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return bool_label(value)
    if isinstance(value, (str, int, float)):
        return scraper.clean_text(value)
    if isinstance(value, list):
        if all(not isinstance(item, (dict, list)) for item in value):
            return "; ".join(scraper.clean_text(item) for item in value if scraper.clean_text(item))
        return compact_json(value)
    if isinstance(value, dict):
        return compact_json(value)
    return scraper.clean_text(value)


def flatten_object_fields(obj: dict[str, Any], prefix: str = "") -> dict[str, str]:
    result: dict[str, str] = {}
    for key, value in obj.items():
        path = f"{prefix}.{key}" if prefix else str(key)
        if isinstance(value, dict):
            result.update(flatten_object_fields(value, path))
            continue
        normalized = normalize_nested_value(value)
        if normalized:
            result[f"json.{path}"] = normalized
    return result


def enhanced_object_fields(obj: dict[str, Any]) -> dict[str, str]:
    fields = flatten_object_fields(obj)

    notes_text = rich_text_from_plate_notes(obj.get("plate_notes"))
    if notes_text:
        fields["json.plate_notes_text"] = notes_text

    suitable_for = labels_from_list(obj.get("officeSuitableFor"), "label", "name", "title")
    if suitable_for:
        fields["json.officeSuitableFor.labels"] = suitable_for

    metro_stations = labels_from_list(obj.get("metro_stations"), "stationName", "name")
    if metro_stations:
        fields["json.metro_stations.labels"] = metro_stations

    lat = scraper.clean_text(obj.get("la"))
    lon = scraper.clean_text(obj.get("lo"))
    if lat and lon:
        fields["json.coordinates"] = f"{lat}, {lon}"

    for path, mapping in MAPPED_RAW_FIELDS.items():
        value = nested_get(obj, path)
        mapped = map_code(value, mapping)
        if mapped:
            fields[f"json.{path}.label"] = mapped

    for path in (
        "ventilation",
        "vent",
        "conditioner",
        "fireAlarm",
        "remoteSecurity",
        "video",
        "separate_entrance",
        "freestanding",
        "is_exclusive",
        "online_showing",
        "additional_meta.private_entrance",
        "additional_meta.ventilation",
        "additional_meta.vent",
        "additional_meta.conditioner",
        "additional_meta.fireAlarm",
        "additional_meta.video",
        "additional_meta.furniture",
        "additional_meta.upload_ramp",
        "additional_meta.fencing",
        "additional_meta.gas",
        "additional_meta.hasCraneBeam",
        "additional_meta.isStoreRoom",
        "additional_meta.construction_on_the_ground",
    ):
        value = nested_get(obj, path)
        label = bool_label(value)
        if label:
            fields[f"json.{path}.label"] = label

    return fields


def cached_detail_files(detail_dir: Path) -> dict[str, Path]:
    files: dict[str, Path] = {}
    for path in sorted(detail_dir.glob("*.html")):
        object_id = path.stem.split("_", 1)[0]
        if object_id and object_id not in files:
            files[object_id] = path
    return files


def list_items_from_cache(list_dir: Path) -> list[tuple[dict[str, Any], dict[str, Any], Path]]:
    rows: list[tuple[dict[str, Any], dict[str, Any], Path]] = []
    for path in sorted(list_dir.glob("*.html")):
        data = load_cached_data(path)
        items = data.get("lists", {}).get("commerce") or []
        rows.extend((item, data, path) for item in items)
    return rows


def unique_items(
    items: list[tuple[dict[str, Any], dict[str, Any], Path]]
) -> list[tuple[dict[str, Any], dict[str, Any], Path]]:
    seen: set[str] = set()
    result: list[tuple[dict[str, Any], dict[str, Any], Path]] = []
    for item, data, path in items:
        object_id = scraper.object_id(item)
        key = str(object_id or "")
        if key and key not in seen:
            seen.add(key)
            result.append((item, data, path))
    return result


def detail_characteristics(detail_path: Path) -> tuple[dict[str, Any], dict[str, str]]:
    html = detail_path.read_text(encoding="utf-8")
    data = scraper.load_page_data(html)
    characteristics = scraper.extract_characteristics(html)
    detail_obj = data.get("objects", {}).get("commerceObject") or {}
    for key, value in scraper.detail_fields_from_object(detail_obj).items():
        if key in EXCLUDED_DERIVED_FIELDS:
            continue
        if not characteristics.get(key):
            characteristics[key] = value
    return data, characteristics


def normalize_values(row: dict[str, Any]) -> None:
    for column, mapping in LOCAL_VALUE_FIXES.items():
        value = row.get(column)
        if value is None:
            continue
        key = str(value).strip()
        if key in mapping:
            row[column] = mapping[key]

    repair = row.get("Ремонт")
    if isinstance(repair, str):
        for prefix in (
            "С ремонтом / ",
            "Требует ремонта / ",
            "Дизайн / Евро / ",
        ):
            if repair.startswith(prefix):
                row["Ремонт"] = repair[len(prefix) :].strip()
                break

    # These are optional amenities on the listing page. A false JSON default
    # means the amenity is not shown, rather than an explicit "Нет" value.
    for column in OPTIONAL_FEATURE_FIELDS:
        if scraper.clean_text(row.get(column)).lower() in {"нет", "no"}:
            row[column] = ""


def build_rows(cache_dir: Path) -> list[dict[str, Any]]:
    list_dir = cache_dir / "list"
    detail_dir = cache_dir / "detail"
    if not list_dir.exists():
        raise FileNotFoundError(f"List cache directory not found: {list_dir}")
    if not detail_dir.exists():
        raise FileNotFoundError(f"Detail cache directory not found: {detail_dir}")

    detail_files = cached_detail_files(detail_dir)
    rows: list[dict[str, Any]] = []

    for item, list_data, _list_path in unique_items(list_items_from_cache(list_dir)):
        row = scraper.base_row(list_data, item)
        add_metro_details(row, item)
        object_id = str(row.get("object_id") or "")
        detail_path = detail_files.get(object_id)

        if not detail_path:
            row["detail_error"] = "detail cache missing"
            rows.append(row)
            continue

        try:
            detail_data, characteristics = detail_characteristics(detail_path)
            detail_obj = detail_data.get("objects", {}).get("commerceObject") or {}
            if detail_obj:
                row.update(scraper.base_row(detail_data, detail_obj))
                add_metro_details(row, detail_obj)
                row.update(enhanced_object_fields(detail_obj))
            row.update(characteristics)
            normalize_values(row)
        except Exception as exc:
            row["detail_error"] = scraper.clean_text(exc)

        rows.append(row)

    return rows


def clean_export_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return ILLEGAL_CHARACTERS_RE.sub("", scraper.clean_text(value))


def clean_export_header(value: Any) -> str:
    return clean_export_value(value)


def card_value(row: dict[str, Any], keys: tuple[str, ...]) -> str:
    for key in keys:
        value = clean_export_value(row.get(key))
        if value:
            return value
    return ""


def flatten_card_columns() -> list[tuple[str, tuple[str, ...]]]:
    columns: list[tuple[str, tuple[str, ...]]] = []
    for _group_title, fields in CARD_FIELD_GROUPS:
        columns.extend(fields)
    return columns


def save_formatted_excel(rows: list[dict[str, Any]], path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "commerce"
    card_columns = flatten_card_columns()
    used_source_keys = {"object_id"}
    for _label, keys in card_columns:
        used_source_keys.update(keys)

    headers = TABLE_BASE_COLUMNS + [label for label, _keys in card_columns if label != "Код объекта"]
    for row in rows:
        for key in row:
            if key not in used_source_keys and key not in {"detail_error"} and key not in headers:
                headers.append(key)
    if any(row.get("detail_error") for row in rows):
        headers.append("detail_error")

    ws.append([clean_export_header(header) for header in headers])

    source_by_header = {
        "Код объекта": ("object_id",),
        "url": ("url",),
        "Тип": ("Тип",),
        "Адрес": ("Адрес",),
        "Цена": ("Цена",),
        "Цена за м²": ("Цена за м²",),
        "Метро": ("Метро",),
        "Метро расстояние": ("Метро расстояние",),
        "Метро время, мин": ("Метро время, мин",),
    }
    source_by_header.update({label: keys for label, keys in card_columns})

    for row in rows:
        ws.append([card_value(row, source_by_header.get(header, (header,))) for header in headers])

    wb.save(path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract local cached Etagi commerce HTML pages to Excel."
    )
    parser.add_argument(
        "--cache-dir",
        default=str(ROOT / "raw" / "prodazha_196" / "html_cache"),
        help="Directory with list/ and detail/ cached HTML folders.",
    )
    parser.add_argument(
        "--output",
        default=str(ROOT / "data" / "Купить.xlsx"),
        help="Output .xlsx file.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    rows = build_rows(Path(args.cache_dir))
    if not rows:
        raise RuntimeError("No rows extracted from local cache.")
    save_formatted_excel(rows, args.output)

    missing_details = sum(1 for row in rows if row.get("detail_error"))
    print(f"Saved {len(rows)} rows to {args.output}")
    if missing_details:
        print(f"Rows with detail_error: {missing_details}")


if __name__ == "__main__":
    main()
