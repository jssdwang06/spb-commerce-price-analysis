import re
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent.parent  # project root
INPUT = ROOT / "raw" / "arenda_325" / "etagi_arenda_325_selected.xlsx"
OUTPUT_XLSX = ROOT / "raw" / "arenda_325" / "etagi_arenda_325_analysis.xlsx"
OUTPUT_TXT = ROOT / "raw" / "arenda_325" / "etagi_arenda_325_analysis.txt"


def number(value):
    if pd.isna(value):
        return pd.NA
    text = str(value).replace("\xa0", " ").replace(",", ".")
    match = re.search(r"-?\d+(?:[ .]\d{3})*(?:\.\d+)?|-?\d+(?:\.\d+)?", text)
    if not match:
        return pd.NA
    return float(match.group(0).replace(" ", ""))


def first_metro_distance(value):
    if pd.isna(value):
        return pd.NA
    match = re.search(r"\((\d+(?:[,.]\d+)?)\s*км\)", str(value))
    if not match:
        return pd.NA
    return float(match.group(1).replace(",", "."))


def first_metro_time(value):
    if pd.isna(value):
        return pd.NA
    match = re.search(r"(\d+)\s*мин", str(value))
    if not match:
        return pd.NA
    return float(match.group(1))


def parse_floor(value):
    if pd.isna(value):
        return pd.NA
    match = re.search(r"(-?\d+)\s+из\s+(\d+)", str(value))
    if not match:
        return pd.NA
    return float(match.group(1))


def parse_floors(value):
    if pd.isna(value):
        return pd.NA
    match = re.search(r"(-?\d+)\s+из\s+(\d+)", str(value))
    if not match:
        return pd.NA
    return float(match.group(2))


def clean_data(df):
    df = df.copy()
    df["price"] = df["Цена"].map(number).astype("Float64")
    df["old_price"] = df["Старая цена"].map(number).astype("Float64")
    df["price_m2"] = df["Цена за м²"].map(number).astype("Float64")
    df["area"] = df["Общая площадь"].map(number).astype("Float64")
    df["center_km"] = df["До центра"].map(number).astype("Float64")
    df["metro_km"] = df["Метро"].map(first_metro_distance).astype("Float64")
    df["metro_min"] = df["Метро"].map(first_metro_time).astype("Float64")
    df["power_kw"] = df["Мощность электричества"].map(number).astype("Float64")
    df["ceiling_m"] = df["Высота потолков"].map(number).astype("Float64")
    df["year"] = df["Год постройки"].map(number).astype("Float64")
    df["entrances"] = df["Количество входов"].map(number).astype("Float64")
    df["floor"] = df["Этаж / Этажность"].map(parse_floor).astype("Float64")
    df["floors"] = df["Этаж / Этажность"].map(parse_floors).astype("Float64")
    df["discount_pct"] = (df["old_price"] - df["price"]) / df["old_price"] * 100
    return df


def by_group(df, group_col):
    return (
        df.groupby(group_col, dropna=False)
        .agg(
            count=("Код объекта", "count"),
            price_median=("price", "median"),
            price_m2_median=("price_m2", "median"),
            area_median=("area", "median"),
            metro_km_median=("metro_km", "median"),
        )
        .sort_values("count", ascending=False)
        .reset_index()
    )


def numeric_summary(df):
    rows = []
    for col, label in [
        ("price", "Цена, руб./мес."),
        ("price_m2", "Цена за м², руб./мес."),
        ("area", "Площадь, м²"),
        ("center_km", "До центра, км"),
        ("metro_km", "До метро, км"),
        ("metro_min", "До метро, мин."),
        ("power_kw", "Мощность, кВт"),
        ("ceiling_m", "Высота потолков, м"),
        ("year", "Год постройки"),
    ]:
        s = df[col].dropna()
        rows.append(
            {
                "metric": label,
                "count": int(s.count()),
                "min": round(float(s.min()), 2) if len(s) else None,
                "q25": round(float(s.quantile(0.25)), 2) if len(s) else None,
                "median": round(float(s.median()), 2) if len(s) else None,
                "mean": round(float(s.mean()), 2) if len(s) else None,
                "q75": round(float(s.quantile(0.75)), 2) if len(s) else None,
                "max": round(float(s.max()), 2) if len(s) else None,
            }
        )
    return pd.DataFrame(rows)


def top_rows(df, sort_col, ascending=True, n=15):
    cols = [
        "Код объекта",
        "Тип",
        "Адрес",
        "Цена",
        "Цена за м²",
        "Общая площадь",
        "Метро",
        "Ремонт",
        "Расположение",
        "url",
    ]
    return df.sort_values(sort_col, ascending=ascending)[cols].head(n)


def discount_rows(df, n=15):
    cols = [
        "Код объекта",
        "Тип",
        "Адрес",
        "Цена",
        "Старая цена",
        "Цена за м²",
        "discount_pct",
        "url",
    ]
    return df[df["discount_pct"].notna()].sort_values("discount_pct", ascending=False)[cols].head(n)


def value_candidates(df, n=25):
    cols = [
        "Код объекта",
        "Тип",
        "Адрес",
        "Цена",
        "Цена за м²",
        "Общая площадь",
        "Метро",
        "Ремонт",
        "Расположение",
        "value_score",
        "url",
    ]
    type_median = df.groupby("Тип")["price_m2"].transform("median")
    work = df.copy()
    work["value_score"] = work["price_m2"] / type_median
    mask = (
        work["price_m2"].notna()
        & (work["price_m2"] >= 500)
        & work["area"].notna()
        & (work["area"] >= 30)
        & (work["area"] <= 500)
        & (work["value_score"] <= 0.75)
    )
    near_metro = work["metro_km"].isna() | (work["metro_km"] <= 1.5)
    return work[mask & near_metro].sort_values(["value_score", "price_m2"]).head(n)[cols]


def coverage(df):
    rows = []
    for col in df.columns:
        if col in {
            "price",
            "old_price",
            "price_m2",
            "area",
            "center_km",
            "metro_km",
            "metro_min",
            "power_kw",
            "ceiling_m",
            "year",
            "entrances",
            "floor",
            "floors",
            "discount_pct",
        }:
            continue
        filled = df[col].notna() & (df[col].astype(str).str.strip() != "")
        rows.append(
            {
                "column": col,
                "filled": int(filled.sum()),
                "missing": int((~filled).sum()),
                "filled_pct": round(float(filled.mean() * 100), 1),
            }
        )
    return pd.DataFrame(rows).sort_values(["filled_pct", "filled"], ascending=[True, True])


def format_money(value):
    if pd.isna(value):
        return ""
    return f"{float(value):,.0f}".replace(",", " ")


def write_report(df, summary, by_type, by_location, by_repair):
    lines = []
    lines.append(f"Objects: {len(df)}")
    lines.append(f"Median rent: {format_money(df['price'].median())} руб./мес.")
    lines.append(f"Median price per m2: {format_money(df['price_m2'].median())} руб./мес. за м²")
    lines.append(f"Median area: {df['area'].median():.1f} м²")
    lines.append(f"Median metro distance: {df['metro_km'].median():.1f} км")
    lines.append(f"Median distance to center: {df['center_km'].median():.1f} км")
    lines.append("")
    lines.append("By type:")
    lines.append(by_type.to_string(index=False))
    lines.append("")
    lines.append("By location:")
    lines.append(by_location.to_string(index=False))
    lines.append("")
    lines.append("By repair:")
    lines.append(by_repair.to_string(index=False))
    OUTPUT_TXT.write_text("\n".join(lines), encoding="utf-8")


def main():
    df = pd.read_excel(INPUT)
    df = clean_data(df)

    summary = numeric_summary(df)
    by_type = by_group(df, "Тип")
    by_location = by_group(df, "Расположение")
    by_repair = by_group(df, "Ремонт")
    by_layout = by_group(df, "Планировка")
    by_parking = by_group(df, "Парковка")

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        by_type.to_excel(writer, sheet_name="by_type", index=False)
        by_location.to_excel(writer, sheet_name="by_location", index=False)
        by_repair.to_excel(writer, sheet_name="by_repair", index=False)
        by_layout.to_excel(writer, sheet_name="by_layout", index=False)
        by_parking.to_excel(writer, sheet_name="by_parking", index=False)
        top_rows(df, "price_m2", True).to_excel(writer, sheet_name="cheap_m2", index=False)
        top_rows(df, "price_m2", False).to_excel(writer, sheet_name="expensive_m2", index=False)
        top_rows(df, "price", True).to_excel(writer, sheet_name="cheap_total", index=False)
        top_rows(df, "price", False).to_excel(writer, sheet_name="expensive_total", index=False)
        discount_rows(df).to_excel(writer, sheet_name="discounts", index=False)
        value_candidates(df).to_excel(writer, sheet_name="value_candidates", index=False)
        coverage(df).to_excel(writer, sheet_name="coverage", index=False)
        df.to_excel(writer, sheet_name="cleaned_data", index=False)

    write_report(df, summary, by_type, by_location, by_repair)
    print(f"Saved analysis workbook: {OUTPUT_XLSX}")
    print(f"Saved text report: {OUTPUT_TXT}")
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
