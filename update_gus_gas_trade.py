import os
import re
import time
import requests
import pandas as pd
from pathlib import Path
from datetime import datetime
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ============================================================
# SETTINGS
# ============================================================

BASE_URL = "https://api-dbw.stat.gov.pl/api"

OUT = Path("gus_dbw_trade")
OUT.mkdir(exist_ok=True)

OUTPUT_FILE = OUT / "gus_gas_TJ_2018_plus_IMPORT_EXPORT.xlsx"
SECTIONS_FILE = OUT / "03_section_positions.xlsx"

API_KEY = os.getenv("GUS_DBW_API_KEY")

HEADERS = {"accept": "application/json"}
if API_KEY:
    HEADERS["X-ClientId"] = API_KEY

ID_PRZEKROJ = 1434

START_YEAR_IF_NO_FILE = 2018

MONTHS = {
    247: 1,
    248: 2,
    249: 3,
    250: 4,
    251: 5,
    252: 6,
    253: 7,
    254: 8,
    255: 9,
    256: 10,
    257: 11,
    258: 12,
}

MONTH_TO_ID_OKRES = {v: k for k, v in MONTHS.items()}

TARGET_CODES = [
    "27111100",  # Gaz ziemny skroplony [TJ]
    "27112100",  # Gaz ziemny w stanie gazowym [TJ]
]

DATASETS = [
    {
        "sheet": "import_origin_gas_TJ",
        "id_zmienna": 221,
        "description": "Import towarów wg kraju pochodzenia",
    },
    {
        "sheet": "export_gas_TJ",
        "id_zmienna": 220,
        "description": "Eksport towarów",
    },
]

PAGE_SIZE = 5000
MAX_PAGE = 100
SLEEP = 0.7

ONLY_TOTAL_COUNTRY = True

KEY_COLS = [
    "Zmienna",
    "Typ informacji",
    "Kraje towary",
    "CN - uzupełniająca jednostka miary",
    "Jednostka terytorialna",
]

PERIOD_RE = re.compile(r"^\d{4} M\d{2}$")


# ============================================================
# SESSION
# ============================================================

session = requests.Session()

retries = Retry(
    total=5,
    backoff_factor=2,
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=["GET"],
)

adapter = HTTPAdapter(max_retries=retries)
session.mount("https://", adapter)
session.mount("http://", adapter)


# ============================================================
# BASIC HELPERS
# ============================================================

def is_period_col(col):
    return isinstance(col, str) and bool(PERIOD_RE.match(col))


def period_to_tuple(period):
    year, month = period.split(" M")
    return int(year), int(month)


def tuple_to_period(year, month):
    return f"{year} M{month:02d}"


def next_month(year, month):
    if month == 12:
        return year + 1, 1
    return year, month + 1


def month_range(start_year, start_month, end_year, end_month):
    y, m = start_year, start_month

    while (y, m) <= (end_year, end_month):
        yield y, m
        y, m = next_month(y, m)


def get_last_period_from_sheet(path, sheet_name):
    if not path.exists():
        return None

    try:
        df = pd.read_excel(path, sheet_name=sheet_name, nrows=1)
    except Exception:
        return None

    periods = [c for c in df.columns if is_period_col(c)]

    if not periods:
        return None

    periods_sorted = sorted(periods, key=period_to_tuple)
    return periods_sorted[-1]


def get_download_start(path, sheet_name):
    last_period = get_last_period_from_sheet(path, sheet_name)

    if last_period is None:
        return START_YEAR_IF_NO_FILE, 1, None

    y, m = period_to_tuple(last_period)
    ny, nm = next_month(y, m)

    return ny, nm, last_period


def safe_int(x):
    try:
        return int(x)
    except Exception:
        return None


# ============================================================
# API HELPERS
# ============================================================

def get_json(endpoint, params=None, allow_404=False):
    url = f"{BASE_URL}/{endpoint}"

    for attempt in range(1, 6):
        try:
            r = session.get(
                url,
                params=params,
                headers=HEADERS,
                timeout=90,
            )

            time.sleep(SLEEP)

            if r.status_code == 404 and allow_404:
                return None

            if r.status_code != 200:
                print("URL:", r.url)
                print("STATUS:", r.status_code)
                print("TEXT:", r.text[:1000])
                r.raise_for_status()

            return r.json()

        except requests.exceptions.RequestException as e:
            print(f"Request failed, attempt {attempt}/5:", e)
            time.sleep(5 * attempt)

    raise RuntimeError("API request failed after 5 attempts")


def extract_rows(obj):
    if obj is None:
        return None

    if isinstance(obj, list):
        return obj

    candidates = []

    def walk(x):
        if isinstance(x, list):
            if x and all(isinstance(i, dict) for i in x):
                candidates.append(x)
            for i in x:
                walk(i)
        elif isinstance(x, dict):
            for v in x.values():
                walk(v)

    walk(obj)

    return max(candidates, key=len) if candidates else []


def fetch_page(id_zmienna, year, id_okres, page):
    params = {
        "id-zmienna": id_zmienna,
        "id-przekroj": ID_PRZEKROJ,
        "id-rok": year,
        "id-okres": id_okres,
        "ile-na-stronie": PAGE_SIZE,
        "numer-strony": page,
        "lang": "pl",
    }

    obj = get_json(
        "variable/variable-data-section",
        params=params,
        allow_404=True,
    )

    return extract_rows(obj)


# ============================================================
# METADATA
# ============================================================

def load_sections():
    if not SECTIONS_FILE.exists():
        raise FileNotFoundError(
            f"Brakuje pliku: {SECTIONS_FILE}. "
            "Ten plik musi być w repozytorium, bo z niego bierzemy ID produktów i krajów."
        )

    return pd.read_excel(SECTIONS_FILE)


def get_target_product_ids(sections):
    products = sections[
        (sections["id-przekroj"] == ID_PRZEKROJ)
        & (sections["symbol"].astype(str).isin(TARGET_CODES))
    ].copy()

    if products.empty:
        raise RuntimeError(
            "Nie znaleziono kodów 27111100 i 27112100 dla przekroju 1434."
        )

    print("\nTarget products:")
    print(
        products[
            [
                "id-przekroj",
                "id-wymiar",
                "nazwa-wymiar",
                "id-pozycja",
                "symbol",
                "nazwa-pozycja",
            ]
        ].to_string(index=False)
    )

    return products["id-pozycja"].astype(int).tolist()


def get_total_country_ids(sections):
    countries = sections[
        (sections["id-przekroj"] == ID_PRZEKROJ)
        & (sections["id-wymiar"] == 734)
        & (sections["nazwa-pozycja"].astype(str).str.lower().eq("ogółem"))
    ].copy()

    ids = countries["id-pozycja"].astype(int).tolist()

    if not ids:
        raise RuntimeError("Nie znaleziono kraju 'Ogółem' dla przekroju 1434.")

    return ids


# ============================================================
# DOWNLOAD
# ============================================================

def filter_rows(rows, target_product_ids, total_country_ids):
    if not rows:
        return []

    out = []

    for row in rows:
        product_id = safe_int(row.get("id-pozycja-3"))
        country_id = safe_int(row.get("id-pozycja-2"))

        if product_id not in target_product_ids:
            continue

        if ONLY_TOTAL_COUNTRY and country_id not in total_country_ids:
            continue

        out.append(row)

    return out


def download_dataset(dataset, target_product_ids, total_country_ids, start_year, start_month, end_year, end_month):
    all_rows = []

    for year, month in month_range(start_year, start_month, end_year, end_month):
        id_okres = MONTH_TO_ID_OKRES[month]

        print("\nDownloading:")
        print(
            "dataset:", dataset["sheet"],
            "id_zmienna:", dataset["id_zmienna"],
            "year:", year,
            "month:", month,
            "id_okres:", id_okres,
        )

        for page in range(MAX_PAGE + 1):
            rows = fetch_page(
                id_zmienna=dataset["id_zmienna"],
                year=year,
                id_okres=id_okres,
                page=page,
            )

            if rows is None:
                print("page:", page, "does not exist - stopping this month")
                break

            matched = filter_rows(
                rows=rows,
                target_product_ids=target_product_ids,
                total_country_ids=total_country_ids,
            )

            print(
                "page:", page,
                "rows:", len(rows),
                "matched:", len(matched),
            )

            for row in matched:
                row["dataset"] = dataset["sheet"]
                row["dataset_description"] = dataset["description"]
                row["year"] = year
                row["month"] = month
                row["date"] = f"{year}-{month:02d}-01"
                row["page_downloaded"] = page

            all_rows.extend(matched)

            if len(rows) < PAGE_SIZE:
                print("last page reached")
                break

    if not all_rows:
        return pd.DataFrame()

    df = pd.json_normalize(all_rows)
    df = df.drop_duplicates()

    return df


# ============================================================
# LAYOUT
# ============================================================

def add_names(df, sections):
    if df.empty:
        return df

    territory = sections[
        (sections["id-przekroj"] == ID_PRZEKROJ)
        & (sections["id-wymiar"] == 2)
    ][["id-pozycja", "nazwa-pozycja", "symbol"]].drop_duplicates(subset=["id-pozycja"])

    territory = territory.rename(
        columns={
            "id-pozycja": "id-pozycja-1",
            "nazwa-pozycja": "territory_name",
            "symbol": "territory_symbol",
        }
    )

    countries = sections[
        (sections["id-przekroj"] == ID_PRZEKROJ)
        & (sections["id-wymiar"] == 734)
    ][["id-pozycja", "nazwa-pozycja", "symbol"]].drop_duplicates(subset=["id-pozycja"])

    countries = countries.rename(
        columns={
            "id-pozycja": "id-pozycja-2",
            "nazwa-pozycja": "country_name",
            "symbol": "country_symbol",
        }
    )

    products = sections[
        (sections["id-przekroj"] == ID_PRZEKROJ)
        & (sections["id-wymiar"] == 1181)
        & (sections["symbol"].astype(str).isin(TARGET_CODES))
    ][["id-pozycja", "nazwa-pozycja", "symbol"]].drop_duplicates(subset=["id-pozycja"])

    products = products.rename(
        columns={
            "id-pozycja": "id-pozycja-3",
            "nazwa-pozycja": "product_name",
            "symbol": "product_code",
        }
    )

    df = df.merge(territory, on="id-pozycja-1", how="left")
    df = df.merge(countries, on="id-pozycja-2", how="left")
    df = df.merge(products, on="id-pozycja-3", how="left")

    return df


def make_gus_layout(df, sections):
    if df.empty:
        return pd.DataFrame()

    df = add_names(df, sections)

    df["wartosc"] = (
        df["wartosc"]
        .astype(str)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    df["wartosc"] = pd.to_numeric(df["wartosc"], errors="coerce")

    df["period"] = (
        df["year"].astype(int).astype(str)
        + " M"
        + df["month"].astype(int).astype(str).str.zfill(2)
    )

    df["typ_informacji"] = "[-]"

    df["territory_name"] = df["territory_name"].fillna("POLSKA")
    df["country_name"] = df["country_name"].fillna("")
    df["product_name"] = df["product_name"].fillna(df["product_code"].astype(str))

    key_cols = [
        "dataset_description",
        "typ_informacji",
        "country_name",
        "product_name",
        "territory_name",
        "period",
    ]

    df = df.drop_duplicates(subset=key_cols + ["wartosc"])

    layout = df.pivot_table(
        index=[
            "dataset_description",
            "typ_informacji",
            "country_name",
            "product_name",
            "territory_name",
        ],
        columns="period",
        values="wartosc",
        aggfunc="first",
    ).reset_index()

    period_cols = sorted(
        [c for c in layout.columns if is_period_col(c)],
        key=period_to_tuple,
    )

    layout = layout[
        [
            "dataset_description",
            "typ_informacji",
            "country_name",
            "product_name",
            "territory_name",
        ]
        + period_cols
    ]

    layout = layout.rename(
        columns={
            "dataset_description": "Zmienna",
            "typ_informacji": "Typ informacji",
            "country_name": "Kraje towary",
            "product_name": "CN - uzupełniająca jednostka miary",
            "territory_name": "Jednostka terytorialna",
        }
    )

    return layout


def read_existing_sheet(path, sheet_name):
    if not path.exists():
        return pd.DataFrame()

    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()


def merge_existing_with_new(existing_df, new_df):
    if existing_df.empty:
        return new_df.copy()

    if new_df.empty:
        return existing_df.copy()

    for col in KEY_COLS:
        if col not in existing_df.columns:
            existing_df[col] = ""
        if col not in new_df.columns:
            new_df[col] = ""

    combined = pd.concat([existing_df, new_df], ignore_index=True, sort=False)

    period_cols = sorted(
        [c for c in combined.columns if is_period_col(c)],
        key=period_to_tuple,
    )

    def last_notna(series):
        non_empty = series.dropna()
        if non_empty.empty:
            return pd.NA
        return non_empty.iloc[-1]

    merged = (
        combined
        .groupby(KEY_COLS, dropna=False, as_index=False)
        .agg(last_notna)
    )

    merged = merged[KEY_COLS + period_cols]

    return merged


# ============================================================
# EXCEL FORMATTING
# ============================================================

def autofit_excel(path):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = load_workbook(path)

    header_fill = PatternFill("solid", fgColor="EDE7F6")
    header_font = Font(bold=True, color="5E35B1")
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="center", wrap_text=True)

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter

            for cell in col:
                value = cell.value
                if value is None:
                    continue
                max_len = max(max_len, len(str(value)))

            width = min(max(max_len + 2, 10), 45)
            ws.column_dimensions[col_letter].width = width

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0'

    wb.save(path)


# ============================================================
# MAIN
# ============================================================

def main():
    now = datetime.now()
    end_year = now.year
    end_month = now.month

    sections = load_sections()
    target_product_ids = get_target_product_ids(sections)
    total_country_ids = get_total_country_ids(sections)

    print("\nTotal country IDs:")
    print(total_country_ids)

    final_layouts = {}
    summary_rows = []
    any_new_data = False

    for dataset in DATASETS:
        sheet = dataset["sheet"]

        print("\n==============================")
        print("DATASET:", sheet)
        print("==============================")

        start_year, start_month, last_existing_period = get_download_start(OUTPUT_FILE, sheet)

        print("Last existing period:", last_existing_period)
        print("Download from:", tuple_to_period(start_year, start_month))
        print("Download to:", tuple_to_period(end_year, end_month))

        existing_df = read_existing_sheet(OUTPUT_FILE, sheet)

        if (start_year, start_month) > (end_year, end_month):
            print("No new months to download.")
            final_df = existing_df
            raw_df = pd.DataFrame()
            new_layout_df = pd.DataFrame()
        else:
            raw_df = download_dataset(
                dataset=dataset,
                target_product_ids=target_product_ids,
                total_country_ids=total_country_ids,
                start_year=start_year,
                start_month=start_month,
                end_year=end_year,
                end_month=end_month,
            )

            new_layout_df = make_gus_layout(raw_df, sections)

            if not new_layout_df.empty:
                any_new_data = True

            final_df = merge_existing_with_new(existing_df, new_layout_df)

        final_layouts[sheet] = final_df

        period_cols = [c for c in final_df.columns if is_period_col(c)] if not final_df.empty else []
        last_final_period = sorted(period_cols, key=period_to_tuple)[-1] if period_cols else None

        summary_rows.append(
            {
                "dataset": sheet,
                "description": dataset["description"],
                "last_existing_period_before_update": last_existing_period,
                "download_started_from": tuple_to_period(start_year, start_month),
                "download_attempted_to": tuple_to_period(end_year, end_month),
                "new_raw_rows": len(raw_df),
                "new_layout_rows": len(new_layout_df),
                "final_rows": len(final_df),
                "last_final_period_after_update": last_final_period,
                "only_total_country": ONLY_TOTAL_COUNTRY,
                "updated_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
            }
        )

    summary_df = pd.DataFrame(summary_rows)

    if OUTPUT_FILE.exists() and not any_new_data:
        print("\nNo new data found. Existing file was not changed.")
        print(summary_df)
        return

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="summary", index=False)

        for sheet_name, layout_df in final_layouts.items():
            if not layout_df.empty:
                layout_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    autofit_excel(OUTPUT_FILE)

    print("\nSaved:", OUTPUT_FILE)
    print(summary_df)


if __name__ == "__main__":
    main()
