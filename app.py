import os
import re
import time
from datetime import datetime

import pandas as pd
import streamlit as st
from PIL import Image

st.set_page_config(page_title="Sunshine Islands Bus App", layout="wide")


def get_route_name(code: str) -> str:
    code = str(code).strip()
    if not code:
        return code

    last_char = code[-1]
    if last_char == "M" and len(code) >= 2 and code[-2] in ["A", "E", "W", "N", "S"]:
        return code[:-2] + code[-1]

    if last_char.isdigit() and len(code) >= 2:
        if code[-2].isdigit():
            return code
        if code[-2] in ["A", "E", "W", "N", "S"]:
            return code[:-2]

    if last_char in ["A", "E", "W", "N", "S"] and len(code) >= 2:
        return code[:-1]

    return code


def is_route_eligible_for_bbi(current_route: str, rule_route_str: str) -> bool:
    curr = str(current_route).strip().upper()
    rules = [r.strip().upper() for r in str(rule_route_str).split(",")]
    for rule in rules:
        if not rule:
            continue
        if curr == rule or curr.startswith(rule):
            return True
    return False


def is_circular_route(route_code: str) -> bool:
    letters_only = "".join(c for c in str(route_code) if c.isalpha())
    return (letters_only[-1] if letters_only else "") == "A"


def is_route_deleted(condition_val) -> bool:
    if condition_val is None or pd.isna(condition_val):
        return False
    s = str(condition_val).strip().lower()
    return "deleted" in s or "已取消" in s


def get_destination_and_classification(route_code: str, route_df: pd.DataFrame, lang: str):
    route_df_sorted = route_df.sort_values("stop_order")
    is_circ = is_circular_route(route_code)

    if is_circ and not route_df_sorted.empty:
        first_row = route_df_sorted.iloc[0]
        if lang == "ZH":
            disp = first_row.get("Display_info_chi", None)
            dest = (
                disp
                if pd.notna(disp) and str(disp).strip() not in ["", "-"]
                else (
                    first_row.get("stop_chinese_name")
                    if pd.notna(first_row.get("stop_chinese_name"))
                    else first_row.get("stop_english_name")
                )
            )
        else:
            disp = first_row.get("Display_info_eng", None)
            dest = (
                disp
                if pd.notna(disp) and str(disp).strip() not in ["", "-"]
                else first_row.get("stop_english_name")
            )
    else:
        name_col = "stop_chinese_name" if lang == "ZH" else "stop_english_name"
        dest = (
            route_df_sorted.iloc[-1][name_col]
            if not route_df_sorted.empty and pd.notna(route_df_sorted.iloc[-1][name_col])
            else "-"
        )

    labels = []
    if not str(route_code)[-1].isalpha():
        labels.append("特別班次" if lang == "ZH" else "Special Departure")
    if str(route_code).startswith("N"):
        labels.append("通宵路線" if lang == "ZH" else "Overnight Route")

    letters_only = "".join(c for c in str(route_code) if c.isalpha())
    last_letter = letters_only[-1] if letters_only else ""
    dir_map = {
        "A": ("循環線", "Circular Route"),
        "E": ("東行", "Eastbound"),
        "W": ("西行", "Westbound"),
        "N": ("北行", "Northbound"),
        "S": ("南行", "Southbound"),
    }
    if last_letter in dir_map:
        labels.append(dir_map[last_letter][0] if lang == "ZH" else dir_map[last_letter][1])

    cond_val = (
        route_df_sorted.iloc[0].get("route_condition", None)
        if "route_condition" in route_df_sorted.columns
        else None
    )
    if cond_val and pd.notna(cond_val) and str(cond_val).strip() not in ["-", "nan", ""]:
        for item in re.split(r"[,/+\n]", str(cond_val).strip()):
            clean = item.strip()
            if clean.lower() not in ["normal", "night", "regular", "none"]:
                labels.append(clean)

    if not labels:
        labels.append("常規路線" if lang == "ZH" else "Regular Route")

    return str(dest), " | ".join(labels)


@st.cache_data
def load_workbook_data(excel_file):
    xls = pd.ExcelFile(excel_file)
    return {
        "stops": pd.read_excel(xls, "Route_stops") if "Route_stops" in xls.sheet_names else None,
        "timetable": pd.read_excel(xls, "Timetable") if "Timetable" in xls.sheet_names else None,
        "two_way": (
            pd.read_excel(xls, "two_way_selectional_fare")
            if "two_way_selectional_fare" in xls.sheet_names
            else None
        ),
        "z7": (
            pd.read_excel(xls, "Z7_interchange_discount")
            if "Z7_interchange_discount" in xls.sheet_names
            else None
        ),
        "northern": (
            pd.read_excel(xls, "Northern_Interchange_discount")
            if "Northern_Interchange_discount" in xls.sheet_names
            else None
        ),
    }


def parse_hhmm_to_minutes(hhmm_str: str) -> int:
    try:
        parts = str(hhmm_str).strip().split(":")
        if len(parts) >= 2:
            h, m = int(parts[0]), int(parts[1])
            if 0 <= h <= 29 and 0 <= m <= 59:
                return h * 60 + m
    except ValueError:
        pass
    now = datetime.now()
    return now.hour * 60 + now.minute


def format_minutes_to_hhmm(total_mins: float) -> str:
    m_int = int(round(total_mins)) % (24 * 60)
    return f"{m_int // 60:02d}:{m_int % 60:02d}"


def get_stop_offset_minutes(df_route: pd.DataFrame, row: pd.Series) -> float:
    for col in df_route.columns:
        if str(col).strip().lower().replace(" ", "_") == "travel_time_from_terminus":
            val = row.get(col)
            if pd.notna(val):
                try:
                    return float(val)
                except ValueError:
                    pass
    try:
        return float(max(0, int(row["stop_order"]) - 1) * 2)
    except Exception:
        return 0.0


def get_next_arrival(df_timetable, route_code, df_route, board_row, earliest_min):
    offset = get_stop_offset_minutes(df_route, board_row)
    if df_timetable is None or df_timetable.empty:
        return earliest_min

    tt_df = df_timetable[df_timetable["route_code"] == route_code]
    if tt_df.empty:
        return earliest_min

    slots = []
    for _, r in tt_df.iterrows():
        dep_s = str(r.get("departure_time", "")).strip()
        int_s = str(r.get("interval", "")).strip()
        if not dep_s or dep_s == "-":
            continue
        try:
            p = dep_s.split(":")
            slots.append((int(p[0]) * 60 + int(p[1]), int_s))
        except ValueError:
            continue

    if not slots:
        return earliest_min

    slots.sort(key=lambda x: x[0])
    deps = []
    for i, (start_m, int_s) in enumerate(slots):
        next_m = slots[i + 1][0] if i + 1 < len(slots) else 1439
        if int_s in ["-", "nan", "", "None"]:
            deps.append(start_m)
        else:
            try:
                step = int(int_s)
                if step > 0:
                    cur = start_m
                    while cur < next_m:
                        deps.append(cur)
                        cur += step
                else:
                    deps.append(start_m)
            except ValueError:
                deps.append(start_m)

    for dep_m in sorted(set(deps)):
        arr_m = dep_m + offset
        if arr_m >= earliest_min:
            return arr_m
    return None


def get_boarding_fare(row: pd.Series) -> float:
    val = row.get("boarding_fare")
    if pd.notna(val) and str(val).strip() not in ["", "-", "nan", "None"]:
        try:
            return float(val)
        except ValueError:
            pass
    return 0.0


def get_interchange_fee(df_z7, df_northern, r_from: str, r_to: str):
    for df_i in (df_z7, df_northern):
        if df_i is not None and not df_i.empty:
            for _, r in df_i.iterrows():
                if is_route_eligible_for_bbi(r_from, str(r.get("From_route", ""))) and is_route_eligible_for_bbi(
                    r_to, str(r.get("To_route", ""))
                ):
                    fee = r.get("Fee")
                    if pd.notna(fee):
                        try:
                            return float(fee)
                        except ValueError:
                            pass
    return None


def run_navi(data, origin, dest, dep_str, lang, show_deleted):
    df_stops = data["stops"]
    if df_stops is None or df_stops.empty or not origin or not dest or origin == dest:
        return []

    start_q = parse_hhmm_to_minutes(dep_str)
    name_col = "stop_chinese_name" if lang == "ZH" else "stop_english_name"
    df = df_stops.copy()
    df["_label"] = df[name_col].fillna(df.get("stop_english_name", "")).astype(str).str.strip()

    groups = {}
    for r_code, r_df in df.groupby("route_code", sort=False):
        cond_val = r_df.iloc[0].get("route_condition", None) if "route_condition" in r_df.columns else None
        if not show_deleted and is_route_deleted(cond_val):
            continue
        groups[str(r_code)] = r_df.sort_values("stop_order").reset_index(drop=True)

    results = []

    # Direct
    for r_code, r_df in groups.items():
        o_rows = r_df[r_df["_label"] == origin]
        d_rows = r_df[r_df["_label"] == dest]
        for _, o_r in o_rows.iterrows():
            for _, d_r in d_rows.iterrows():
                if int(o_r["stop_order"]) < int(d_r["stop_order"]):
                    b_arr = get_next_arrival(data["timetable"], r_code, r_df, o_r, start_q)
                    if b_arr is None:
                        continue
                    ride = max(0.0, get_stop_offset_minutes(r_df, d_r) - get_stop_offset_minutes(r_df, o_r))
                    d_arr = b_arr + ride
                    fare = get_boarding_fare(o_r)
                    results.append(
                        {
                            "type": "Direct",
                            "dep_time": format_minutes_to_hhmm(b_arr),
                            "arr_time": format_minutes_to_hhmm(d_arr),
                            "total_mins": d_arr - start_q,
                            "total_fare": fare,
                            "legs": [
                                {
                                    "route_code": r_code,
                                    "route_name": get_route_name(r_code),
                                    "board": origin,
                                    "off": dest,
                                    "dep": format_minutes_to_hhmm(b_arr),
                                    "arr": format_minutes_to_hhmm(d_arr),
                                    "fare": fare,
                                }
                            ],
                        }
                    )

    # 1-Transfer
    origin_map = {}
    for r_code, r_df in groups.items():
        for _, o_r in r_df[r_df["_label"] == origin].iterrows():
            for _, m_r in r_df[r_df["stop_order"] > o_r["stop_order"]].iterrows():
                m_stop = m_r["_label"]
                if m_stop and m_stop not in (origin, dest):
                    origin_map.setdefault(m_stop, []).append((r_code, r_df, o_r, m_r))

    for r2_code, r2_df in groups.items():
        for _, d_r in r2_df[r2_df["_label"] == dest].iterrows():
            for _, m2_r in r2_df[r2_df["stop_order"] < d_r["stop_order"]].iterrows():
                m_stop = m2_r["_label"]
                if m_stop in origin_map:
                    for r1_code, r1_df, o_r, m1_r in origin_map[m_stop]:
                        if r1_code == r2_code:
                            continue
                        l1_dep = get_next_arrival(data["timetable"], r1_code, r1_df, o_r, start_q)
                        if l1_dep is None:
                            continue
                        m1 = max(0.0, get_stop_offset_minutes(r1_df, m1_r) - get_stop_offset_minutes(r1_df, o_r))
                        l1_arr = l1_dep + m1

                        l2_dep = get_next_arrival(data["timetable"], r2_code, r2_df, m2_r, l1_arr + 3.0)
                        if l2_dep is None:
                            continue
                        m2 = max(0.0, get_stop_offset_minutes(r2_df, d_r) - get_stop_offset_minutes(r2_df, m2_r))
                        l2_arr = l2_dep + m2

                        f1 = get_boarding_fare(o_r)
                        f2 = get_boarding_fare(m2_r)
                        inter = get_interchange_fee(data["z7"], data["northern"], r1_code, r2_code)
                        leg2_fare = max(0.0, inter) if inter is not None else f2
                        results.append(
                            {
                                "type": "1 Transfer",
                                "dep_time": format_minutes_to_hhmm(l1_dep),
                                "arr_time": format_minutes_to_hhmm(l2_arr),
                                "total_mins": l2_arr - start_q,
                                "total_fare": f1 + leg2_fare,
                                "legs": [
                                    {
                                        "route_code": r1_code,
                                        "route_name": get_route_name(r1_code),
                                        "board": origin,
                                        "off": m_stop,
                                        "dep": format_minutes_to_hhmm(l1_dep),
                                        "arr": format_minutes_to_hhmm(l1_arr),
                                        "fare": f1,
                                    },
                                    {
                                        "route_code": r2_code,
                                        "route_name": get_route_name(r2_code),
                                        "board": m_stop,
                                        "off": dest,
                                        "dep": format_minutes_to_hhmm(l2_dep),
                                        "arr": format_minutes_to_hhmm(l2_arr),
                                        "fare": leg2_fare,
                                    },
                                ],
                            }
                        )

    return results


with st.sidebar:
    lang = st.radio("Language / 語言", ["EN", "ZH"], horizontal=True)
    uploaded_excel = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx", "xls"])
    default_excel = "Data Excel.xlsx"
    excel_source = uploaded_excel if uploaded_excel is not None else (default_excel if os.path.exists(default_excel) else None)

st.title("🚌 Sunshine Islands Bus App" if lang == "EN" else "🚌 陽光群島巴士應用程式")

if not excel_source:
    st.warning("Please place `Data Excel.xlsx` in the app folder or upload it in the sidebar.")
    st.stop()

data = load_workbook_data(excel_source)
df_stops = data["stops"]

tab_routes, tab_navi, tab_map = st.tabs(
    ["Routes", "Navi", "Map"] if lang == "EN" else ["路線", "導航", "地圖"]
)

with tab_routes:
    col_s, col_d = st.columns([3, 1])
    with col_s:
        query = st.text_input("Search route or stop" if lang == "EN" else "搜尋路線或車站", "").strip().lower()
    with col_d:
        show_del_routes = st.checkbox("Show Deleted Routes" if lang == "EN" else "顯示已取消路線", value=False)

    if df_stops is not None and not df_stops.empty:
        name_col = "stop_chinese_name" if lang == "ZH" else "stop_english_name"
        group_cols = ["route_code"] + (["route_condition"] if "route_condition" in df_stops.columns else [])

        for key, r_df in df_stops.groupby(group_cols, sort=False, dropna=False):
            r_code = key[0] if isinstance(key, tuple) else key
            cond_v = key[1] if isinstance(key, tuple) else None
            if not show_del_routes and is_route_deleted(cond_v):
                continue

            r_sorted = r_df.sort_values("stop_order")
            r_name = get_route_name(r_code)
            first_s = (
                r_sorted.iloc[0][name_col]
                if pd.notna(r_sorted.iloc[0][name_col])
                else r_sorted.iloc[0]["stop_english_name"]
            )
            last_s, cls_label = get_destination_and_classification(r_code, r_sorted, lang)

            if query:
                hit = (
                    query in str(r_code).lower()
                    or query in str(r_name).lower()
                    or query in str(cls_label).lower()
                    or query in str(first_s).lower()
                    or query in str(last_s).lower()
                    or r_sorted[name_col].astype(str).str.lower().str.contains(query, na=False).any()
                )
                if not hit:
                    continue

            with st.expander(f"**{r_name}** [{cls_label}] — {first_s} ➔ {last_s}"):
                st.dataframe(r_sorted[["stop_order", "stop_chinese_name", "stop_english_name", "boarding_fare"]].dropna(how="all", axis=1), use_container_width=True)

with tab_navi:
    if df_stops is not None and not df_stops.empty:
        name_col = "stop_chinese_name" if lang == "ZH" else "stop_english_name"
        stops_list = sorted(
            list({s.strip() for s in df_stops[name_col].dropna().astype(str) if s.strip() and s.strip() != "-"})
        )
        c1, c2 = st.columns(2)
        with c1:
            origin = st.selectbox("Origin Stop" if lang == "EN" else "起點站", stops_list)
        with c2:
            dest = st.selectbox("Destination Stop" if lang == "EN" else "終點站", stops_list, index=min(1, len(stops_list) - 1))

        c3, c4, c5 = st.columns([1, 1, 2])
        with c3:
            dep_str = st.text_input("Departure (HH:MM)" if lang == "EN" else "出發時間 (HH:MM)", time.strftime("%H:%M"))
        with c4:
            sort_mode = st.radio("Sort" if lang == "EN" else "排序", ["Fastest", "Cheapest"], horizontal=True)
        with c5:
            show_del_navi = st.checkbox("Show Deleted Routes" if lang == "EN" else "顯示已取消路線", value=show_del_routes, key="navi_del")

        if st.button("Find Routes" if lang == "EN" else "搜尋路線"):
            res = run_navi(data, origin, dest, dep_str, lang, show_del_navi)
            res = sorted(
                res,
                key=lambda x: (x["total_fare"], x["total_mins"]) if sort_mode == "Cheapest" else (x["total_mins"], x["total_fare"]),
            )
            if not res:
                st.info("No travelling ways found between these two stops for this time.")
            else:
                for i, item in enumerate(res[:15], 1):
                    st.markdown(
                        f"**#{i} {item['type']} ({item['dep_time']} ➔ {item['arr_time']})** — "
                        f"Total: **{int(round(item['total_mins']))}m** | Fare: **${item['total_fare']:.1f}**"
                    )
                    for leg in item["legs"]:
                        st.write(
                            f"• {leg['route_name']} ({leg['route_code']}) [{leg['dep']}–{leg['arr']}]: "
                            f"{leg['board']} ➔ {leg['off']} (${leg['fare']:.1f})"
                        )
                    st.divider()

with tab_map:
    map_file = "map.jpg"[cite: 3]
    if os.path.exists(map_file):
        width = st.slider("Map Width (px)", min_value=300, max_value=1400, value=750, step=50)
        st.image(Image.open(map_file), width=width)
    else:
        st.warning("`map.jpg` not found in the app directory.")