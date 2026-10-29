# -*- coding: utf-8 -*-
"""事件地區分析系統 (Streamlit 版)

由 Google Apps Script + HTML 版本改寫。
功能：Excel 匯入 → GPS（Excel 座標優先，失敗再用地址轉換）→ 篩選 → 分群 →
      事件地圖 / 熱力圖 / 熱點排行 / 點擊統計 → 匯出 Excel
"""
import html
import io
import re
import time

import folium
import numpy as np
import pandas as pd
import requests
import streamlit as st
from folium.plugins import Fullscreen, HeatMap
from scipy.sparse import coo_matrix
from scipy.sparse.csgraph import connected_components
from scipy.spatial import cKDTree
from streamlit_folium import st_folium

st.set_page_config(page_title="事件地區分析系統", layout="wide")

DEFAULT_CENTER = [23.7, 120.9]
ROC_RE = re.compile(r"^(\d{2,3})年(\d{1,2})月(\d{1,2})日(?:\s*(\d{1,2})時(\d{1,2})分)?$")
TIME_RE = re.compile(r"(?:^|\s|T)(\d{1,2}):(\d{2})")


# ============================================================
# 地址轉 GPS
# ============================================================
def _get_google_key():
    try:
        return st.secrets.get("GOOGLE_MAPS_API_KEY", "")
    except Exception:
        return ""


@st.cache_data(show_spinner=False, ttl=60 * 60 * 24 * 30)
def geocode_address(address: str, google_key: str):
    """回傳 (ok, lat, lng, formatted, status, error)。有 Google key 用 Google，否則用 Nominatim(OSM)。"""
    address = str(address or "").strip()
    if not address:
        return (False, None, None, "", "EMPTY_ADDRESS", "地址為空白")
    try:
        if google_key:
            r = requests.get(
                "https://maps.googleapis.com/maps/api/geocode/json",
                params={"address": address, "key": google_key,
                        "language": "zh-TW", "region": "tw"},
                timeout=15,
            ).json()
            status = r.get("status", "UNKNOWN_STATUS")
            if status == "OK" and r.get("results"):
                top = r["results"][0]
                loc = top["geometry"]["location"]
                return (True, float(loc["lat"]), float(loc["lng"]),
                        top.get("formatted_address", ""), status, "")
            return (False, None, None, "", status, f"Google Geocoder 狀態：{status}")
        # Nominatim：限制每秒 1 次
        time.sleep(1.1)
        r = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": address, "format": "json", "limit": 1,
                    "countrycodes": "tw", "accept-language": "zh-TW"},
            headers={"User-Agent": "event-area-analysis/1.0"},
            timeout=15,
        ).json()
        if r:
            return (True, float(r[0]["lat"]), float(r[0]["lon"]),
                    r[0].get("display_name", ""), "OK", "")
        return (False, None, None, "", "ZERO_RESULTS", "Nominatim 查無結果")
    except Exception as e:  # noqa: BLE001
        return (False, None, None, "", "EXCEPTION", str(e))


# ============================================================
# 日期 / 時間處理
# ============================================================
def parse_datetime(value):
    """回傳 (顯示文字, 當天分鐘數或 None)。"""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "", None
    if isinstance(value, (pd.Timestamp,)) or hasattr(value, "hour"):
        ts = pd.Timestamp(value)
        if pd.isna(ts):
            return "", None
        return ts.strftime("%Y-%m-%d %H:%M"), ts.hour * 60 + ts.minute

    s = str(value).strip()
    if not s:
        return "", None

    m = ROC_RE.match(s)
    if m:
        year, month, day = int(m[1]) + 1911, int(m[2]), int(m[3])
        has_time = m[4] is not None
        hour, minute = int(m[4] or 0), int(m[5] or 0)
        try:
            ts = pd.Timestamp(year=year, month=month, day=day, hour=hour, minute=minute)
        except ValueError:
            return "", None
        if has_time:
            return ts.strftime("%Y-%m-%d %H:%M"), hour * 60 + minute
        return ts.strftime("%Y-%m-%d"), None

    tm = TIME_RE.search(s)
    if tm:
        h, mi = int(tm[1]), int(tm[2])
        if 0 <= h <= 23 and 0 <= mi <= 59:
            ts = pd.to_datetime(s, errors="coerce")
            text = ts.strftime("%Y-%m-%d %H:%M") if not pd.isna(ts) else s
            return text, h * 60 + mi
    ts = pd.to_datetime(s, errors="coerce")
    if not pd.isna(ts):
        return ts.strftime("%Y-%m-%d"), None
    return s, None


# ============================================================
# 處理 Excel 資料
# A 編號 / B 日期 / C 管轄單位 / D 地點 / I 受傷人數 / K 經度 / L 緯度
# ============================================================
def read_table(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded, header=None, dtype=object, encoding="utf-8-sig")
        except UnicodeDecodeError:
            uploaded.seek(0)
            df = pd.read_csv(uploaded, header=None, dtype=object, encoding="big5")
    else:
        df = pd.read_excel(uploaded, header=None, dtype=object)
    for c in range(df.shape[1], 12):  # 補足欄位到 L
        df[c] = None
    return df


def valid_gps(lat, lng):
    return (lat is not None and lng is not None
            and np.isfinite(lat) and np.isfinite(lng)
            and lat != 0 and lng != 0
            and -90 <= lat <= 90 and -180 <= lng <= 180)


def to_float(v):
    try:
        f = float(v)
        return f if np.isfinite(f) else None
    except (TypeError, ValueError):
        return None


def process_excel(raw: pd.DataFrame, google_key: str, progress=None) -> pd.DataFrame:
    rows = []
    body = raw.iloc[1:]
    total = len(body)
    for n, (_, r) in enumerate(body.iterrows(), start=1):
        rid = r[0]
        if rid is None or (isinstance(rid, float) and np.isnan(rid)) or str(rid).strip() == "":
            continue
        date_text, minutes = parse_datetime(r[1])
        unit = "" if pd.isna(r[2]) else str(r[2]).strip()
        address = "" if pd.isna(r[3]) else str(r[3]).strip()
        injured = to_float(r[8]) or 0
        lng, lat = to_float(r[10]), to_float(r[11])

        source = status = error = formatted = ""
        if valid_gps(lat, lng):
            source, status = "Excel K/L", "成功"
        else:
            ok, glat, glng, formatted, gstatus, error = geocode_address(address, google_key)
            source = "地址轉換"
            if ok:
                lat, lng, status = glat, glng, "成功"
            else:
                lat = lng = None
                status = "失敗"
        rows.append({
            "id": str(rid).strip(), "date": date_text, "minutes": minutes,
            "unit": unit, "address": address, "injured": int(injured),
            "lng": lng, "lat": lat,
            "gps_source": source, "gps_status": status,
            "gps_error": error, "formatted_address": formatted,
        })
        if progress and (n % 5 == 0 or n == total):
            progress.progress(n / total, text=f"處理中 {n}/{total}")
    return pd.DataFrame(rows)


# ============================================================
# 分群（距離 <= radius 的事件連成同一群，等同原本的 BFS）
# ============================================================
@st.cache_data(show_spinner=False)
def compute_labels(lat: tuple, lng: tuple, radius: float):
    lat = np.asarray(lat)
    lng = np.asarray(lng)
    n = len(lat)
    if n == 0:
        return np.array([], dtype=int)
    lat0 = np.deg2rad(lat.mean())
    xy = np.column_stack([lng * 111320 * np.cos(lat0), lat * 110574])
    pairs = cKDTree(xy).query_pairs(radius, output_type="ndarray")
    if len(pairs) == 0:
        return np.arange(1, n + 1)
    g = coo_matrix((np.ones(len(pairs)), (pairs[:, 0], pairs[:, 1])), shape=(n, n))
    _, labels = connected_components(g, directed=False)
    return labels + 1


def haversine(lat1, lng1, lat2, lng2):
    r = 6371000.0
    p1, p2 = np.radians(lat1), np.radians(lat2)
    dphi = p2 - p1
    dlmb = np.radians(np.asarray(lng2) - lng1)
    a = np.sin(dphi / 2) ** 2 + np.cos(p1) * np.cos(p2) * np.sin(dlmb / 2) ** 2
    return 2 * r * np.arcsin(np.sqrt(a))


def build_groups(df: pd.DataFrame, radius: float):
    """df 需為已含有效 GPS 的資料；回傳 (df 加群組欄位, groups DataFrame)。"""
    df = df.copy()
    if df.empty:
        df["group_id"] = pd.Series(dtype=int)
        df["group_count"] = pd.Series(dtype=int)
        df["group_injured"] = pd.Series(dtype=int)
        return df, pd.DataFrame(columns=["id", "count", "injured", "lat", "lng"])
    df["group_id"] = compute_labels(tuple(df["lat"]), tuple(df["lng"]), float(radius))
    g = df.groupby("group_id")
    df["group_count"] = g["id"].transform("count")
    df["group_injured"] = g["injured"].transform("sum")
    groups = g.agg(count=("id", "count"), injured=("injured", "sum"),
                   lat=("lat", "mean"), lng=("lng", "mean")).reset_index()
    groups = groups.rename(columns={"group_id": "id"})
    return df, groups


def group_color(count):
    return "#dc2626" if count >= 10 else "#eab308" if count >= 4 else "#16a34a"


def group_label(count):
    return "高密度" if count >= 10 else "中密度" if count >= 4 else "一般"


def rank_groups(groups: pd.DataFrame, top_n: int) -> pd.DataFrame:
    """事件數相同並列相同名次（與原版相同）。"""
    s = groups.sort_values("count", ascending=False).reset_index(drop=True)
    ranks, last, cur = [], None, 0
    for i, c in enumerate(s["count"]):
        if c != last:
            cur, last = i + 1, c
        ranks.append(cur)
    s["rank"] = ranks
    return s[s["rank"] <= top_n]


# ============================================================
# 地圖
# ============================================================
def new_map():
    m = folium.Map(location=DEFAULT_CENTER, zoom_start=7, tiles="OpenStreetMap",
                   control_scale=True)
    Fullscreen(title="全螢幕", title_cancel="離開全螢幕").add_to(m)
    return m


def fit(m, df):
    if df.empty:
        return
    if len(df) == 1:
        m.location = [df.iloc[0]["lat"], df.iloc[0]["lng"]]
        m.options["zoom"] = 15
        return
    m.fit_bounds([[df["lat"].min(), df["lng"].min()],
                  [df["lat"].max(), df["lng"].max()]], padding=(30, 30), max_zoom=16)


def count_icon(text, color, size=34, font=13):
    return folium.DivIcon(
        icon_size=(size, size), icon_anchor=(size // 2, size // 2),
        html=(f'<div style="background:{color};color:#fff;border:2px solid #fff;'
              f'border-radius:50%;width:{size}px;height:{size}px;display:flex;'
              f'align-items:center;justify-content:center;font-size:{font}px;'
              f'font-weight:bold;box-shadow:0 1px 5px rgba(0,0,0,.4)">{text}</div>'),
    )


def group_popup(g):
    return (f"<b>群組 {int(g['id'])}</b><br>事件數：{int(g['count'])}"
            f"<br>受傷人數：{int(g['injured'])}"
            f"<br>群組中心：{g['lat']:.6f}, {g['lng']:.6f}")


def build_event_map(pts, groups, radius):
    m = new_map()
    for _, r in pts.iterrows():
        pop = (f"<b>編號：</b>{html.escape(r['id'])}<br><b>日期：</b>{html.escape(r['date'])}"
               f"<br><b>管轄單位：</b>{html.escape(r['unit'])}<br><b>地點：</b>{html.escape(r['address'])}"
               f"<br><b>受傷人數：</b>{r['injured']}<br><b>群組：</b>{int(r['group_id'])}")
        folium.CircleMarker([r["lat"], r["lng"]], radius=5, color="#ef4444", weight=1,
                            fill=True, fill_color="#ef4444", fill_opacity=0.75,
                            popup=folium.Popup(pop, max_width=300)).add_to(m)
    for _, g in groups[groups["count"] > 2].iterrows():
        color = group_color(g["count"])
        folium.Circle([g["lat"], g["lng"]], radius=max(radius, 30), color=color,
                      weight=2, fill=True, fill_color=color, fill_opacity=0.16,
                      popup=folium.Popup(group_popup(g), max_width=260)).add_to(m)
        folium.Marker([g["lat"], g["lng"]], icon=count_icon(int(g["count"]), color),
                      popup=folium.Popup(group_popup(g), max_width=260),
                      z_index_offset=1000).add_to(m)
    fit(m, pts)
    return m


HEAT_GRADIENT = {0.0: "#172554", 0.125: "#1d4ed8", 0.25: "#22c55e",
                 0.5: "#facc15", 0.75: "#f97316", 1.0: "#991b1b"}


def best_point(members: pd.DataFrame, radius):
    """群組內「半徑/2 範圍內鄰居最多」的事件位置，作為熱點標記位置。"""
    lats, lngs = members["lat"].to_numpy(), members["lng"].to_numpy()
    best, best_i = -1, 0
    for i in range(len(members)):
        c = int((haversine(lats[i], lngs[i], lats, lngs) <= radius / 2).sum())
        if c > best:
            best, best_i = c, i
    return lats[best_i], lngs[best_i]


def build_heat_map(pts, groups, ranked, radius):
    m = new_map()
    if not pts.empty:
        max_count = max(int(groups["count"].max()), 1)
        HeatMap(pts[["lat", "lng", "group_count"]].astype(float).values.tolist(),
                radius=20, blur=12, min_opacity=0.35, max_zoom=17,
                max_val=float(max_count), gradient=HEAT_GRADIENT).add_to(m)
        for _, g in ranked.iterrows():
            members = pts[pts["group_id"] == g["id"]]
            lat, lng = best_point(members, radius)
            addr = members.iloc[0]["address"] or "—"
            pop = (f"<b>Top {int(g['rank'])} 熱點 (群組 {int(g['id'])})</b>"
                   f"<br>事件數：{int(g['count'])}<br>受傷人數：{int(g['injured'])}"
                   f"<br>代表地點：{html.escape(addr)}")
            folium.Marker([lat, lng], icon=count_icon(f"#{int(g['rank'])}", "#dc2626", 28, 12),
                          popup=folium.Popup(pop, max_width=280),
                          z_index_offset=2000).add_to(m)
    fit(m, pts)
    return m


# ============================================================
# 熱點排行表（含 24 小時密集度）
# ============================================================
def time_bar_html(members: pd.DataFrame) -> str:
    counts = [0] * 24
    for mnt in members["minutes"].dropna():
        counts[int(mnt) // 60] += 1
    mx = max(max(counts), 1)
    cells = ""
    for h, c in enumerate(counts):
        bg, fg = "#f1f5f9", "#64748b"
        if c > 0:
            ratio = c / mx
            gb = int(255 - ratio * 220)
            bg = f"rgb(255,{gb},{gb})"
            fg = "#fff" if ratio > 0.5 else "#991b1b"
        cells += (f'<div title="{h:02d}時: {c} 起事件" style="flex:1;height:20px;background:{bg};'
                  f'color:{fg};border-radius:2px;font-size:9px;display:flex;align-items:center;'
                  f'justify-content:center">{c if c else ""}</div>')
    axis = ('<div style="display:flex;justify-content:space-between;font-size:10px;'
            'color:#94a3b8"><span>00</span><span>06</span><span>12</span><span>18</span>'
            '<span>23</span></div>')
    return f'<div style="display:flex;gap:2px;min-width:240px">{cells}</div>{axis}'


def hotspot_table_html(pts, ranked) -> str:
    head = ("<tr><th>名次</th><th>群組</th><th>事件數</th><th>受傷人數</th>"
            "<th>24小時時間密集度 (0~23時)</th><th>中心座標</th><th>代表地點</th></tr>")
    body = ""
    for _, g in ranked.iterrows():
        members = pts[pts["group_id"] == g["id"]]
        addr = html.escape(members.iloc[0]["address"] or "—")
        body += (f"<tr><td><b>#{int(g['rank'])}</b></td><td>群組 {int(g['id'])}</td>"
                 f"<td><b>{int(g['count'])}</b></td><td>{int(g['injured'])}</td>"
                 f"<td>{time_bar_html(members)}</td>"
                 f"<td>{g['lat']:.6f}, {g['lng']:.6f}</td><td>{addr}</td></tr>")
    style = ("<style>.ht{border-collapse:collapse;width:100%;font-size:13px}"
             ".ht th{background:#f1f5f9;color:#334155;text-align:left;padding:8px}"
             ".ht td{border-bottom:1px solid #e2e8f0;padding:8px;vertical-align:top}</style>")
    return f'{style}<div style="overflow-x:auto"><table class="ht">{head}{body}</table></div>'


# ============================================================
# 匯出
# ============================================================
def export_excel(pts_all: pd.DataFrame, groups: pd.DataFrame) -> bytes:
    detail = pd.DataFrame({
        "編號": pts_all["id"], "日期": pts_all["date"], "管轄單位": pts_all["unit"],
        "地點": pts_all["address"], "受傷人數": pts_all["injured"],
        "經度": pts_all["lng"], "緯度": pts_all["lat"],
        "GPS來源": pts_all["gps_source"], "GPS狀態": pts_all["gps_status"],
        "GPS錯誤": pts_all["gps_error"],
        "群組編號": pts_all["group_id"], "群組事件數": pts_all["group_count"],
        "群組受傷人數": pts_all["group_injured"],
    })
    grp = pd.DataFrame({
        "群組編號": groups["id"], "事件數": groups["count"], "受傷人數": groups["injured"],
        "中心經度": groups["lng"], "中心緯度": groups["lat"],
        "密度": groups["count"].map(group_label),
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        detail.to_excel(w, sheet_name="事件明細", index=False)
        grp.to_excel(w, sheet_name="群組分析", index=False)
    return buf.getvalue()


# ============================================================
# UI
# ============================================================
st.title("事件地區分析系統")

# ---- 側邊欄：匯入 + 篩選 ----
with st.sidebar:
    st.header("一、匯入事件資料")
    uploaded = st.file_uploader("選擇 Excel 檔案", type=["xlsx", "xls", "csv"])
    st.caption("欄位：A 編號、B 日期、C 管轄單位、D 地點、I 受傷人數、K 經度、L 緯度（第一列為標題）")
    google_key = _get_google_key()
    st.caption("地址轉換：" + ("Google Geocoding API" if google_key
                          else "OpenStreetMap Nominatim（較慢，每秒 1 筆）"))
    if st.button("匯入並分析", type="primary", use_container_width=True):
        if uploaded is None:
            st.warning("請先選擇檔案。")
        else:
            try:
                raw = read_table(uploaded)
                if len(raw) <= 1:
                    st.error("Excel 沒有資料，或只有標題列。")
                else:
                    bar = st.progress(0.0, text="處理中…")
                    st.session_state["data"] = process_excel(raw, google_key, bar)
                    bar.empty()
            except Exception as e:  # noqa: BLE001
                st.error(f"讀取失敗：{e}")

    st.header("二、篩選條件")
    c1, c2 = st.columns(2)
    start_t = c1.time_input("開始時間", value=pd.Timestamp("00:00").time(), key="start_t")
    end_t = c2.time_input("結束時間", value=pd.Timestamp("23:59").time(), key="end_t")
    data_all = st.session_state.get("data")
    units = sorted(u for u in (data_all["unit"].unique() if data_all is not None else []) if u)
    unit_sel = st.selectbox("管轄單位", ["全部單位"] + units)
    radius = st.slider("事件分群距離（公尺）", 15, 500, 100, 5)
    click_radius = st.slider("熱力圖點擊統計半徑（公尺）", 10, 300, 50, 10)
    st.caption("時間篩選依日期欄位中的時間；沒有時間的資料會保留。開始晚於結束視為跨午夜。")

if data_all is None or data_all.empty:
    st.info("請先在左側選擇 Excel 檔案並按「匯入並分析」。")
    st.stop()

# ---- 篩選 ----
df = data_all.copy()
if unit_sel != "全部單位":
    df = df[df["unit"] == unit_sel]
s_min = start_t.hour * 60 + start_t.minute
e_min = end_t.hour * 60 + end_t.minute
if not (s_min == 0 and e_min == 1439):
    m_ = df["minutes"]
    if s_min > e_min:
        keep = (m_ >= s_min) | (m_ <= e_min)
    else:
        keep = (m_ >= s_min) & (m_ <= e_min)
    df = df[keep | m_.isna()]

# ---- 分群 ----
has_gps = df["lat"].notna() & df["lng"].notna()
pts, groups = build_groups(df[has_gps].reset_index(drop=True), radius)
ok_ids = set(pts["id"])

# ---- 三、事件統計 ----
st.subheader("三、事件統計")
cols = st.columns(6)
cols[0].metric("篩選後事件數", len(df))
cols[1].metric("受傷總人數", int(df["injured"].sum()))
cols[2].metric("事件群組數", len(groups))
cols[3].metric("GPS 成功", int(((df["gps_source"] == "Excel K/L") & (df["gps_status"] == "成功")).sum()))
cols[4].metric("地址轉換成功", int(((df["gps_source"] == "地址轉換") & (df["gps_status"] == "成功")).sum()))
cols[5].metric("GPS 失敗", int((df["gps_status"] != "成功").sum()))

# 把群組資訊合併回完整明細
detail = df.merge(pts[["id", "group_id", "group_count", "group_injured"]].drop_duplicates("id"),
                  on="id", how="left")

# ---- 四、事件分布地圖 ----
st.subheader("四、事件分布地圖")
st.caption("綠 / 黃 / 紅圈代表事件群組（>2 筆才顯示；≥4 黃、≥10 紅），圈中數字為事件數。"
           "地圖右上角可全螢幕。")
if pts.empty:
    st.warning("沒有具有效 GPS 的事件。")
else:
    st_folium(build_event_map(pts, groups, radius), height=700, use_container_width=True,
              returned_objects=[], key="event_map")

# ---- 五、熱力圖 ----
st.subheader("五、事件熱力圖")
if not pts.empty:
    top_n = st.selectbox("熱點排行顯示名次", [5, 10, 20, 50], index=1)
    ranked = rank_groups(groups, top_n)
    st.caption(f"目前顯示 {len(pts)} 筆有效 GPS 事件；最高群組事件數（最紅基準）："
               f"{int(groups['count'].max())}。點擊地圖可統計周圍 {click_radius} 公尺內事件。")
    out = st_folium(build_heat_map(pts, groups, ranked, radius), height=700,
                    use_container_width=True, returned_objects=["last_clicked"], key="heat_map")

    clicked = (out or {}).get("last_clicked")
    if clicked:
        clat, clng = clicked["lat"], clicked["lng"]
        d = haversine(clat, clng, pts["lat"].to_numpy(), pts["lng"].to_numpy())
        near = pts[d <= click_radius].copy()
        near["距點擊位置(m)"] = d[d <= click_radius].round().astype(int)
        near = near.sort_values("距點擊位置(m)")
        st.markdown("##### 熱力圖位置統計")
        a, b, c = st.columns(3)
        a.metric("附近事件數", len(near))
        b.metric("附近受傷人數", int(near["injured"].sum()))
        c.metric("統計半徑", f"{click_radius} m")
        st.caption(f"點擊座標：{clat:.6f}, {clng:.6f}")
        if near.empty:
            st.info(f"此位置 {click_radius} 公尺內沒有事件。")
        else:
            st.dataframe(near[["id", "date", "unit", "address", "injured", "距點擊位置(m)"]]
                         .rename(columns={"id": "編號", "date": "日期", "unit": "管轄單位",
                                          "address": "地點", "injured": "受傷人數"}),
                         use_container_width=True, hide_index=True)

    st.markdown("##### 事件密集熱點排行")
    st.caption("事件數相同的群組並列相同名次，並同步標記於熱力圖中。")
    st.markdown(hotspot_table_html(pts, ranked), unsafe_allow_html=True)

# ---- 六、事件明細 ----
st.subheader("六、事件明細")
show = detail.rename(columns={
    "id": "編號", "date": "日期", "unit": "管轄單位", "address": "地點",
    "injured": "受傷人數", "lng": "經度", "lat": "緯度", "gps_source": "GPS 來源",
    "gps_status": "GPS 狀態", "group_id": "群組"})
st.dataframe(show[["編號", "日期", "管轄單位", "地點", "受傷人數", "經度", "緯度",
                   "GPS 來源", "GPS 狀態", "群組"]],
             use_container_width=True, hide_index=True)

if not pts.empty:
    export_df = detail[detail["id"].isin(ok_ids) | True]
    st.download_button("匯出分析結果 (Excel)", data=export_excel(export_df, groups),
                       file_name="事件地區分析結果.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
