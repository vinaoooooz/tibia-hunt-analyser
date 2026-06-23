import streamlit as st
import pandas as pd
import sys
import os
import re
import json
from pathlib import Path
from datetime import datetime
from statistics import mean
from io import BytesIO
from PIL import Image as PilImage

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from hunt_analyzer_excel import (
    parse_hunt_session, parse_hunt_session_json, group_sessions,
    compute_group_metrics, compute_global_kills_h,
    download_creature_image, split_multi_session, _clean_group_monsters,
    load_json_sessions, _tibia_log_dir, MINOR_THRESHOLD
)

st.set_page_config(page_title="Tibia Hunt Analyzer", layout="wide")
st.title("Tibia Hunt Analyzer")

if "data_ready" not in st.session_state:
    st.session_state.data_ready = False


def process_sessions(raw_sessions):
    groups = group_sessions(raw_sessions)
    group_metrics = [compute_group_metrics(g) for g in groups]
    top5 = compute_global_kills_h(raw_sessions)
    return groups, group_metrics, top5


def load_example_json():
    script_dir = Path(__file__).parent
    json_path = script_dir / "test_session.json"
    if json_path.exists():
        data = json.loads(json_path.read_text(encoding="utf-8"))
        s = parse_hunt_session_json(data)
        if s and "hours" in s:
            s["source"] = json_path.stem
            return [s]
    return None


def load_example_txt():
    script_dir = Path(__file__).parent
    hunts_dir = script_dir / "hunts"
    if hunts_dir.is_dir():
        raw = []
        for f in sorted(hunts_dir.glob("*.txt")):
            raw.extend(split_multi_session(f.read_text(encoding="utf-8"), f.stem))
        sessions = []
        for text, source in raw:
            s = parse_hunt_session(text)
            if s and "hours" in s:
                s["source"] = source
                sessions.append(s)
        if sessions:
            return sessions
    return None


def render_creature_image(name, width=48):
    buf = download_creature_image(name)
    if buf:
        img = PilImage.open(buf)
        st.image(img, width=width)
        return True
    return False


def fmt_number(n):
    if n is None:
        return "-"
    return f"{n:,}"


def fmt_hours(h):
    return f"{int(h):02d}:{round((h % 1) * 60):02d}"


def get_deviation_pct(val, avg):
    if avg and avg > 0:
        return round((val - avg) / avg * 100, 1)
    return 0


# ----- SIDEBAR -----
with st.sidebar:
    st.header("Entrada de Dados")

    input_mode = st.radio("Modo:", ["Pasta do Tibia", "Exemplo", "Colar texto", "Upload TXT"])

    sessions = None

    if input_mode == "Pasta do Tibia":
        tibia_dir = _tibia_log_dir()
        if tibia_dir is None:
            st.warning("Pasta de log do Tibia nao encontrada.")
        else:
            st.caption(f"Carregando de: `{tibia_dir}`")
            if st.button("Carregar da pasta do Tibia", type="primary", use_container_width=True):
                sessions = load_json_sessions(tibia_dir)

    elif input_mode == "Exemplo":
        st.caption("Carrega os arquivos da pasta hunts/ e test_session.json")
        if st.button("Carregar exemplo", type="primary", use_container_width=True):
            sessions = load_example_txt()
            if not sessions:
                sessions = load_example_json()

    elif input_mode == "Colar texto":
        text = st.text_area("Cole as sessoes (separadas por 2 linhas em branco):", height=200)
        if st.button("Processar", type="primary", use_container_width=True):
            chunks = [c.strip() for c in text.split("\n\n") if c.strip()]
            sessions = []
            for i, chunk in enumerate(chunks, 1):
                s = parse_hunt_session(chunk)
                if s and "hours" in s:
                    s["source"] = f"Sessao {i}"
                    sessions.append(s)

    elif input_mode == "Upload TXT":
        uploaded = st.file_uploader("Selecione arquivos .txt", type="txt", accept_multiple_files=True)
        if uploaded and st.button("Processar", type="primary", use_container_width=True):
            sessions = []
            for f in uploaded:
                try:
                    text = f.read().decode("utf-8")
                    for chunk, source in split_multi_session(text, Path(f.name).stem):
                        s = parse_hunt_session(chunk)
                        if s and "hours" in s:
                            s["source"] = source
                            sessions.append(s)
                except Exception as e:
                    st.error(f"Erro em {f.name}: {e}")

    if sessions is not None:
        if len(sessions) == 0:
            st.warning("Nenhuma sessao valida encontrada.")
        else:
            groups, group_metrics, top5 = process_sessions(sessions)
            st.session_state.sessions = sessions
            st.session_state.groups = groups
            st.session_state.group_metrics = group_metrics
            st.session_state.top5 = top5
            st.session_state.data_ready = True
            st.success(f"{len(sessions)} sessoes carregadas, {len(groups)} grupos")
            st.rerun()


# ----- MAIN CONTENT -----
if not st.session_state.get("data_ready"):
    st.info("Carregue dados no menu lateral para comecar.")
    st.stop()

sessions = st.session_state.sessions
groups = st.session_state.groups
group_metrics = st.session_state.group_metrics
top5 = st.session_state.top5

sorted_groups = sorted(group_metrics, key=lambda g: g["avg_raw_xp_h"], reverse=True)

tab_names = ["Resumo"] + [gm["name"].title() for gm in sorted_groups]
tabs = st.tabs(tab_names)

# ===== TAB RESUMO =====
with tabs[0]:
    st.subheader("TOP EXP")
    exp_data = []
    for rank, gm in enumerate(sorted_groups, 1):
        exp_data.append({
            "Hunt": gm["name"],
            "Sessoes": gm["count"],
            "Horas": fmt_hours(gm["total_hours"]),
            "Media Raw XP/h": f"{gm['avg_raw_xp_h']:,}",
            "Media Profit/h": f"{gm['avg_profit_h']:,}",
            "Total Kills": gm["total_kills"],
            "Rank": f"#{rank}",
        })
    st.dataframe(pd.DataFrame(exp_data), use_container_width=True, hide_index=True)

    chart_data = pd.DataFrame({
        "Hunt": [gm["name"] for gm in sorted_groups],
        "Raw XP/h": [gm["avg_raw_xp_h"] for gm in sorted_groups],
        "Profit/h": [gm["avg_profit_h"] for gm in sorted_groups],
    })
    st.subheader("Comparativo XP/h por Hunt")
    st.bar_chart(chart_data, x="Hunt", y=["Raw XP/h", "Profit/h"])

    # TOP 5 Profit/h + Preferred List side by side
    col1, gap, col2 = st.columns([5, 0.5, 5])

    with col1:
        st.subheader("TOP 5 Profit/h")
        profit_sorted = sorted(group_metrics, key=lambda g: g["avg_profit_h"], reverse=True)[:5]
        avg_profit = mean(g["avg_profit_h"] for g in profit_sorted) if profit_sorted else 0
        profit_rows = []
        for rank, gm in enumerate(profit_sorted, 1):
            dev = get_deviation_pct(gm["avg_profit_h"], avg_profit)
            profit_rows.append({
                "#": rank,
                "Hunt": gm["name"],
                "Profit/h": gm["avg_profit_h"],
                "+/-% Media": f"{dev:+.1f}%",
                "Criatura": gm["monsters"][0][0] if gm["monsters"] else "",
            })
        pdf = pd.DataFrame(profit_rows)
        st.dataframe(pdf.drop(columns=["Criatura"]), use_container_width=True, hide_index=True,
                     column_config={"Profit/h": st.column_config.NumberColumn(format="%d")})

        for _, row in pdf.iterrows():
            ic = st.columns([1, 5])
            with ic[0]:
                render_creature_image(row["Criatura"], width=48)
            with ic[1]:
                st.caption(f"**{row['Hunt']}** — {row['Profit/h']:,} profit/h")

    with col2:
        st.subheader("Preferred List - Top 5 Kills/h")
        if top5:
            avg_kh = mean(item[1] for item in top5) if top5 else 0
            pl_rows = []
            for idx, (cname, kh) in enumerate(top5, 1):
                dev = get_deviation_pct(kh, avg_kh)
                pl_rows.append({
                    "#": idx,
                    "Criatura": cname,
                    "Kills/h": kh,
                    "+/-% Media": f"{dev:+.1f}%",
                })
            pdf2 = pd.DataFrame(pl_rows)
            st.dataframe(pdf2, use_container_width=True, hide_index=True,
                         column_config={"Kills/h": st.column_config.NumberColumn(format="%d")})

            for _, row in pdf2.iterrows():
                ic = st.columns([1, 5])
                with ic[0]:
                    render_creature_image(row["Criatura"], width=48)
                with ic[1]:
                    dev = row["+/-% Media"]
                    dev_str = f" ({dev})" if dev else ""
                    st.caption(f"**{row['Criatura']}** — {row['Kills/h']:,} kills/h{dev_str}")

# ===== TAB DETALHE POR GRUPO =====
for tab_idx, gm in enumerate(sorted_groups):
    with tabs[tab_idx + 1]:
        st.subheader(f"Hunt: {gm['name'].title()}")

        # Metrics
        mc = st.columns(6)
        mc[0].metric("Sessoes", gm["count"])
        mc[1].metric("Horas total", fmt_hours(gm["total_hours"]))
        mc[2].metric("Media Raw XP/h", f"{gm['avg_raw_xp_h']:,}")
        mc[3].metric("Media Profit/h", f"{gm['avg_profit_h']:,}")
        mc[4].metric("Total kills", gm["total_kills"])
        mc[5].metric("Total Raw XP", f"{gm['total_raw_xp']:,}")

        # Creatures table
        st.subheader("Criaturas Mortas")
        total_hours = gm["total_hours"]
        creature_rows = []
        for idx, (cname, ccount) in enumerate(gm["monsters"], 1):
            kh = round(ccount / total_hours) if total_hours > 0 else 0
            pct = round(ccount / gm["total_kills"] * 100, 1) if gm["total_kills"] > 0 else 0
            creature_rows.append({
                "#": idx,
                "Criatura": cname,
                "Kills": ccount,
                "Kills/h": kh,
                "% Total": f"{pct}%",
            })
        cdf = pd.DataFrame(creature_rows)
        st.dataframe(cdf, use_container_width=True, hide_index=True,
                     column_config={"Kills": st.column_config.NumberColumn(format="%d"),
                                    "Kills/h": st.column_config.NumberColumn(format="%d")})

        icon_cols = st.columns(6)
        for i, (cname, _) in enumerate(gm["monsters"][:6]):
            with icon_cols[i]:
                render_creature_image(cname, width=64)

        # Historical trend
        sessions_with_date = [(s.get("date", ""), s.get("calc_raw_xp_h", 0))
                              for s in gm["sessions"] if s.get("date")]
        sessions_with_date.sort(key=lambda x: x[0])

        if len(sessions_with_date) >= 2:
            st.subheader("Tendencia Historica")
            mid = len(sessions_with_date) // 2
            first_half = [x[1] for x in sessions_with_date[:mid]]
            second_half = [x[1] for x in sessions_with_date[mid:]]
            avg_first = mean(first_half) if first_half else 0
            avg_second = mean(second_half) if second_half else 0
            if avg_first > 0:
                ratio = avg_second / avg_first
                if ratio >= 1.05:
                    trend = "Melhorando"
                elif ratio <= 0.95:
                    trend = "Piorando"
                else:
                    trend = "Estavel"
            else:
                trend = "-"

            tc = st.columns(5)
            tc[0].metric("Tendencia", trend)
            tc[1].metric("Media 1a metade", f"{avg_first:,.0f}")
            tc[2].metric("Media 2a metade", f"{avg_second:,.0f}")
            best_xp = max(x[1] for x in sessions_with_date)
            worst_xp = min(x[1] for x in sessions_with_date)
            best_date = max(sessions_with_date, key=lambda x: x[1])[0]
            worst_date = min(sessions_with_date, key=lambda x: x[1])[0]
            tc[3].metric("Melhor sessao", f"{best_xp:,}", best_date[-5:])
            tc[4].metric("Pior sessao", f"{worst_xp:,}", worst_date[-5:])

        # Sessions detail
        st.subheader("Detalhes das Sessoes neste Grupo")
        group_avg = gm["avg_raw_xp_h"]
        sess_rows = []
        n_sessions = len(gm["sessions"])
        for idx, s in enumerate(gm["sessions"]):
            xp = s.get("calc_raw_xp_h", 0)
            profit = s.get("calc_profit_h", 0)
            balance = s.get("balance", 0)
            kills = s.get("total_kills", 0)
            damage = s.get("damage", 0)
            dmg_h = round(damage / s.get("hours", 1)) if s.get("hours", 0) > 0 else 0
            dev = get_deviation_pct(xp, group_avg)
            sess_rows.append({
                "Sessao": s.get("source", f"Sessao {idx+1}"),
                "Duracao": s.get("duration_str", ""),
                "Raw XP Gain": s.get("raw_xp_gain", 0),
                "Raw XP/h": xp,
                "Profit/h": profit,
                "Balance": balance,
                "Kills": kills,
                "Damage": damage,
                "Damage/h": dmg_h,
                "% Media": dev,
            })

        sdf = pd.DataFrame(sess_rows)

        def color_threshold(val):
            if isinstance(val, (int, float)) and group_avg > 0 and val != 0:
                deviation = (val - group_avg) / group_avg
                if deviation < -0.30:
                    return "background-color: #FFCCCC"
                elif deviation > 0.30:
                    return "background-color: #CCFFCC"
            return ""

        styled = sdf.style.map(color_threshold, subset=["Raw XP/h"])
        st.dataframe(styled, use_container_width=True, hide_index=True,
                     column_config={
                         "Raw XP Gain": st.column_config.NumberColumn(format="%d"),
                         "Raw XP/h": st.column_config.NumberColumn(format="%d"),
                         "Profit/h": st.column_config.NumberColumn(format="%d"),
                         "Balance": st.column_config.NumberColumn(format="%d"),
                         "Kills": st.column_config.NumberColumn(format="%d"),
                         "Damage": st.column_config.NumberColumn(format="%d"),
                         "Damage/h": st.column_config.NumberColumn(format="%d"),
                     })

        # Line chart
        if n_sessions > 0:
            st.subheader("XP/h | Profit/h | Dano/h por Sessao")
            chart_df = pd.DataFrame({
                "Sessao": [f"S{s+1}" for s in range(n_sessions)],
                "Raw XP/h": [s.get("calc_raw_xp_h", 0) for s in gm["sessions"]],
                "Profit/h": [s.get("calc_profit_h", 0) for s in gm["sessions"]],
                "Damage/h": [round(s.get("damage", 0) / max(s.get("hours", 1), 0.01))
                             for s in gm["sessions"]],
            })
            st.line_chart(chart_df, x="Sessao", y=["Raw XP/h", "Profit/h", "Damage/h"],
                          use_container_width=True)

        # Kills per session per creature
        all_cnames = [cname for cname, _ in gm["monsters"]]
        if all_cnames:
            st.subheader("Kills por Sessao / Criatura")
            km_rows = []
            for s in gm["sessions"]:
                row_data = {"Sessao": s.get("source", "")}
                row_total = 0
                for cname in all_cnames:
                    k = s.get("monsters", {}).get(cname, 0)
                    row_data[cname.capitalize()] = k
                    row_total += k
                row_data["Total"] = row_total
                km_rows.append(row_data)
            kmdf = pd.DataFrame(km_rows)
            st.dataframe(kmdf, use_container_width=True, hide_index=True)

        # Items
        all_items = gm.get("items", [])
        if all_items:
            st.subheader("Itens Looteados")
            item_rows = []
            for idx, (iname, icount) in enumerate(all_items, 1):
                ih = round(icount / total_hours) if total_hours > 0 else 0
                item_rows.append({"#": idx, "Item": iname, "Qty": icount, "Qty/h": ih})
            st.dataframe(pd.DataFrame(item_rows), use_container_width=True, hide_index=True,
                         column_config={"Qty": st.column_config.NumberColumn(format="%d"),
                                        "Qty/h": st.column_config.NumberColumn(format="%d")})
