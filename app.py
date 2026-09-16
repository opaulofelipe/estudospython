from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
import unicodedata
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from filelock import FileLock

# -----------------------------------------------------------------------------
# Configuração
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Painel de Estudos",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).resolve().parent
PLANILHA = Path(os.getenv("STUDY_SHEET_PATH", BASE_DIR / "Estudos.xlsx"))
DATA_DIR = Path(os.getenv("STUDY_DATA_DIR", BASE_DIR / "data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)
PROGRESS_FILE = DATA_DIR / "progresso.json"
LOCK_FILE = DATA_DIR / "progresso.lock"
TIMEZONE = ZoneInfo("America/Sao_Paulo")

PALETTE = {
    "muted": "#8C93A8",
    "purple": "#62466B",
    "plum": "#45364B",
    "ink": "#2D2327",
    "text": "#F7F3F7",
    "soft": "#C8C4CF",
    "success": "#B8D8C0",
    "warning": "#E1C699",
}

DIAS_ORDEM = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
DIA_MAP = {
    0: "Segunda",
    1: "Terça",
    2: "Quarta",
    3: "Quinta",
    4: "Sexta",
    5: "Sábado",
    6: "Domingo",
}

# -----------------------------------------------------------------------------
# Estilos
# -----------------------------------------------------------------------------
st.markdown(
    f"""
    <style>
    :root {{
        --muted: {PALETTE['muted']};
        --purple: {PALETTE['purple']};
        --plum: {PALETTE['plum']};
        --ink: {PALETTE['ink']};
        --text: {PALETTE['text']};
        --soft: {PALETTE['soft']};
        --success: {PALETTE['success']};
        --warning: {PALETTE['warning']};
    }}

    .stApp {{
        background:
          radial-gradient(circle at 10% 0%, rgba(140,147,168,.12), transparent 30%),
          linear-gradient(180deg, #2D2327 0%, #261D21 100%);
    }}

    .block-container {{
        max-width: 1380px;
        padding-top: 1.8rem;
        padding-bottom: 4rem;
    }}

    h1, h2, h3 {{ letter-spacing: -0.025em; }}

    .eyebrow {{
        color: var(--muted);
        text-transform: uppercase;
        font-size: .76rem;
        font-weight: 700;
        letter-spacing: .13em;
        margin-bottom: .35rem;
    }}

    .hero-title {{
        color: var(--text);
        font-size: clamp(2rem, 4vw, 3.35rem);
        line-height: 1.02;
        font-weight: 800;
        margin: 0;
    }}

    .hero-subtitle {{
        color: var(--soft);
        font-size: 1rem;
        margin-top: .7rem;
        max-width: 760px;
    }}

    .section-title {{
        font-size: 1.16rem;
        font-weight: 750;
        color: var(--text);
        margin: .25rem 0 .2rem;
    }}

    .section-subtitle {{
        color: var(--muted);
        font-size: .9rem;
        margin-bottom: .7rem;
    }}

    .metric-card, .today-card, .subject-summary {{
        border: 1px solid rgba(140,147,168,.22);
        background: linear-gradient(145deg, rgba(69,54,75,.72), rgba(45,35,39,.88));
        box-shadow: 0 12px 35px rgba(0,0,0,.10);
        border-radius: 18px;
        padding: 1rem 1.05rem;
    }}

    .metric-label {{
        color: var(--muted);
        font-size: .77rem;
        font-weight: 700;
        letter-spacing: .04em;
        text-transform: uppercase;
    }}

    .metric-value {{
        color: var(--text);
        font-size: 1.75rem;
        font-weight: 800;
        margin-top: .18rem;
        line-height: 1.1;
    }}

    .metric-help {{
        color: var(--soft);
        font-size: .77rem;
        margin-top: .22rem;
    }}

    .today-name {{
        color: var(--text);
        font-size: 1rem;
        font-weight: 750;
        line-height: 1.2;
    }}

    .today-meta {{
        color: var(--muted);
        font-size: .78rem;
        margin-top: .25rem;
    }}

    .today-next {{
        color: var(--soft);
        font-size: .84rem;
        margin-top: .7rem;
    }}

    .pill {{
        display: inline-block;
        color: var(--text);
        background: rgba(140,147,168,.13);
        border: 1px solid rgba(140,147,168,.24);
        padding: .28rem .56rem;
        border-radius: 999px;
        font-size: .73rem;
        margin-right: .3rem;
    }}

    div[data-testid="stMetric"] {{
        background: linear-gradient(145deg, rgba(69,54,75,.70), rgba(45,35,39,.86));
        border: 1px solid rgba(140,147,168,.20);
        padding: .85rem 1rem;
        border-radius: 16px;
    }}

    div[data-testid="stMetric"] label {{ color: var(--muted) !important; }}

    div[data-testid="stProgress"] > div > div > div > div {{
        background: linear-gradient(90deg, var(--muted), #B0A3BA);
    }}

    div[data-testid="stButton"] button {{
        min-height: 2.55rem;
        border-radius: 12px;
        border: 1px solid rgba(140,147,168,.28);
        font-weight: 700;
        transition: transform .12s ease, border-color .12s ease, background .12s ease;
    }}

    div[data-testid="stButton"] button:hover {{
        transform: translateY(-1px);
        border-color: var(--muted);
    }}

    div[data-testid="stTabs"] button[role="tab"] {{
        height: 3.1rem;
        font-weight: 750;
    }}

    div[data-testid="stTabs"] button[aria-selected="true"] {{
        color: #FFFFFF;
    }}

    div[data-testid="stSelectbox"] label,
    div[data-testid="stTextInput"] label {{
        color: var(--soft);
        font-weight: 650;
    }}

    .lesson-hint {{
        color: var(--muted);
        font-size: .83rem;
        padding: .6rem 0 .35rem;
    }}

    .small-note {{
        color: var(--muted);
        font-size: .78rem;
    }}

    @media (max-width: 700px) {{
        .block-container {{ padding: 1.05rem .8rem 3rem; }}
        .hero-title {{ font-size: 2.05rem; }}
        .metric-value {{ font-size: 1.45rem; }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Dados da planilha
# -----------------------------------------------------------------------------
def normalize_text(value: object) -> str:
    text = str(value or "").strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    return text.casefold()


def canonical_day(value: object) -> str:
    raw = str(value or "").strip()
    key = normalize_text(raw)
    aliases = {
        "segunda": "Segunda",
        "segunda-feira": "Segunda",
        "terca": "Terça",
        "terca-feira": "Terça",
        "quarta": "Quarta",
        "quarta-feira": "Quarta",
        "quinta": "Quinta",
        "quinta-feira": "Quinta",
        "sexta": "Sexta",
        "sexta-feira": "Sexta",
        "sabado": "Sábado",
        "domingo": "Domingo",
    }
    return aliases.get(key, raw.title())


def subject_key(name: str, institution: str) -> str:
    # O nome da disciplina é a identidade estável. Assim, trocar instituição ou dia
    # na planilha não apaga o progresso existente.
    raw = normalize_text(name)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def read_sheet(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {path.name}")

    xls = pd.ExcelFile(path, engine="openpyxl")
    preferred = next((s for s in xls.sheet_names if normalize_text(s) == "estudos"), xls.sheet_names[0])
    df = pd.read_excel(path, sheet_name=preferred, engine="openpyxl")

    normalized_cols = {normalize_text(c): c for c in df.columns}
    required = {
        "disciplina": "Disciplina",
        "aulas totais": "Aulas totais",
        "dia": "Dia",
        "instituicao": "Instituição",
    }
    missing = [label for norm, label in required.items() if norm not in normalized_cols]
    if missing:
        raise ValueError("Colunas obrigatórias ausentes: " + ", ".join(missing))

    out = pd.DataFrame(
        {
            "Disciplina": df[normalized_cols["disciplina"]],
            "Aulas totais": df[normalized_cols["aulas totais"]],
            "Dia": df[normalized_cols["dia"]],
            "Instituição": df[normalized_cols["instituicao"]],
        }
    )
    out = out.dropna(subset=["Disciplina"]).copy()
    out["Disciplina"] = out["Disciplina"].astype(str).str.strip()
    out["Instituição"] = out["Instituição"].fillna("").astype(str).str.strip()
    out["Dia"] = out["Dia"].apply(canonical_day)
    out["Aulas totais"] = pd.to_numeric(out["Aulas totais"], errors="coerce").fillna(0).astype(int)
    out = out[out["Aulas totais"] > 0].copy()
    out["key"] = [subject_key(n, i) for n, i in zip(out["Disciplina"], out["Instituição"])]
    out = out.reset_index(drop=True)

    duplicated = out[out["key"].duplicated(keep=False)]
    if not duplicated.empty:
        names = ", ".join(sorted(set(duplicated["Disciplina"].tolist())))
        st.warning(
            "Há nomes de disciplina duplicados. Para preservar o progresso corretamente, "
            f"mantenha cada nome de disciplina único. Duplicadas: {names}."
        )
    return out


# -----------------------------------------------------------------------------
# Persistência do progresso
# -----------------------------------------------------------------------------
def empty_store() -> dict:
    return {"version": 1, "subjects": {}, "last_sync": None}


def load_store() -> dict:
    if not PROGRESS_FILE.exists():
        return empty_store()
    try:
        with FileLock(str(LOCK_FILE), timeout=5):
            with PROGRESS_FILE.open("r", encoding="utf-8") as f:
                data = json.load(f)
        if not isinstance(data, dict) or "subjects" not in data:
            return empty_store()
        return data
    except Exception:
        return empty_store()


def save_store(data: dict) -> None:
    data["version"] = 1
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with FileLock(str(LOCK_FILE), timeout=5):
        fd, temp_path = tempfile.mkstemp(prefix="progresso_", suffix=".json", dir=DATA_DIR)
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            os.replace(temp_path, PROGRESS_FILE)
        finally:
            if os.path.exists(temp_path):
                os.unlink(temp_path)


def reconcile_store(store: dict, df: pd.DataFrame) -> dict:
    subjects = store.setdefault("subjects", {})
    for row in df.to_dict("records"):
        entry = subjects.setdefault(row["key"], {"completed": []})
        entry.setdefault("completed", [])
        entry["meta"] = {
            "disciplina": row["Disciplina"],
            "instituicao": row["Instituição"],
            "dia": row["Dia"],
        }
    store["last_sync"] = datetime.now(TIMEZONE).isoformat(timespec="seconds")
    return store


def completed_lessons(store: dict, key: str, total: int) -> set[int]:
    raw = store.get("subjects", {}).get(key, {}).get("completed", [])
    result = set()
    for item in raw:
        try:
            lesson = int(item)
        except (TypeError, ValueError):
            continue
        if 1 <= lesson <= total:
            result.add(lesson)
    return result


def set_completed(store: dict, key: str, lessons: set[int]) -> None:
    entry = store.setdefault("subjects", {}).setdefault(key, {"completed": []})
    entry["completed"] = sorted(int(x) for x in lessons if int(x) >= 1)
    save_store(store)


def first_missing(completed: set[int], total: int) -> int | None:
    return next((n for n in range(1, total + 1) if n not in completed), None)


# -----------------------------------------------------------------------------
# Estado da aplicação
# -----------------------------------------------------------------------------
try:
    studies = read_sheet(PLANILHA)
except Exception as exc:
    st.error("Não foi possível carregar a planilha de estudos.")
    st.code(str(exc))
    st.info("Mantenha o arquivo Estudos.xlsx na mesma pasta do app.py e preserve as quatro colunas esperadas.")
    st.stop()

if "store" not in st.session_state:
    st.session_state.store = reconcile_store(load_store(), studies)
    save_store(st.session_state.store)

store = st.session_state.store


def refresh_data() -> None:
    global studies, store
    studies = read_sheet(PLANILHA)
    st.session_state.store = reconcile_store(load_store(), studies)
    save_store(st.session_state.store)
    store = st.session_state.store


# -----------------------------------------------------------------------------
# Ações e diálogos
# -----------------------------------------------------------------------------
@st.dialog("Aulas anteriores pendentes", width="small")
def previous_lessons_dialog(key: str, lesson: int, total: int, discipline: str):
    done = completed_lessons(st.session_state.store, key, total)
    previous_pending = [n for n in range(1, lesson) if n not in done]

    st.markdown(f"**{discipline} — aula {lesson}**")
    st.write(
        f"Há {len(previous_pending)} aula(s) anterior(es) ainda não marcadas. "
        f"Deseja considerar todas as aulas de 1 a {lesson} como concluídas?"
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button(f"Marcar 1–{lesson}", type="primary", use_container_width=True):
            done.update(range(1, lesson + 1))
            set_completed(st.session_state.store, key, done)
            st.rerun()
    with c2:
        if st.button(f"Só a aula {lesson}", use_container_width=True):
            done.add(lesson)
            set_completed(st.session_state.store, key, done)
            st.rerun()

    if st.button("Cancelar", use_container_width=True):
        st.rerun()


@st.dialog("Concluir todas as aulas", width="small")
def complete_all_dialog(key: str, total: int, discipline: str):
    st.write(f"Marcar as **{total} aulas** de **{discipline}** como concluídas?")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Concluir tudo", type="primary", use_container_width=True):
            set_completed(st.session_state.store, key, set(range(1, total + 1)))
            st.rerun()
    with c2:
        if st.button("Cancelar", use_container_width=True):
            st.rerun()


@st.dialog("Limpar progresso", width="small")
def clear_subject_dialog(key: str, discipline: str):
    st.write(f"Remover todas as aulas concluídas de **{discipline}**?")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Limpar", type="primary", use_container_width=True):
            set_completed(st.session_state.store, key, set())
            st.rerun()
    with c2:
        if st.button("Cancelar", use_container_width=True):
            st.rerun()


@st.dialog("Concluir todas as disciplinas", width="small")
def complete_everything_dialog():
    total_classes = int(studies["Aulas totais"].sum())
    st.write(
        f"Esta ação marcará **{total_classes} aulas** de **{len(studies)} disciplinas** como concluídas."
    )
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Concluir tudo", type="primary", use_container_width=True):
            for row in studies.to_dict("records"):
                entry = st.session_state.store.setdefault("subjects", {}).setdefault(row["key"], {})
                entry["completed"] = list(range(1, int(row["Aulas totais"]) + 1))
            save_store(st.session_state.store)
            st.rerun()
    with c2:
        if st.button("Cancelar", use_container_width=True):
            st.rerun()


def toggle_lesson(row: dict, lesson: int) -> None:
    key = row["key"]
    total = int(row["Aulas totais"])
    done = completed_lessons(st.session_state.store, key, total)

    if lesson in done:
        done.remove(lesson)
        set_completed(st.session_state.store, key, done)
        st.rerun()
        return

    previous_pending = [n for n in range(1, lesson) if n not in done]
    if previous_pending:
        previous_lessons_dialog(key, lesson, total, row["Disciplina"])
    else:
        done.add(lesson)
        set_completed(st.session_state.store, key, done)
        st.rerun()


def mark_next(row: dict) -> None:
    total = int(row["Aulas totais"])
    done = completed_lessons(st.session_state.store, row["key"], total)
    nxt = first_missing(done, total)
    if nxt is not None:
        done.add(nxt)
        set_completed(st.session_state.store, row["key"], done)
        st.rerun()


# -----------------------------------------------------------------------------
# Cálculos
# -----------------------------------------------------------------------------
def row_stats(row: dict) -> dict:
    total = int(row["Aulas totais"])
    done = completed_lessons(st.session_state.store, row["key"], total)
    count = len(done)
    return {
        "done": done,
        "count": count,
        "remaining": max(total - count, 0),
        "pct": (count / total * 100) if total else 0,
        "next": first_missing(done, total),
    }


def aggregate_stats(df: pd.DataFrame) -> dict:
    total = 0
    done = 0
    completed_subjects = 0
    for row in df.to_dict("records"):
        stats = row_stats(row)
        t = int(row["Aulas totais"])
        total += t
        done += stats["count"]
        if stats["count"] == t:
            completed_subjects += 1
    return {
        "total": total,
        "done": done,
        "remaining": max(total - done, 0),
        "pct": (done / total * 100) if total else 0,
        "completed_subjects": completed_subjects,
    }


now = datetime.now(TIMEZONE)
today_name = DIA_MAP[now.weekday()]
today_df = studies[studies["Dia"] == today_name].copy()
overall = aggregate_stats(studies)

# -----------------------------------------------------------------------------
# Cabeçalho
# -----------------------------------------------------------------------------
header_left, header_right = st.columns([4.6, 1.4], vertical_alignment="center")
with header_left:
    st.markdown('<div class="eyebrow">Painel pessoal</div>', unsafe_allow_html=True)
    st.markdown('<h1 class="hero-title">Estudos</h1>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="hero-subtitle">Acompanhe seu avanço, registre aulas concluídas e veja o foco de hoje. '
        f'Hoje é <b>{today_name.lower()}</b>, {now.strftime("%d/%m/%Y")}.</div>',
        unsafe_allow_html=True,
    )
with header_right:
    if st.button("↻  Sincronizar planilha", use_container_width=True, help="Relê Estudos.xlsx e preserva o progresso das disciplinas existentes."):
        try:
            refresh_data()
            st.toast("Planilha sincronizada.", icon="✓")
            st.rerun()
        except Exception as exc:
            st.error(str(exc))

st.write("")

# KPIs
m1, m2, m3, m4 = st.columns(4)
metrics = [
    (m1, "Progresso geral", f"{overall['pct']:.0f}%", f"{overall['done']} de {overall['total']} aulas"),
    (m2, "Aulas concluídas", str(overall["done"]), "Somando todas as disciplinas"),
    (m3, "Aulas restantes", str(overall["remaining"]), "Para concluir o plano atual"),
    (m4, "Disciplinas", str(len(studies)), f"{overall['completed_subjects']} concluída(s)"),
]
for col, label, value, help_text in metrics:
    with col:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">{label}</div>'
            f'<div class="metric-value">{value}</div><div class="metric-help">{help_text}</div></div>',
            unsafe_allow_html=True,
        )

st.write("")

# -----------------------------------------------------------------------------
# Matérias do dia
# -----------------------------------------------------------------------------
st.markdown('<div class="section-title">Matérias do dia</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="section-subtitle">Prioridade de {today_name.lower()} com acesso rápido à próxima aula.</div>',
    unsafe_allow_html=True,
)

if today_df.empty:
    st.info("Não há disciplinas programadas para hoje.")
else:
    cols = st.columns(min(3, len(today_df)))
    for idx, row in enumerate(today_df.to_dict("records")):
        stats = row_stats(row)
        with cols[idx % len(cols)]:
            next_text = "Plano concluído" if stats["next"] is None else f"Próxima: aula {stats['next']}"
            st.markdown(
                f'<div class="today-card"><div class="today-name">{row["Disciplina"]}</div>'
                f'<div class="today-meta">{row["Instituição"]}</div>'
                f'<div class="today-next">{next_text} · {stats["count"]}/{int(row["Aulas totais"])} concluídas</div></div>',
                unsafe_allow_html=True,
            )
            st.progress(stats["pct"] / 100)
            if stats["next"] is not None:
                if st.button(
                    f"✓ Concluir aula {stats['next']}",
                    key=f"today_next_{row['key']}",
                    use_container_width=True,
                ):
                    mark_next(row)
            else:
                st.button("✓ Concluído", key=f"today_done_{row['key']}", disabled=True, use_container_width=True)

st.write("")

# -----------------------------------------------------------------------------
# Abas
# -----------------------------------------------------------------------------
tab_progress, tab_overview = st.tabs(["✓ Registrar aulas", "◔ Visão geral"])

with tab_progress:
    st.write("")
    controls_left, controls_next, controls_all, controls_clear = st.columns([2.7, 1.25, 1.05, 1.0], vertical_alignment="bottom")

    with controls_left:
        labels = studies["Disciplina"].tolist()
        default_index = 0
        if not today_df.empty:
            first_today = today_df.iloc[0]["Disciplina"]
            if first_today in labels:
                default_index = labels.index(first_today)
        selected_name = st.selectbox("Disciplina", labels, index=default_index, key="subject_selector")

    selected_row = studies[studies["Disciplina"] == selected_name].iloc[0].to_dict()
    selected_stats = row_stats(selected_row)
    total_selected = int(selected_row["Aulas totais"])

    with controls_next:
        if selected_stats["next"] is not None:
            if st.button(
                f"✓ Próxima ({selected_stats['next']})",
                type="primary",
                use_container_width=True,
            ):
                mark_next(selected_row)
        else:
            st.button("✓ Concluída", disabled=True, use_container_width=True)

    with controls_all:
        if st.button("Concluir todas", use_container_width=True):
            complete_all_dialog(selected_row["key"], total_selected, selected_row["Disciplina"])

    with controls_clear:
        if st.button("Limpar", use_container_width=True):
            clear_subject_dialog(selected_row["key"], selected_row["Disciplina"])

    st.write("")

    # Resumo da disciplina
    summary_a, summary_b, summary_c = st.columns([3.2, 1, 1])
    with summary_a:
        st.markdown(
            f'<div class="subject-summary"><span class="pill">{selected_row["Dia"]}</span>'
            f'<span class="pill">{selected_row["Instituição"]}</span>'
            f'<div style="font-size:1.35rem;font-weight:800;color:#fff;margin-top:.7rem">{selected_row["Disciplina"]}</div>'
            f'<div style="color:{PALETTE["soft"]};font-size:.88rem;margin-top:.25rem">'
            f'{selected_stats["count"]} de {total_selected} aulas concluídas · {selected_stats["pct"]:.0f}%</div></div>',
            unsafe_allow_html=True,
        )
    with summary_b:
        st.metric("Concluídas", selected_stats["count"])
    with summary_c:
        st.metric("Restantes", selected_stats["remaining"])

    st.progress(selected_stats["pct"] / 100)

    st.markdown(
        '<div class="lesson-hint">Clique em uma aula para marcar ou desmarcar. '
        'Se você escolher uma aula deixando anteriores pendentes, o painel perguntará se deseja preencher o intervalo.</div>',
        unsafe_allow_html=True,
    )

    # Grade de aulas: 8 por linha em desktop; os botões se adaptam à largura disponível.
    buttons_per_row = 8
    for start in range(1, total_selected + 1, buttons_per_row):
        cols = st.columns(buttons_per_row)
        for offset, lesson in enumerate(range(start, min(start + buttons_per_row, total_selected + 1))):
            with cols[offset]:
                is_done = lesson in selected_stats["done"]
                label = f"✓ {lesson}" if is_done else str(lesson)
                if st.button(
                    label,
                    key=f"lesson_{selected_row['key']}_{lesson}",
                    type="primary" if is_done else "secondary",
                    use_container_width=True,
                    help="Clique para desmarcar" if is_done else "Clique para marcar como concluída",
                ):
                    toggle_lesson(selected_row, lesson)

with tab_overview:
    st.write("")

    # Um gráfico de rosca por disciplina.
    # O ângulo dourado gera uma cor pastel exclusiva para cada matéria,
    # sem repetir tons dentro da lista atual de disciplinas.
    overview_subjects = studies.to_dict("records")
    chart_columns = 3

    for start_idx in range(0, len(overview_subjects), chart_columns):
        cols = st.columns(chart_columns, gap="large")
        batch = overview_subjects[start_idx : start_idx + chart_columns]

        for offset, row in enumerate(batch):
            subject_index = start_idx + offset
            stats = row_stats(row)
            total = int(row["Aulas totais"])

            # Distribuição de matizes pelo ângulo dourado: visualmente distinta
            # mesmo quando novas disciplinas são acrescentadas à planilha.
            hue = (subject_index * 137.508 + 215) % 360
            subject_color = f"hsl({hue:.2f}, 58%, 76%)"

            with cols[offset]:
                fig = go.Figure(
                    data=[
                        go.Pie(
                            labels=["Concluídas", "Restantes"],
                            values=[stats["count"], stats["remaining"]],
                            hole=0.72,
                            sort=False,
                            direction="clockwise",
                            marker=dict(
                                colors=[subject_color, "rgba(140,147,168,0.18)"],
                                line=dict(width=0),
                            ),
                            textinfo="none",
                            hovertemplate="%{label}: %{value} aula(s)<extra></extra>",
                        )
                    ]
                )

                fig.add_annotation(
                    text=(
                        f"<b>{stats['pct']:.0f}%</b>"
                        f"<br><span style='font-size:12px;color:{PALETTE['soft']}'>"
                        f"{stats['count']}/{total} aulas</span>"
                    ),
                    x=0.5,
                    y=0.5,
                    showarrow=False,
                    font=dict(color=PALETTE["text"], size=24),
                    align="center",
                )

                fig.update_layout(
                    title=dict(
                        text=f"<b>{row['Disciplina']}</b>",
                        x=0.5,
                        xanchor="center",
                        y=0.97,
                        font=dict(color=PALETTE["text"], size=16),
                    ),
                    height=300,
                    margin=dict(l=8, r=8, t=55, b=8),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    showlegend=False,
                    font=dict(color=PALETTE["soft"]),
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True,
                    config={"displayModeBar": False},
                    key=f"overview_donut_{row['key']}",
                )

st.markdown(
    '<div class="small-note" style="margin-top:2rem">Fonte de dados: Estudos.xlsx · O progresso fica em data/progresso.json.</div>',
    unsafe_allow_html=True,
)
