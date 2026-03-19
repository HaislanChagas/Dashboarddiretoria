import math
import unicodedata

import gspread
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Dashboard Executivo do Diretor", page_icon="📊", layout="wide")
px.defaults.template = "plotly_dark"

# =========================
# CONFIG
# =========================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1mC29Jya_c5KTVZpHqfRSDLkuVWv4_ASYNISEPQzN2xQ/edit#gid=0"
SHEET_INDICADORES = "1WEA9Mhc8kPj8LdN8h_eB264nLlcI4fZY0lDMjrEsn60"
ABA_BASE_DASHBOARD = "BASE_DASHBOARD"

# COLE AQUI O ID DA PLANILHA GOOGLE SHEETS DA ROLETA
# Exemplo: "1AbCDefGhIJklmnOPqRstUVwxYZ1234567890"
SHEET_ROLETA = "15HV7fUtCJ4AN5kG81fgJjLVf29wLsHQsrQ5H1m0Vq7U"
ABA_ROLETA = "ROLETA"

CACHE_TTL = 300
ABAS_EXCLUIDAS_RANKING = {"Funil de Vendas Geral"}
ETAPAS = ["Leads", "Pasta", "Aprovação", "Proposta", "Venda"]


# =========================
# CSS
# =========================
def load_css():
    with open("style.css", encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)


load_css()


# =========================
# HELPERS
# =========================
def normalizar_texto(texto) -> str:
    if texto is None:
        return ""
    texto = str(texto).strip().lower()
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")


def numero(valor):
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and math.isnan(valor):
            return 0.0
        return float(valor)
    s = str(valor).strip()
    if s == "" or "#DIV/0!" in s.upper():
        return 0.0
    s = s.replace("R$", "").replace("%", "").replace(" ", "")
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def safe_div(a, b):
    return a / b if b not in (0, None) else 0.0


def fmt_int(v):
    return f"{int(round(numero(v))):,}".replace(",", ".")


def fmt_num(v, dec=1):
    return f"{numero(v):,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(v, dec=1):
    return f"{numero(v) * 100:.{dec}f}%".replace(".", ",")


def fmt_money(v, dec=0):
    return f"R$ {numero(v):,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def card_kpi(titulo: str, valor: str, nota: str = ""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">{titulo}</div>
            <div class="kpi-value">{valor}</div>
            <div class="kpi-note">{nota}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def estilizar_fig(fig):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#0b1224",
        font=dict(color="#f3f7ff", family="Inter, Segoe UI, Arial"),
        title_font=dict(size=18, color="#f3f7ff"),
        margin=dict(l=20, r=20, t=55, b=20),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(color="#cbd5e1"),
            bgcolor="rgba(0,0,0,0)"
        ),
    )
    fig.update_xaxes(
        showgrid=True,
        gridcolor="rgba(255,255,255,0.06)",
        zeroline=False,
        linecolor="rgba(255,255,255,0.10)",
        tickfont=dict(color="#cbd5e1"),
        title_font=dict(color="#cbd5e1")
    )
    fig.update_yaxes(
        showgrid=True,
        gridcolor="rgba(255,255,255,0.06)",
        zeroline=False,
        linecolor="rgba(255,255,255,0.10)",
        tickfont=dict(color="#cbd5e1"),
        title_font=dict(color="#cbd5e1")
    )
    return fig


# =========================
# GOOGLE SHEETS
# =========================
@st.cache_resource
def conectar_google():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets.readonly",
            "https://www.googleapis.com/auth/drive.readonly",
        ],
    )
    return gspread.authorize(creds)


@st.cache_resource
def abrir_planilha():
    gc = conectar_google()
    return gc.open_by_url(SHEET_URL)


@st.cache_data(ttl=CACHE_TTL)
def listar_abas():
    sh = abrir_planilha()
    return [ws.title for ws in sh.worksheets()]


@st.cache_data(ttl=CACHE_TTL)
def carregar_aba_raw(nome_aba: str):
    sh = abrir_planilha()
    ws = sh.worksheet(nome_aba)
    return ws.get_all_values()


# =========================
# PARSER FUNIL
# =========================
def detectar_coluna_ancora(valores):
    max_cols = max(len(l) for l in valores) if valores else 0
    for c in range(max_cols):
        amostras = []
        for r in [0, 1, 2, 5, 6, 7]:
            if r < len(valores) and c < len(valores[r]):
                amostras.append(normalizar_texto(valores[r][c]))
        bloco = " | ".join(amostras)
        if any(chave in bloco for chave in ["mes", "gerente", "funil de vendas", "marco", "março"]):
            return c
    return 0


def get_cell(valores, r, c):
    if r < 0 or c < 0:
        return ""
    if r >= len(valores) or c >= len(valores[r]):
        return ""
    return valores[r][c]


def parse_aba_funil(nome_aba, valores):
    if not valores:
        return None, pd.DataFrame()

    anc = detectar_coluna_ancora(valores)

    mes_nome = get_cell(valores, 0, anc)
    dias_mes = numero(get_cell(valores, 1, anc + 1))
    meta_mes = numero(get_cell(valores, 1, anc + 2))
    dia_atual = numero(get_cell(valores, 2, anc + 1))
    gerente = get_cell(valores, 5, anc + 1)

    resumo = {
        "aba": nome_aba,
        "mes_nome": mes_nome,
        "dias_mes": dias_mes,
        "meta_mes": meta_mes,
        "dia_atual": dia_atual,
        "gerente": gerente,
    }

    linhas = []
    etapa_row_start = 7

    for i, etapa in enumerate(ETAPAS):
        r = etapa_row_start + i
        etapa_nome = get_cell(valores, r, anc)
        if normalizar_texto(etapa_nome) != normalizar_texto(etapa):
            etapa_nome = etapa

        linhas.append(
            {
                "aba": nome_aba,
                "gerente": gerente,
                "mes_nome": mes_nome,
                "dias_mes": dias_mes,
                "meta_mes": meta_mes,
                "dia_atual": dia_atual,
                "etapa": etapa_nome,
                "target_total": numero(get_cell(valores, r, anc + 1)),
                "target_pct": numero(get_cell(valores, r, anc + 2)),
                "expected_now": numero(get_cell(valores, r, anc + 5)),
                "expected_pct": numero(get_cell(valores, r, anc + 6)),
                "actual_now": numero(get_cell(valores, r, anc + 9)),
                "actual_pct": numero(get_cell(valores, r, anc + 10)),
                "gap": numero(get_cell(valores, r, anc + 13)),
                "atingimento": safe_div(
                    numero(get_cell(valores, r, anc + 9)),
                    numero(get_cell(valores, r, anc + 5)),
                ),
            }
        )

    return resumo, pd.DataFrame(linhas)


@st.cache_data(ttl=CACHE_TTL)
def carregar_dados_dashboard():
    abas = listar_abas()
    resumos, frames, erros = [], [], []

    for aba in abas:
        try:
            valores = carregar_aba_raw(aba)
            resumo, df = parse_aba_funil(aba, valores)
            if resumo:
                resumos.append(resumo)
            if not df.empty:
                frames.append(df)
        except Exception as e:
            erros.append(f"{aba}: {type(e).__name__}: {e}")

    return pd.DataFrame(resumos), (pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()), erros


# =========================
# INDICADORES - BASE_DASHBOARD
# =========================
@st.cache_data(ttl=CACHE_TTL)
def carregar_indicadores():
    try:
        gc = conectar_google()
        sh = gc.open_by_key(SHEET_INDICADORES)
        ws = sh.worksheet(ABA_BASE_DASHBOARD)

        valores = ws.get_all_values()
        if not valores or len(valores) < 2:
            return pd.DataFrame(), f"A aba {ABA_BASE_DASHBOARD} está vazia ou sem linhas suficientes."

        header = [str(x).strip() for x in valores[0]]
        linhas = valores[1:]
        linhas = [linha for linha in linhas if any(str(c).strip() != "" for c in linha)]

        if not linhas:
            return pd.DataFrame(), f"A aba {ABA_BASE_DASHBOARD} não possui dados abaixo do cabeçalho."

        max_cols = len(header)
        linhas_pad = []
        for linha in linhas:
            if len(linha) < max_cols:
                linha = linha + [""] * (max_cols - len(linha))
            else:
                linha = linha[:max_cols]
            linhas_pad.append(linha)

        df = pd.DataFrame(linhas_pad, columns=header)

        mapa_colunas = {normalizar_texto(c): c for c in df.columns}

        obrigatorias = {
            "Mes": ["mes", "mês"],
            "Equipe": ["equipe", "time", "operacao", "operação", "gerente"],
            "Vendas": ["vendas", "venda"],
            "VGV": ["vgv"],
        }

        opcionais = {
            "Conversao": ["conversao", "conversão"],
            "Corretor_Ativo": ["corretor_ativo", "corretor ativo", "corretorativo"],
            "Roleta": ["roleta"],
            "Roleta_por_Ativo": ["roleta_por_ativo", "roleta por ativo"],
            "IPC": ["ipc"],
            "Equipe_Produtiva": ["equipe_produtiva", "equipe produtiva"],
            "Equipe_Produtiva_Rate": ["equipe_produtiva_rate", "equipe produtiva rate", "equipe produtiva %"],
            "Quarentena": ["quarentena"],
        }

        mapeadas = {}

        for nome_padrao, aliases in obrigatorias.items():
            encontrada = None
            for alias in aliases:
                if alias in mapa_colunas:
                    encontrada = mapa_colunas[alias]
                    break
            if not encontrada:
                return pd.DataFrame(), f"Coluna obrigatória ausente em {ABA_BASE_DASHBOARD}: {nome_padrao}. Colunas encontradas: {list(df.columns)}"
            mapeadas[nome_padrao] = encontrada

        for nome_padrao, aliases in opcionais.items():
            encontrada = None
            for alias in aliases:
                if alias in mapa_colunas:
                    encontrada = mapa_colunas[alias]
                    break
            mapeadas[nome_padrao] = encontrada

        colunas_usar = [mapeadas["Mes"], mapeadas["Equipe"], mapeadas["Vendas"], mapeadas["VGV"]]
        for nome_opt in opcionais:
            if mapeadas[nome_opt]:
                colunas_usar.append(mapeadas[nome_opt])

        df = df[colunas_usar].copy()

        rename_map = {
            mapeadas["Mes"]: "Mes",
            mapeadas["Equipe"]: "Equipe",
            mapeadas["Vendas"]: "Vendas",
            mapeadas["VGV"]: "VGV",
        }
        for nome_opt in opcionais:
            if mapeadas[nome_opt]:
                rename_map[mapeadas[nome_opt]] = nome_opt

        df = df.rename(columns=rename_map)

        for col in [
            "Vendas", "VGV", "Conversao", "Corretor_Ativo", "Roleta",
            "Roleta_por_Ativo", "IPC", "Equipe_Produtiva",
            "Equipe_Produtiva_Rate", "Quarentena"
        ]:
            if col in df.columns:
                df[col] = df[col].apply(numero)

        df["Mes"] = df["Mes"].astype(str).str.strip()
        df["Equipe"] = df["Equipe"].astype(str).str.strip()

        df = df[(df["Mes"] != "") & (df["Equipe"] != "")].copy()
        df = df[(df["Vendas"] > 0) | (df["VGV"] > 0) | (df.get("Roleta", 0) > 0)].copy()

        df["Ticket_Medio"] = df.apply(lambda x: safe_div(x["VGV"], x["Vendas"]), axis=1)

        ordem_meses = {
            "jan": 1, "janeiro": 1,
            "fev": 2, "fevereiro": 2,
            "mar": 3, "marco": 3, "março": 3,
            "abr": 4, "abril": 4,
            "mai": 5, "maio": 5,
            "jun": 6, "junho": 6,
            "jul": 7, "julho": 7,
            "ago": 8, "agosto": 8,
            "set": 9, "setembro": 9,
            "out": 10, "outubro": 10,
            "nov": 11, "novembro": 11,
            "dez": 12, "dezembro": 12,
        }

        df["ordem_mes"] = df["Mes"].apply(lambda x: ordem_meses.get(normalizar_texto(x), 999))
        df = df.sort_values(["ordem_mes", "Equipe"]).reset_index(drop=True)

        return df, None

    except Exception as e:
        return pd.DataFrame(), f"{type(e).__name__}: {e}"


# =========================
# ROLETA DIÁRIA
# =========================
@st.cache_data(ttl=CACHE_TTL)
def carregar_roleta_diaria():
    try:
        if not SHEET_ROLETA or SHEET_ROLETA == "COLE_AQUI_O_ID_DA_PLANILHA_ROLETA":
            return pd.DataFrame(), "Configure o ID da planilha da roleta em SHEET_ROLETA."

        gc = conectar_google()
        sh = gc.open_by_key(SHEET_ROLETA)
        ws = sh.worksheet(ABA_ROLETA)

        valores = ws.get_all_values()
        if not valores or len(valores) < 3:
            return pd.DataFrame(), "A aba ROLETA está vazia ou sem dados suficientes."

        # No arquivo exemplo, o cabeçalho real estava na 2ª linha (índice 1)
        # Aqui tentamos detectar automaticamente entre as primeiras linhas.
        header_idx = None
        for i, linha in enumerate(valores[:10]):
            linha_norm = [normalizar_texto(c) for c in linha]
            texto_linha = " | ".join(linha_norm)
            if "data" in texto_linha and ("total" in texto_linha or "roleta" in texto_linha):
                header_idx = i
                break

        if header_idx is None:
            return pd.DataFrame(), f"Não encontrei o cabeçalho da aba ROLETA. Primeiras linhas: {valores[:5]}"

        header = [str(x).strip() for x in valores[header_idx]]
        linhas = valores[header_idx + 1:]

        linhas = [linha for linha in linhas if any(str(c).strip() != "" for c in linha)]
        if not linhas:
            return pd.DataFrame(), "Não há linhas de dados na aba ROLETA."

        max_cols = len(header)
        linhas_pad = []
        for linha in linhas:
            if len(linha) < max_cols:
                linha = linha + [""] * (max_cols - len(linha))
            else:
                linha = linha[:max_cols]
            linhas_pad.append(linha)

        df = pd.DataFrame(linhas_pad, columns=header)
        mapa = {normalizar_texto(c): c for c in df.columns}

        col_data = mapa.get("data")
        col_manha = mapa.get("roleta manha") or mapa.get("roleta manhã")
        col_total = mapa.get("total")
        col_noite = mapa.get("roleta noite")
        col_rn_total = mapa.get("r.n total") or mapa.get("rn total")

        if not col_data or not col_total:
            return pd.DataFrame(), f"Colunas encontradas na roleta: {list(df.columns)}"

        cols = [col_data, col_total]
        if col_manha:
            cols.append(col_manha)
        if col_noite:
            cols.append(col_noite)
        if col_rn_total:
            cols.append(col_rn_total)

        df = df[cols].copy()

        rename_map = {
            col_data: "Dia",
            col_total: "Roleta_Total",
        }
        if col_manha:
            rename_map[col_manha] = "Roleta_Manha"
        if col_noite:
            rename_map[col_noite] = "Roleta_Noite"
        if col_rn_total:
            rename_map[col_rn_total] = "RN_Total"

        df = df.rename(columns=rename_map)

        for col in df.columns:
            if col != "Dia":
                df[col] = df[col].apply(numero)

        df["Dia"] = df["Dia"].apply(numero)
        df = df[df["Dia"] > 0].copy()
        df["Dia"] = df["Dia"].astype(int)
        df = df.sort_values("Dia").reset_index(drop=True)

        return df, None

    except Exception as e:
        return pd.DataFrame(), f"{type(e).__name__}: {e}"


def valor_etapa(df, etapa, coluna):
    temp = df[df["etapa"] == etapa]
    if temp.empty:
        return 0.0
    return numero(temp.iloc[0][coluna])


# =========================
# RENDER - PERFORMANCE FINANCEIRA
# =========================
def render_financeiro():
    df_indicadores, erro_ind = carregar_indicadores()

    if erro_ind:
        st.warning(f"Indicadores: {erro_ind}")

    if df_indicadores.empty:
        st.warning("Sem dados financeiros disponíveis.")
        return

    consolidado_mes = (
        df_indicadores.groupby(["Mes", "ordem_mes"], as_index=False)
        .agg({"Vendas": "sum", "VGV": "sum"})
        .sort_values("ordem_mes")
        .reset_index(drop=True)
    )
    consolidado_mes["Ticket_Medio"] = consolidado_mes.apply(
        lambda x: safe_div(x["VGV"], x["Vendas"]), axis=1
    )

    ultimo = consolidado_mes.iloc[-1]
    vendas_mes = numero(ultimo["Vendas"])
    vgv_mes = numero(ultimo["VGV"])
    ticket_medio = safe_div(vgv_mes, vendas_mes)

    if len(consolidado_mes) > 1:
        penultimo = consolidado_mes.iloc[-2]
        crescimento_vgv = safe_div(vgv_mes - numero(penultimo["VGV"]), numero(penultimo["VGV"]))
        crescimento_vendas = safe_div(vendas_mes - numero(penultimo["Vendas"]), numero(penultimo["Vendas"]))
    else:
        crescimento_vgv = 0.0
        crescimento_vendas = 0.0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        card_kpi("Vendas do mês", fmt_int(vendas_mes), f"Mês: {ultimo['Mes']}")
    with c2:
        card_kpi("VGV do mês", fmt_money(vgv_mes, 0), "Resultado financeiro atual")
    with c3:
        card_kpi("Ticket médio", fmt_money(ticket_medio, 0), "VGV / vendas")
    with c4:
        card_kpi("Crescimento VGV", fmt_pct(crescimento_vgv), f"Vendas: {fmt_pct(crescimento_vendas)}")

    g1, g2 = st.columns(2)

    with g1:
        fig_vgv = px.line(
            consolidado_mes,
            x="Mes",
            y="VGV",
            markers=True,
            title="Evolução mensal do VGV",
        )
        fig_vgv.update_traces(line_color="#4dd099", marker_color="#4dd099")
        st.plotly_chart(estilizar_fig(fig_vgv), use_container_width=True)

    with g2:
        fig_vendas = px.bar(
            consolidado_mes,
            x="Mes",
            y="Vendas",
            title="Volume mensal de vendas",
            text_auto=".0f",
            color_discrete_sequence=["#143554"],
        )
        st.plotly_chart(estilizar_fig(fig_vendas), use_container_width=True)

    fig_ticket = px.line(
        consolidado_mes,
        x="Mes",
        y="Ticket_Medio",
        markers=True,
        title="Evolução do ticket médio",
    )
    fig_ticket.update_traces(line_color="#60a5fa", marker_color="#60a5fa")
    st.plotly_chart(estilizar_fig(fig_ticket), use_container_width=True)

    st.markdown("<div class='section-title'>Ranking por equipe</div>", unsafe_allow_html=True)

    ranking = (
        df_indicadores
        .groupby("Equipe", as_index=False)
        .agg({
            "Vendas": "sum",
            "VGV": "sum"
        })
        .sort_values("VGV", ascending=False)
        .reset_index(drop=True)
    )
    ranking["Ticket"] = ranking.apply(lambda x: safe_div(x["VGV"], x["Vendas"]), axis=1)

    r1, r2 = st.columns(2)

    with r1:
        fig_rank_vgv = px.bar(
            ranking.sort_values("VGV", ascending=True),
            x="VGV",
            y="Equipe",
            orientation="h",
            title="Ranking por VGV",
            text_auto=".0f",
            color_discrete_sequence=["#4dd099"]
        )
        st.plotly_chart(estilizar_fig(fig_rank_vgv), use_container_width=True)

    with r2:
        fig_rank_vendas = px.bar(
            ranking.sort_values("Vendas", ascending=True),
            x="Vendas",
            y="Equipe",
            orientation="h",
            title="Ranking por Volume de Vendas",
            text_auto=".0f",
            color_discrete_sequence=["#143554"]
        )
        st.plotly_chart(estilizar_fig(fig_rank_vendas), use_container_width=True)

    st.dataframe(
        ranking.rename(columns={
            "Equipe": "Equipe",
            "Vendas": "Vendas",
            "VGV": "VGV",
            "Ticket": "Ticket Médio"
        }),
        use_container_width=True,
        height=260,
    )

    if {"Mes", "Equipe", "VGV"}.issubset(df_indicadores.columns):
        st.markdown("<div class='section-title'>VGV por equipe e mês</div>", unsafe_allow_html=True)
        pivot_vgv = (
            df_indicadores
            .pivot_table(index="Mes", columns="Equipe", values="VGV", aggfunc="sum", fill_value=0)
            .reset_index()
        )
        fig_comp = px.bar(
            pivot_vgv,
            x="Mes",
            y=[col for col in pivot_vgv.columns if col != "Mes"],
            barmode="group",
            title="Comparativo mensal de VGV por equipe",
        )
        st.plotly_chart(estilizar_fig(fig_comp), use_container_width=True)

    st.markdown("<div class='section-title'>Leitura executiva financeira</div>", unsafe_allow_html=True)

    if crescimento_vgv < 0:
        st.markdown(
            f"""
            <div class="alert-bad">
                <b>Alerta:</b> o VGV caiu {fmt_pct(abs(crescimento_vgv))} em relação ao mês anterior.<br>
                Revise volume, mix de produto e taxa de fechamento.
            </div>
            """,
            unsafe_allow_html=True,
        )
    elif crescimento_vgv > 0.15:
        st.markdown(
            f"""
            <div class="alert-good">
                <b>Destaque:</b> o VGV cresceu {fmt_pct(crescimento_vgv)} em relação ao mês anterior.<br>
                Mantenha a estratégia que está puxando resultado.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <div class="alert-warn">
                <b>Status:</b> crescimento financeiro estável.<br>
                Há espaço para acelerar volume e ticket médio.
            </div>
            """,
            unsafe_allow_html=True,
        )

    tabela_cols = ["Mes", "Equipe", "Vendas", "VGV", "Ticket_Medio"]
    for extra in ["Conversao", "Corretor_Ativo", "Roleta", "IPC", "Equipe_Produtiva", "Quarentena"]:
        if extra in df_indicadores.columns:
            tabela_cols.append(extra)

    tabela_fin = df_indicadores[tabela_cols].copy()
    tabela_fin = tabela_fin.rename(columns={
        "Mes": "Mês",
        "Equipe": "Equipe",
        "Vendas": "Vendas",
        "VGV": "VGV",
        "Ticket_Medio": "Ticket Médio",
        "Conversao": "Conversão",
        "Corretor_Ativo": "Corretor Ativo",
        "Roleta": "Roleta",
        "IPC": "IPC",
        "Equipe_Produtiva": "Equipe Produtiva",
        "Quarentena": "Quarentena",
    })
    st.dataframe(tabela_fin, use_container_width=True, height=300)


# =========================
# RENDER - ROLETA
# =========================
def render_roleta():
    df_roleta, erro_roleta = carregar_roleta_diaria()
    df_indicadores, erro_ind = carregar_indicadores()

    if erro_roleta:
        st.warning(f"Roleta diária: {erro_roleta}")
    if erro_ind:
        st.warning(f"Indicadores: {erro_ind}")

    if df_roleta.empty:
        st.warning("Sem dados da roleta diária.")
        return

    roleta_total_mes = numero(df_roleta["Roleta_Total"].sum())
    roleta_manha = numero(df_roleta["Roleta_Manha"].sum()) if "Roleta_Manha" in df_roleta.columns else 0
    roleta_noite = numero(df_roleta["Roleta_Noite"].sum()) if "Roleta_Noite" in df_roleta.columns else 0
    rn_total = numero(df_roleta["RN_Total"].sum()) if "RN_Total" in df_roleta.columns else 0

    vendas_mes = 0
    vgv_mes = 0
    conv_roleta = 0
    vgv_por_roleta = 0

    if not df_indicadores.empty:
        consolidado_mes = (
            df_indicadores.groupby(["Mes", "ordem_mes"], as_index=False)
            .agg({"Vendas": "sum", "VGV": "sum"})
            .sort_values("ordem_mes")
            .reset_index(drop=True)
        )

        if not consolidado_mes.empty:
            ultimo = consolidado_mes.iloc[-1]
            vendas_mes = numero(ultimo["Vendas"])
            vgv_mes = numero(ultimo["VGV"])
            conv_roleta = safe_div(vendas_mes, roleta_total_mes)
            vgv_por_roleta = safe_div(vgv_mes, roleta_total_mes)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        card_kpi("Roleta total mês", fmt_int(roleta_total_mes), "Total acumulado")
    with c2:
        card_kpi("Roleta manhã", fmt_int(roleta_manha), "Volume manhã")
    with c3:
        card_kpi("Roleta noite", fmt_int(roleta_noite), "Volume noite")
    with c4:
        card_kpi("RN total", fmt_int(rn_total), "Retorno / RN acumulado")

    c5, c6, c7, c8 = st.columns(4)
    with c5:
        card_kpi("Vendas do mês", fmt_int(vendas_mes), "Base financeira")
    with c6:
        card_kpi("Conversão da roleta", fmt_pct(conv_roleta), "Vendas / roleta")
    with c7:
        card_kpi("VGV do mês", fmt_money(vgv_mes, 0), "Base financeira")
    with c8:
        card_kpi("VGV por roleta", fmt_money(vgv_por_roleta, 0), "VGV / roleta")

    g1, g2 = st.columns(2)

    with g1:
        fig_roleta_total = px.line(
            df_roleta,
            x="Dia",
            y="Roleta_Total",
            markers=True,
            title="Evolução diária da roleta total",
        )
        fig_roleta_total.update_traces(line_color="#4dd099", marker_color="#4dd099")
        st.plotly_chart(estilizar_fig(fig_roleta_total), use_container_width=True)

    with g2:
        cols_plot = []
        if "Roleta_Manha" in df_roleta.columns:
            cols_plot.append("Roleta_Manha")
        if "Roleta_Noite" in df_roleta.columns:
            cols_plot.append("Roleta_Noite")

        if cols_plot:
            fig_turnos = px.bar(
                df_roleta,
                x="Dia",
                y=cols_plot,
                barmode="group",
                title="Roleta por turno",
            )
            st.plotly_chart(estilizar_fig(fig_turnos), use_container_width=True)

    if "RN_Total" in df_roleta.columns:
        fig_rn = px.line(
            df_roleta,
            x="Dia",
            y="RN_Total",
            markers=True,
            title="Evolução diária do RN total",
        )
        fig_rn.update_traces(line_color="#60a5fa", marker_color="#60a5fa")
        st.plotly_chart(estilizar_fig(fig_rn), use_container_width=True)

    st.markdown("<div class='section-title'>Leitura executiva da roleta</div>", unsafe_allow_html=True)

    if conv_roleta < 0.03:
        st.markdown(
            f"""
            <div class="alert-bad">
                <b>Alerta:</b> baixa conversão da roleta no mês.<br>
                Roleta total: {fmt_int(roleta_total_mes)} |
                Vendas: {fmt_int(vendas_mes)} |
                Conversão: {fmt_pct(conv_roleta)}
            </div>
            """,
            unsafe_allow_html=True,
        )
    elif conv_roleta >= 0.03 and conv_roleta < 0.06:
        st.markdown(
            """
            <div class="alert-warn">
                <b>Status:</b> conversão da roleta em nível intermediário.<br>
                Há espaço para melhorar abordagem, qualificação e fechamento.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"""
            <div class="alert-good">
                <b>Destaque:</b> a roleta está sendo bem aproveitada.<br>
                Conversão atual: {fmt_pct(conv_roleta)} |
                VGV por roleta: {fmt_money(vgv_por_roleta, 0)}
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.dataframe(df_roleta, use_container_width=True, height=320)


# =========================
# HEADER
# =========================
st.markdown(
    """
    <div class="brand-box">
        <div class="brand-left">
            <div class="brand-logo">T3 IMOVEIS</div>
            <div class="brand-divider"></div>
            <div>
                <div class="brand-title">Gestão Roque</div>
                <div class="brand-subtitle">Consolidado do funil, ranking de operações, gaps de execução e leitura de performance em tempo real.</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Controles")
    if st.button("Atualizar dados", use_container_width=True):
        st.cache_data.clear()
        st.rerun()


# =========================
# CARGA PRINCIPAL
# =========================
df_resumo, df_funil, erros = carregar_dados_dashboard()

if erros:
    with st.expander("Avisos de leitura", expanded=False):
        for e in erros:
            st.warning(e)

if df_funil.empty:
    st.error("Não foi possível montar o dashboard a partir da planilha.")
    st.stop()

df_geral = df_funil[df_funil["aba"] == "Funil de Vendas Geral"].copy()
df_ops = df_funil[~df_funil["aba"].isin(ABAS_EXCLUIDAS_RANKING)].copy()

with st.sidebar:
    abas_ops = sorted(df_ops["aba"].dropna().unique().tolist())
    operacao_filtro = st.selectbox("Operação", ["Todas"] + abas_ops, index=0)

df_ops_filtrado = df_ops[df_ops["aba"] == operacao_filtro].copy() if operacao_filtro != "Todas" else df_ops.copy()


# =========================
# KPIS TOPO
# =========================
mes_nome = df_funil["mes_nome"].dropna().iloc[0] if not df_funil["mes_nome"].dropna().empty else "-"
dias_mes = numero(df_funil["dias_mes"].dropna().iloc[0]) if not df_funil["dias_mes"].dropna().empty else 0
dia_atual = numero(df_funil["dia_atual"].dropna().iloc[0]) if not df_funil["dia_atual"].dropna().empty else 0

leads_esperado = valor_etapa(df_geral, "Leads", "expected_now")
leads_atual = valor_etapa(df_geral, "Leads", "actual_now")
vendas_esperado = valor_etapa(df_geral, "Venda", "expected_now")
vendas_atual = valor_etapa(df_geral, "Venda", "actual_now")
gap_vendas = valor_etapa(df_geral, "Venda", "gap")
ating_vendas = safe_div(vendas_atual, vendas_esperado)
ritmo_mes = safe_div(dia_atual, dias_mes)

k1, k2, k3, k4, k5 = st.columns(5)
with k1:
    card_kpi("Mês / Dia", f"{mes_nome} • {int(dia_atual)}/{int(dias_mes)}", f"Ritmo do mês: {fmt_pct(ritmo_mes)}")
with k2:
    card_kpi("Leads esperados hoje", fmt_int(leads_esperado), f"Atual: {fmt_int(leads_atual)}")
with k3:
    card_kpi("Vendas esperadas hoje", fmt_num(vendas_esperado, 2), f"Atual: {fmt_num(vendas_atual, 2)}")
with k4:
    card_kpi("Gap de vendas", fmt_num(gap_vendas, 2), "Resultado acumulado vs esperado")
with k5:
    card_kpi("Atingimento de vendas", fmt_pct(ating_vendas), "Atual / esperado")


# =========================
# TABS
# =========================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["Visão Executiva", "Ranking de Operações", "Funil Consolidado", "Diagnóstico", "Performance Financeira", "Roleta"]
)


with tab1:
    st.markdown("<div class='section-title'>Comparativo executivo do consolidado</div>", unsafe_allow_html=True)
    base_exec = df_geral[["etapa", "expected_now", "actual_now", "gap"]].copy()

    fig_exec = go.Figure()
    fig_exec.add_bar(name="Esperado até hoje", x=base_exec["etapa"], y=base_exec["expected_now"], marker_color="#143554")
    fig_exec.add_bar(name="Realizado", x=base_exec["etapa"], y=base_exec["actual_now"], marker_color="#4dd099")
    fig_exec.update_layout(
        barmode="group",
        title="Esperado x Realizado por etapa",
        xaxis_title="Etapa",
        yaxis_title="Volume",
        legend_title=""
    )
    st.plotly_chart(estilizar_fig(fig_exec), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("<div class='section-title'>Gaps por etapa</div>", unsafe_allow_html=True)
        fig_gap = px.bar(
            base_exec,
            x="etapa",
            y="gap",
            text_auto=".2f",
            title="Gap acumulado por etapa",
            color_discrete_sequence=["#4dd099"],
        )
        st.plotly_chart(estilizar_fig(fig_gap), use_container_width=True)

    with c2:
        st.markdown("<div class='section-title'>Conversão atual do funil</div>", unsafe_allow_html=True)
        fig_funil = go.Figure(
            go.Funnel(
                y=df_geral["etapa"],
                x=df_geral["actual_now"],
                textinfo="value+percent initial",
                marker={"color": ["#143554", "#1f4f73", "#2b6a88", "#39a88a", "#4dd099"]},
            )
        )
        fig_funil.update_layout(title="Funil real do consolidado")
        st.plotly_chart(estilizar_fig(fig_funil), use_container_width=True)


with tab2:
    st.markdown("<div class='section-title'>Ranking por operação</div>", unsafe_allow_html=True)
    rank_vendas = df_ops[df_ops["etapa"] == "Venda"].copy()
    rank_vendas["atingimento_pct"] = rank_vendas["atingimento"] * 100
    rank_vendas = rank_vendas[["aba", "expected_now", "actual_now", "gap", "atingimento_pct"]].sort_values("actual_now", ascending=False)

    st.dataframe(
        rank_vendas.rename(
            columns={
                "aba": "Operação",
                "expected_now": "Esperado",
                "actual_now": "Realizado",
                "gap": "Gap",
                "atingimento_pct": "Atingimento %",
            }
        ),
        use_container_width=True,
        height=320,
    )

    c1, c2 = st.columns(2)
    with c1:
        fig_rank = px.bar(
            rank_vendas.sort_values("actual_now", ascending=True),
            x="actual_now",
            y="aba",
            orientation="h",
            text_auto=".2f",
            title="Ranking de vendas realizadas",
            color_discrete_sequence=["#4dd099"],
        )
        st.plotly_chart(estilizar_fig(fig_rank), use_container_width=True)

    with c2:
        fig_ating = px.bar(
            rank_vendas.sort_values("atingimento_pct", ascending=True),
            x="atingimento_pct",
            y="aba",
            orientation="h",
            text_auto=".1f",
            title="Atingimento de vendas (%)",
            color_discrete_sequence=["#143554"],
        )
        st.plotly_chart(estilizar_fig(fig_ating), use_container_width=True)

    st.markdown("<div class='section-title'>Detalhe da operação selecionada</div>", unsafe_allow_html=True)

    if operacao_filtro == "Todas":
        st.info("Selecione uma operação na barra lateral para ver o detalhe individual.")
    else:
        base_op = df_ops_filtrado.copy()
        if base_op.empty:
            st.warning("Sem dados para a operação selecionada.")
        else:
            top1, top2, top3, top4 = st.columns(4)

            vendas_esp = numero(base_op.loc[base_op["etapa"] == "Venda", "expected_now"].sum())
            vendas_real = numero(base_op.loc[base_op["etapa"] == "Venda", "actual_now"].sum())
            leads_esp = numero(base_op.loc[base_op["etapa"] == "Leads", "expected_now"].sum())
            leads_real = numero(base_op.loc[base_op["etapa"] == "Leads", "actual_now"].sum())

            with top1:
                card_kpi("Operação", operacao_filtro, "Equipe selecionada")
            with top2:
                card_kpi("Leads", fmt_num(leads_real, 2), f"Esperado: {fmt_num(leads_esp, 2)}")
            with top3:
                card_kpi("Vendas", fmt_num(vendas_real, 2), f"Esperado: {fmt_num(vendas_esp, 2)}")
            with top4:
                card_kpi("Atingimento Vendas", fmt_pct(safe_div(vendas_real, vendas_esp)), "Realizado / esperado")

            d1, d2 = st.columns(2)
            with d1:
                fig_op = go.Figure()
                fig_op.add_bar(name="Esperado", x=base_op["etapa"], y=base_op["expected_now"], marker_color="#143554")
                fig_op.add_bar(name="Realizado", x=base_op["etapa"], y=base_op["actual_now"], marker_color="#4dd099")
                fig_op.update_layout(
                    barmode="group",
                    title=f"Esperado x Realizado • {operacao_filtro}",
                    xaxis_title="Etapa",
                    yaxis_title="Volume",
                )
                st.plotly_chart(estilizar_fig(fig_op), use_container_width=True)

            with d2:
                fig_funil_op = go.Figure(
                    go.Funnel(
                        y=base_op["etapa"],
                        x=base_op["actual_now"],
                        textinfo="value+percent initial",
                        marker={"color": ["#143554", "#1f4f73", "#2b6a88", "#39a88a", "#4dd099"]},
                    )
                )
                fig_funil_op.update_layout(title=f"Funil atual • {operacao_filtro}")
                st.plotly_chart(estilizar_fig(fig_funil_op), use_container_width=True)

            base_gap = base_op[["etapa", "gap"]].copy()
            fig_gap_op = px.bar(
                base_gap,
                x="etapa",
                y="gap",
                text_auto=".2f",
                title=f"Gap por etapa • {operacao_filtro}",
                color_discrete_sequence=["#4dd099"],
            )
            st.plotly_chart(estilizar_fig(fig_gap_op), use_container_width=True)

            st.dataframe(
                base_op[["etapa", "target_total", "expected_now", "actual_now", "gap", "atingimento"]]
                .assign(atingimento=lambda x: x["atingimento"] * 100)
                .rename(
                    columns={
                        "etapa": "Etapa",
                        "target_total": "Meta Mês",
                        "expected_now": "Esperado Hoje",
                        "actual_now": "Realizado",
                        "gap": "Gap",
                        "atingimento": "Atingimento %",
                    }
                ),
                use_container_width=True,
                height=260,
            )


with tab3:
    st.markdown("<div class='section-title'>Meta total, esperado e realizado</div>", unsafe_allow_html=True)
    fig_full = go.Figure()
    fig_full.add_bar(name="Meta total do mês", x=df_geral["etapa"], y=df_geral["target_total"], marker_color="#dce5ec")
    fig_full.add_bar(name="Onde deveria estar hoje", x=df_geral["etapa"], y=df_geral["expected_now"], marker_color="#143554")
    fig_full.add_bar(name="Onde estou", x=df_geral["etapa"], y=df_geral["actual_now"], marker_color="#4dd099")
    fig_full.update_layout(
        barmode="group",
        title="Leitura completa do funil",
        xaxis_title="Etapa",
        yaxis_title="Volume",
        legend_title=""
    )
    st.plotly_chart(estilizar_fig(fig_full), use_container_width=True)

    st.markdown("<div class='section-title'>Tabela consolidada do funil</div>", unsafe_allow_html=True)
    tabela_funil = df_geral[["etapa", "target_total", "expected_now", "actual_now", "gap", "atingimento"]].copy()
    tabela_funil["atingimento"] = tabela_funil["atingimento"] * 100
    st.dataframe(
        tabela_funil.rename(
            columns={
                "etapa": "Etapa",
                "target_total": "Meta do mês",
                "expected_now": "Esperado hoje",
                "actual_now": "Realizado",
                "gap": "Gap",
                "atingimento": "Atingimento %",
            }
        ),
        use_container_width=True,
        height=300,
    )


with tab4:
    st.markdown("<div class='section-title'>Leitura executiva dos gargalos</div>", unsafe_allow_html=True)
    base_diag = df_geral.copy()
    base_diag["atingimento_pct"] = base_diag["atingimento"] * 100
    piores = base_diag.sort_values("atingimento_pct").head(2)
    melhores = base_diag.sort_values("atingimento_pct", ascending=False).head(2)

    for _, row in piores.iterrows():
        st.markdown(
            f"""<div class="alert-bad"><b>Gargalo:</b> {row['etapa']} <br>Esperado até hoje: {fmt_num(row['expected_now'], 2)}<br>Realizado: {fmt_num(row['actual_now'], 2)}<br>Atingimento: {fmt_pct(row['atingimento'])}<br>Gap: {fmt_num(row['gap'], 2)}</div>""",
            unsafe_allow_html=True,
        )

    for _, row in melhores.iterrows():
        st.markdown(
            f"""<div class="alert-good"><b>Destaque:</b> {row['etapa']} <br>Esperado até hoje: {fmt_num(row['expected_now'], 2)}<br>Realizado: {fmt_num(row['actual_now'], 2)}<br>Atingimento: {fmt_pct(row['atingimento'])}<br>Gap: {fmt_num(row['gap'], 2)}</div>""",
            unsafe_allow_html=True,
        )

    vendas_ops = df_ops[df_ops["etapa"] == "Venda"].copy().sort_values("atingimento")
    if not vendas_ops.empty:
        pior_op = vendas_ops.iloc[0]
        melhor_op = vendas_ops.iloc[-1]

        st.markdown(
            f"""<div class="alert-warn"><b>Operação mais pressionada em vendas:</b> {pior_op['aba']}<br>Esperado: {fmt_num(pior_op['expected_now'], 2)} | Realizado: {fmt_num(pior_op['actual_now'], 2)} | Atingimento: {fmt_pct(pior_op['atingimento'])}</div>""",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"""<div class="alert-good"><b>Operação destaque em vendas:</b> {melhor_op['aba']}<br>Esperado: {fmt_num(melhor_op['expected_now'], 2)} | Realizado: {fmt_num(melhor_op['actual_now'], 2)} | Atingimento: {fmt_pct(melhor_op['atingimento'])}</div>""",
            unsafe_allow_html=True,
        )


with tab5:
    render_financeiro()


with tab6:
    render_roleta()