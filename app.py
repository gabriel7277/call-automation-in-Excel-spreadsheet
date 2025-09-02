import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Dashboard de Atendimentos", layout="wide")

# ==============================
# 1) Upload com cache
# ==============================
@st.cache_data
def load_excel(file):
    try:
        return pd.read_excel(file, engine="openpyxl")
    except:
        return pd.read_excel(file, engine="xlrd")

uploaded_file = st.file_uploader("📂 Envie seu arquivo Excel (.xls ou .xlsx)", type=["xls", "xlsx"])

if uploaded_file is not None:
    df = load_excel(uploaded_file)

    # ==============================
    # 2) Selecionar colunas importantes e converter datas
    # ==============================
    cols_keep = [
        "created_at", "closed_at", "attended_at", "updated_at",
        "last_receive", "last_send", "company_name",
        "channel_typename", "channel_name", "tabulation_comment",
        "customer_name", "agent_login", "closed"
    ]
    df = df[[c for c in cols_keep if c in df.columns]].copy()

    date_cols = ["created_at", "closed_at", "updated_at", "last_receive", "last_send", "attended_at"]
    for col in set(date_cols).intersection(df.columns):
        df[col] = pd.to_datetime(df[col], errors="coerce", utc=True)

    # ==============================
    # 3) Apenas fechados
    # ==============================
    df_closed = df[df.get("closed", 0) == 1].copy()

    # ==============================
    # 4) Filtro de datas
    # ==============================
    if "created_at" in df.columns:
        min_date = df["created_at"].min().date()
        max_date = df["created_at"].max().date()

        st.sidebar.header("📅 Filtros")
        date_range = st.sidebar.date_input(
            "Selecione o período (Criação):",
            value=[min_date, max_date],
            min_value=min_date,
            max_value=max_date,
        )

        if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
            start_date, end_date = date_range
            df = df[(df["created_at"].dt.date >= start_date) & (df["created_at"].dt.date <= end_date)]
            df_closed = df_closed[(df_closed["created_at"].dt.date >= start_date) & (df_closed["created_at"].dt.date <= end_date)]

    # ==============================
    # 5) Filtro avançado por canal
    # ==============================
    if "channel_name" in df.columns:
        canais_disponiveis = df["channel_name"].unique().tolist()
        canais_selecionados = st.sidebar.multiselect(
            "Selecione os canais",
            options=canais_disponiveis,
            default=canais_disponiveis
        )

        df = df[df["channel_name"].isin(canais_selecionados)]
        df_closed = df_closed[df_closed["channel_name"].isin(canais_selecionados)]

    # ==============================
    # 6) Cálculo de tempos (em minutos) com cache
    # ==============================
    @st.cache_data
    def calcula_tempos(df_input):
        df_input = df_input.copy()
        m_total = df_input["created_at"].notna() & df_input["closed_at"].notna()
        m_espera = df_input["created_at"].notna() & df_input["attended_at"].notna()
        m_agente = df_input["attended_at"].notna() & df_input["closed_at"].notna()

        df_input.loc[m_total, "tempo_ciclo_total_min"] = ((df_input.loc[m_total, "closed_at"] - df_input.loc[m_total, "created_at"]).dt.total_seconds() // 60)
        df_input.loc[m_espera, "tempo_espera_cliente_min"] = ((df_input.loc[m_espera, "attended_at"] - df_input.loc[m_espera, "created_at"]).dt.total_seconds() // 60)
        df_input.loc[m_agente, "tempo_atendimento_agente_min"] = ((df_input.loc[m_agente, "closed_at"] - df_input.loc[m_agente, "attended_at"]).dt.total_seconds() // 60)

        for c in ["tempo_ciclo_total_min", "tempo_espera_cliente_min", "tempo_atendimento_agente_min"]:
            if c in df_input.columns:
                df_input[c] = df_input[c].clip(lower=0)

        return df_input

    df = calcula_tempos(df)
    df_closed = calcula_tempos(df_closed)

    # ==============================
    # 7) Função para formatar tempo
    # ==============================
    def format_time(minutes):
        if pd.isna(minutes) or minutes == 0:
            return "0m"
        hours = int(minutes // 60)
        mins = int(minutes % 60)
        return f"{hours}h {mins}m" if hours > 0 else f"{mins}m"

    # ==============================
    # 8) KPIs
    # ==============================
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📌 Atendimentos Fechados", len(df_closed))
    col2.metric("🏢 Empresas", df["company_name"].nunique() if "company_name" in df else 0)
    col3.metric("📡 Canais", df["channel_typename"].nunique() if "channel_typename" in df else 0)

    media_atend_agente = df_closed["tempo_atendimento_agente_min"].mean(skipna=True)
    col4.metric("⏱️ Tempo Médio Atendimento (Agente → Fechamento)",
                f"{format_time(media_atend_agente)} ({media_atend_agente/60:.1f}h)")

    col5, col6 = st.columns(2)
    media_espera = df_closed["tempo_espera_cliente_min"].mean(skipna=True)
    media_ciclo = df_closed["tempo_ciclo_total_min"].mean(skipna=True)
    col5.metric("⌛ Tempo Médio Espera (Criação → 1º Atendimento)", f"{format_time(media_espera)} ({media_espera/60:.1f}h)")
    col6.metric("🔁 Tempo Médio Ciclo Total (Criação → Fechamento)", f"{format_time(media_ciclo)} ({media_ciclo/60:.1f}h)")

    # ==============================
    # 9) Gráficos
    # ==============================
    colg1, colg2 = st.columns(2)

    # --- Treemap de canais: total x fechados ---
    if "channel_name" in df.columns and len(df) > 0:
        total_por_canal = df.groupby("channel_name").size().reset_index(name="total_atendimentos")
        fechados_por_canal = df_closed.groupby("channel_name").size().reset_index(name="total_fechados")
        ch_count = pd.merge(total_por_canal, fechados_por_canal, on="channel_name", how="left").fillna(0)
        ch_count["total_fechados"] = ch_count["total_fechados"].astype(int)

        fig_ch = px.treemap(
            ch_count,
            path=["channel_name"],
            values="total_atendimentos",
            color="total_fechados",
            color_continuous_scale="Viridis",
            title="📊 Atendimentos por Canal (Total vs Fechados)"
        )
        fig_ch.update_traces(
            texttemplate="%{label}\nTotal: %{value}\nFechados: %{customdata[0]}",
            textinfo="text",
            textfont_size=22,
            customdata=ch_count[['total_fechados']].values
        )
        colg1.plotly_chart(fig_ch, use_container_width=True)

    # --- Gráfico de tabulação ---
    if "tabulation_comment" in df.columns:
        tab_count = df_closed.groupby("tabulation_comment").size().reset_index(name="total")
        if len(tab_count) > 8:
            tab_count = tab_count.sort_values(by="total", ascending=True)
            fig_tab = px.bar(tab_count, y="tabulation_comment", x="total", color="total",
                             orientation="h", color_continuous_scale="Viridis", title="Distribuição por Tabulação")
            fig_tab.update_layout(xaxis_title="Total de Chamados", yaxis_title="Tabulação", yaxis=dict(tickmode="linear"), margin=dict(l=150))
        else:
            tab_count = tab_count.sort_values(by="total", ascending=False)
            fig_tab = px.bar(tab_count, x="tabulation_comment", y="total", color="total",
                             color_continuous_scale="Viridis", title="Distribuição por Tabulação")
            fig_tab.update_layout(xaxis_tickangle=-45, margin=dict(b=150))
        fig_tab.update_traces(text=tab_count["total"], textposition="outside")
        colg2.plotly_chart(fig_tab, use_container_width=True)

    # --- Top 10 clientes ---
    if "customer_name" in df_closed.columns:
        cust_count = df_closed.groupby("customer_name").size().reset_index(name="total").sort_values(by="total", ascending=False).head(10)
        fig_cust = px.bar(cust_count, x="customer_name", y="total", color="customer_name", text="total", title="Top 10 Clientes")
        fig_cust.update_traces(textposition="outside")
        st.plotly_chart(fig_cust, use_container_width=True)

    # --- Top 10 agentes ---
    if "agent_login" in df_closed.columns:
        agent_count = df_closed.groupby("agent_login").size().reset_index(name="total").sort_values(by="total", ascending=False).head(10)
        fig_agent = px.bar(agent_count, x="agent_login", y="total", color="agent_login", text="total", title="🏆 Top Agentes por Chamados Fechados")
        fig_agent.update_traces(textposition="outside")
        st.plotly_chart(fig_agent, use_container_width=True)

    # ==============================
    # 10) Tabela detalhada
    # ==============================
    st.subheader("📋 Atendimentos Fechados - Detalhes")
    st.dataframe(df_closed.head(500), use_container_width=True)
    st.caption(f"Mostrando 500 de {len(df_closed)} registros")

else:
    st.info("⬆️ Envie um arquivo Excel para visualizar o dashboard.")
