import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Dashboard de Atendimentos", layout="wide")

# Upload do arquivo
uploaded_file = st.file_uploader("📂 Envie seu arquivo Excel (.xls ou .xlsx)", type=["xls", "xlsx"])

if uploaded_file is not None:
    # Leitura robusta (tenta openpyxl -> depois xlrd)
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception:
        try:
            df = pd.read_excel(uploaded_file, engine="xlrd")
        except Exception as e:
            st.error(f"Erro ao abrir arquivo: {e}")
            st.stop()

    # 2) Conversão de colunas de data
    for col in ["created_at", "closed_at", "updated_at", "last_receive", "last_send", "attended_at"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # 3) Apenas fechados
    df_closed = df[df.get("closed", 0) == 1].copy()

    # ============================
    # 🔹 Filtro de datas
    if "created_at" in df_closed.columns:
        min_date = df_closed["created_at"].min().date()
        max_date = df_closed["created_at"].max().date()

        st.sidebar.header("📅 Filtros")
        start_date, end_date = st.sidebar.date_input(
            "Selecione o período (Criação):",
            value=[min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )

        # Aplica filtro
        if start_date and end_date:
            mask = (df_closed["created_at"].dt.date >= start_date) & (df_closed["created_at"].dt.date <= end_date)
            df_closed = df_closed.loc[mask]

    # ============================

    # 4) Cálculo de tempos (em minutos)
    m_total = df_closed["created_at"].notna() & df_closed["closed_at"].notna()
    m_espera = df_closed["created_at"].notna() & df_closed["attended_at"].notna()
    m_agente = df_closed["attended_at"].notna() & df_closed["closed_at"].notna()

    df_closed.loc[m_total, "tempo_ciclo_total_min"] = (
        (df_closed.loc[m_total, "closed_at"] - df_closed.loc[m_total, "created_at"]).dt.total_seconds() // 60
    )

    df_closed.loc[m_espera, "tempo_espera_cliente_min"] = (
        (df_closed.loc[m_espera, "attended_at"] - df_closed.loc[m_espera, "created_at"]).dt.total_seconds() // 60
    )

    df_closed.loc[m_agente, "tempo_atendimento_agente_min"] = (
        (df_closed.loc[m_agente, "closed_at"] - df_closed.loc[m_agente, "attended_at"]).dt.total_seconds() // 60
    )

    for c in ["tempo_ciclo_total_min", "tempo_espera_cliente_min", "tempo_atendimento_agente_min"]:
        if c in df_closed.columns:
            df_closed[c] = df_closed[c].clip(lower=0)

    # Formatador de tempo em horas e minutos
    def format_time_minutes(minutes):
        if pd.isna(minutes):
            return "0m"
        minutes = int(minutes)
        horas = minutes // 60
        mins = minutes % 60
        return f"{horas}h {mins}m" if horas > 0 else f"{mins}m"

    # 5) KPIs
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📌 Atendimentos Fechados", len(df_closed))
    col2.metric("🏢 Empresas", df_closed["company_name"].nunique() if "company_name" in df_closed else 0)
    col3.metric("📡 Canais", df_closed["channel_typename"].nunique() if "channel_typename" in df_closed else 0)

    media_atend_agente_min = df_closed["tempo_atendimento_agente_min"].mean(skipna=True)
    col4.metric(
        "⏱️ Tempo Médio de Atendimento (Agente → Fechamento)",
        f"{format_time_minutes(media_atend_agente_min)} ({media_atend_agente_min/60:.1f}h)"
    )

    col5, col6 = st.columns(2)

    media_espera_min = df_closed["tempo_espera_cliente_min"].mean(skipna=True)
    col5.metric(
        "⌛ Tempo Médio de Espera(Criação → 1º Atendimento)",
        f"{format_time_minutes(media_espera_min)} ({media_espera_min/60:.1f}h)"
    )

    media_ciclo_total_min = df_closed["tempo_ciclo_total_min"].mean(skipna=True)
    col6.metric(
        "🔁 Tempo Médio de Ciclo Total (Criação → Fechamento)",
        f"{format_time_minutes(media_ciclo_total_min)} ({media_ciclo_total_min/60:.1f}h)"
    )

    # 6) Gráficos
    colg1, colg2 = st.columns(2)

    if "channel_name" in df_closed.columns:
        ch_count = df_closed.groupby("channel_name").size().reset_index(name="total")
        fig_ch = px.pie(ch_count, names="channel_name", values="total", title="Atendimentos por Canal")
        colg1.plotly_chart(fig_ch, use_container_width=True)

    if "tabulation_comment" in df_closed.columns:
        tab_count = df_closed.groupby("tabulation_comment").size().reset_index(name="total")
        fig_tab = px.bar(tab_count, x="tabulation_comment", y="total", color="tabulation_comment",
                         title="Distribuição por Tabulação")
        colg2.plotly_chart(fig_tab, use_container_width=True)

    if "customer_name" in df_closed.columns:
        cust_count = (df_closed.groupby("customer_name").size().reset_index(name="total")
                      .sort_values(by="total", ascending=False).head(10))
        fig_cust = px.bar(cust_count, x="customer_name", y="total", color="customer_name", title="Top 10 Clientes")
        st.plotly_chart(fig_cust, use_container_width=True)

    if "agent_login" in df_closed.columns:
        # Top 10 agentes por chamados fechados df_closed
        agent_count = (df_closed.groupby("agent_login").size()
                       .reset_index(name="total")
                       .sort_values(by="total", ascending=False)
                       .head(10))

        fig_agent = px.bar(agent_count,
                           x="agent_login",
                           y="total",
                           color="agent_login",
                           text="total",
                           title="🏆 Top 10 Agentes por Chamados Fechados")
        fig_agent.update_traces(textposition="outside")
        st.plotly_chart(fig_agent, use_container_width=True)

    # 7) Tabela detalhada
    st.subheader("📋 Atendimentos Fechados - Detalhes")
    st.dataframe(df_closed, use_container_width=True)
else:
    st.info("⬆️ Envie um arquivo Excel para visualizar o dashboard.")
