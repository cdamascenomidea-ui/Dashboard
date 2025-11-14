import streamlit as st  # pyright: ignore[reportMissingImports]
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np

# Configuração da página
st.set_page_config(
    page_title="Dashboard Braspress - Monitoramento de Entregas",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        text-align: center;
    }
    .status-danger {
        color: #d62728;
        font-weight: bold;
    }
    .status-warning {
        color: #ff7f0e;
        font-weight: bold;
    }
    .status-success {
        color: #2ca02c;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# Função para carregar dados


@st.cache_data
def load_data():
    """Carrega e processa os dados do Excel"""
    df = pd.read_excel(
        'Monitoramento Tarefas  Braspress  13112025 Carlos.xlsx', sheet_name='Braspress')

    # Converter colunas de data
    date_columns = ['Dt.OV', 'Dt.Conhecimento', 'Dt.Trânsito',
                    'Data de previsão de entrega', 'Previsão de Entrega', 'Data da Entrega ']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(
                df[col], origin='1899-12-30', unit='D', errors='coerce')

    # Limpar e converter dias de atraso
    if 'Dias de Atraso - Cliente' in df.columns:
        df['Dias de Atraso - Cliente'] = pd.to_numeric(
            df['Dias de Atraso - Cliente'], errors='coerce')
    
    # Garantir que colunas essenciais existem
    colunas_essenciais = ['Estado', 'Status Final', 'Ocorrências', 'Nome do Cliente']
    for col in colunas_essenciais:
        if col not in df.columns:
            df[col] = ''

    return df


# Carregar dados
try:
    df = load_data()

    # Header
    st.markdown('<div class="main-header">📦 Dashboard Braspress - Monitoramento de Entregas</div>',
                unsafe_allow_html=True)
    st.markdown(
        f"**Última atualização:** {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    st.markdown("---")

    # Sidebar - Filars
    st.sidebar.header("🔍 Filars")

    # Filar por Estado
    estados_disponiveis = ['Todos'] + \
        sorted(df['Estado'].dropna().unique().tolist())
    estado_selecionado = st.sidebar.selectbox("Estado", estados_disponiveis)

    # Filar por Status Final
    status_disponiveis = ['Todos'] + \
        sorted(df['Status Final'].dropna().unique().tolist())
    status_selecionado = st.sidebar.selectbox(
        "Status Final", status_disponiveis)

    # Filar por período de atraso
    max_atraso_data = df['Dias de Atraso - Cliente'].max()
    max_atraso_valor = int(max_atraso_data) if pd.notna(max_atraso_data) else 50
    
    atraso_min, atraso_max = st.sidebar.slider(
        "Dias de Atraso",
        min_value=0,
        max_value=max_atraso_valor,
        value=(0, min(50, max_atraso_valor))
    )

    # Aplicar filars
    df_filtrado = df.copy()
    if estado_selecionado != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Estado'] == estado_selecionado]
    if status_selecionado != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Status Final']
                                  == status_selecionado]
    
    # Filtrar por atraso com segurança para NaN
    df_filtrado = df_filtrado[
        (df_filtrado['Dias de Atraso - Cliente'].fillna(0) >= atraso_min) &
        (df_filtrado['Dias de Atraso - Cliente'].fillna(0) <= atraso_max)
    ]

    # KPIs Principais
    st.header("📊 Indicadores Principais")
    col1, col2, col3, col4, col5 = st.columns(5)

    total_entregas = len(df_filtrado)
    fora_prazo = len(
        df_filtrado[df_filtrado['Status Final'].str.contains('Fora do Prazo', na=False)])
    dentro_prazo = len(df_filtrado[df_filtrado['Status Final'].str.contains(
        'Dentro do Prazo', na=False)])
    taxa_pontualidade = (dentro_prazo / total_entregas *
                         100) if total_entregas > 0 else 0
    atraso_medio = df_filtrado['Dias de Atraso - Cliente'].mean()

    with col1:
        st.metric("Total de Entregas", f"{total_entregas:,}")
    with col2:
        st.metric("Dentro do Prazo",
                  f"{dentro_prazo}", delta=f"{taxa_pontualidade:.1f}%")
    with col3:
        st.metric("Fora do Prazo", f"{fora_prazo}",
                  delta=f"-{(fora_prazo/total_entregas*100):.1f}%" if total_entregas > 0 else "0%", delta_color="inverse")
    with col4:
        st.metric("Atraso Médio", f"{atraso_medio:.1f} dias" if not pd.isna(
            atraso_medio) else "N/A")
    with col5:
        max_atraso = df_filtrado['Dias de Atraso - Cliente'].max()
        st.metric("Atraso Máximo", f"{int(max_atraso)} dias" if pd.notna(max_atraso) else "N/A")

    st.markdown("---")

    # Gráficos - Linha 1
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📈 Distribuição por Status Final")
        status_counts = df_filtrado['Status Final'].value_counts()
        if len(status_counts) > 0:
            fig_status = px.bar(
                x=status_counts.index,
                y=status_counts.values,
                labels={'x': 'Status', 'y': 'Quantidade'},
                color=status_counts.values,
                color_continuous_scale='RdYlGn_r'
            )
            fig_status.update_layout(showlegend=False, height=400)
            fig_status.update_xaxes(tickangle=-45)
            st.plotly_chart(fig_status, use_container_width=True)
        else:
            st.info("Sem dados disponíveis para este filar")

    with col2:
        st.subheader("🗺️ Entregas por Estado (Top 10)")
        estado_counts = df_filtrado['Estado'].value_counts().head(10)
        if len(estado_counts) > 0:
            fig_estados = px.pie(
                values=estado_counts.values,
                names=estado_counts.index,
                hole=0.4
            )
            fig_estados.update_traces(
                textposition='inside', textinfo='percent+label')
            fig_estados.update_layout(height=400)
            st.plotly_chart(fig_estados, use_container_width=True)
        else:
            st.info("Sem dados disponíveis para este filar")

    # Gráficos - Linha 2
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("⚠️ Top 10 Ocorrências")
        ocorrencias_counts = df_filtrado['Ocorrências'].value_counts().head(10)
        if len(ocorrencias_counts) > 0:
            fig_ocorrencias = px.bar(
                x=ocorrencias_counts.values,
                y=ocorrencias_counts.index,
                orientation='h',
                labels={'x': 'Quantidade', 'y': 'Ocorrência'},
                color=ocorrencias_counts.values,
                color_continuous_scale='Reds'
            )
            fig_ocorrencias.update_layout(showlegend=False, height=500, yaxis={
                                          'categoryorder': 'total ascending'})
            st.plotly_chart(fig_ocorrencias, use_container_width=True)
        else:
            st.info("Sem dados de ocorrências disponíveis para o filar selecionado")

    with col2:
        st.subheader("📊 Distribuição de Atrasos")
        atrasos_validos = df_filtrado[df_filtrado['Dias de Atraso - Cliente'].notna()]
        if len(atrasos_validos) > 0:
            fig_atrasos = px.histogram(
                atrasos_validos,
                x='Dias de Atraso - Cliente',
                nbins=20,
                labels={'Dias de Atraso - Cliente': 'Dias de Atraso'},
                color_discrete_sequence=['#ff7f0e']
            )
            fig_atrasos.update_layout(height=500, showlegend=False)
            st.plotly_chart(fig_atrasos, use_container_width=True)
        else:
            st.info("Sem dados de atraso disponíveis para o filar selecionado")

    # Gráfico de Tendência Temporal
    st.subheader("📅 Tendência de Entregas ao Longo do Tempo")
    df_com_data = df_filtrado[df_filtrado['Dt.OV'].notna()].copy()
    if len(df_com_data) > 0:
        df_com_data['Mes'] = df_com_data['Dt.OV'].dt.to_period('M').astype(str)
        entregas_por_mes = df_com_data.groupby(
            'Mes').size().reset_index(name='Quantidade')

        fig_tendencia = px.line(
            entregas_por_mes,
            x='Mes',
            y='Quantidade',
            markers=True,
            labels={'Mes': 'Mês', 'Quantidade': 'Número de Entregas'}
        )
        fig_tendencia.update_traces(line_color='#1f77b4', line_width=3)
        fig_tendencia.update_layout(height=400)
        st.plotly_chart(fig_tendencia, use_container_width=True)
    else:
        st.info("Sem dados de data disponíveis para o filar selecionado")

    # Análise de Performance por Estado
    st.subheader("🎯 Performance por Estado")
    performance_estado = df_filtrado.groupby('Estado').agg({
        'Dias de Atraso - Cliente': 'mean',
        'Nome do Cliente': 'count'
    }).reset_index()
    performance_estado.columns = [
        'Estado', 'Atraso Médio (dias)', 'Total de Entregas']
    performance_estado = performance_estado.sort_values(
        'Total de Entregas', ascending=False).head(10)

    fig_performance = go.Figure()
    fig_performance.add_trace(go.Bar(
        x=performance_estado['Estado'],
        y=performance_estado['Total de Entregas'],
        name='Total de Entregas',
        marker_color='lightblue',
        yaxis='y'
    ))
    fig_performance.add_trace(go.Scatter(
        x=performance_estado['Estado'],
        y=performance_estado['Atraso Médio (dias)'],
        name='Atraso Médio',
        marker_color='red',
        yaxis='y2',
        mode='lines+markers',
        line=dict(width=3)
    ))
    fig_performance.update_layout(
        height=400,
        yaxis=dict(title='Total de Entregas'),
        yaxis2=dict(title='Atraso Médio (dias)', overlaying='y', side='right'),
        hovermode='x unified'
    )
    st.plotly_chart(fig_performance, use_container_width=True)

    # Tabela de Detalhes
    st.subheader("📋 Detalhamento das Entregas")

    # Opção de exportar
    col1, col2 = st.columns([3, 1])
    with col2:
        st.download_button(
            label="📥 Exportar para CSV",
            data=df_filtrado.to_csv(index=False).encode('utf-8'),
            file_name=f'entregas_braspress_{datetime.now().strftime("%Y%m%d")}.csv',
            mime='text/csv'
        )

    # Selecionar colunas para exibir
    colunas_exibir = ['Nome do Cliente', 'Estado', 'Cidade', 'Status Final',
                      'Dias de Atraso - Cliente', 'Ocorrências', 'Status Atual', 'Observação']
    colunas_disponiveis = [
        col for col in colunas_exibir if col in df_filtrado.columns]

    st.dataframe(
        df_filtrado[colunas_disponiveis].head(100),
        use_container_width=True,
        height=400
    )

    # Insights e Alertas
    st.markdown("---")
    st.header("💡 Insights e Alertas")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("### 🚨 Críticos")
        criticos = df_filtrado[df_filtrado['Status Final'].str.contains(
            'Acima de 30', na=False)]
        st.metric("Entregas > 30 dias", len(criticos))
        if len(criticos) > 0:
            st.error(f"Atenção: {len(criticos)} entregas com atraso crítico!")

    with col2:
        st.markdown("### ⚠️ Atenção")
        atencao = df_filtrado[df_filtrado['Status Final'].str.contains(
            '21 a 30|11 a 20', na=False)]
        st.metric("Entregas 11-30 dias", len(atencao))
        if len(atencao) > 0:
            st.warning(f"{len(atencao)} entregas necessitam acompanhamento")

    with col3:
        st.markdown("### ✅ Performance")
        st.metric("Taxa de Pontualidade", f"{taxa_pontualidade:.1f}%")
        if taxa_pontualidade >= 70:
            st.success("Performance dentro do esperado!")
        elif taxa_pontualidade >= 50:
            st.warning("Performance abaixo do ideal")
        else:
            st.error("Performance crítica - necessita ação imediata!")

except FileNotFoundError:
    st.error("⚠️ Arquivo não encontrado! Certifique-se de que o arquivo 'Monitoramento Tarefas  Braspress  30102025 Carlos.xlsx' está no mesmo diretório do script.")
    st.info(
        "📁 Coloque o arquivo Excel na mesma pasta do script Python e execute novamente.")
except Exception as e:
    st.error(f"❌ Erro ao processar os dados: {str(e)}")
    st.info("Verifique se o arquivo Excel está no formato correto e tente novamente.")