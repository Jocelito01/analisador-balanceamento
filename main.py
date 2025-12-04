import streamlit as st
import shutil
import analisador
from analisador import gerar_excel_com_grafico
from datetime import date, datetime, time
import pandas as pd

# Configuração da página
st.set_page_config(page_title="Analisador de Balanceamento", layout="wide")

st.markdown("""
    <style>
    h1, h2, h3 { text-align: center; }
    div[class*="stFileUploader"] label, div[class*="stSelectbox"] label, 
    div[class*="stDateInput"] label, div[class*="stTimeInput"] label {
        font-size: 20px !important; font-weight: bold !important; text-align: center !important;
        display: block !important; width: 100% !important;
    }
    div[data-testid="stSelectbox"] > div > div { text-align: center; }
    div.stButton > button { font-size: 18px !important; font-weight: bold !important; height: 3em; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1>⚙️ Analisador de Balanceamento</h1>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

col_up, col_mod = st.columns(2)
with col_up:
    arquivo = st.file_uploader("📁 Faça upload do arquivo CSV", type=["csv"])
with col_mod:
    modelo = st.selectbox("Selecione o modelo:", ["1119", "1121", "1141","1144","4203", "4147","4203","4282","MB01","MB02","MB03", "TB01-1200", "TB01-1205"])

st.markdown("---")
st.markdown("<h3>📅 Filtro de Período</h3>", unsafe_allow_html=True)
col_d1, col_d2 = st.columns(2)
with col_d1:
    st.markdown("<div style='text-align: center; font-weight: bold; font-size: 18px;'>Início</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1: d_inicio = st.date_input("Data", value=date.today(), key="d_ini")
    with c2: t_inicio = st.time_input("Hora", value=time(0, 0), key="t_ini")
    data_inicio = datetime.combine(d_inicio, t_inicio)
with col_d2:
    st.markdown("<div style='text-align: center; font-weight: bold; font-size: 18px;'>Fim</div>", unsafe_allow_html=True)
    c3, c4 = st.columns(2)
    with c3: d_fim = st.date_input("Data", value=date.today(), key="d_fim")
    with c4: t_fim = st.time_input("Hora", value=time(23, 59), key="t_fim")
    data_fim = datetime.combine(d_fim, t_fim)

st.markdown("<br><br>", unsafe_allow_html=True)
c_esq, c_meio, c_dir = st.columns([1, 2, 1]) 
with c_meio:
    botao_processar = st.button("🔍 PROCESSAR DADOS", use_container_width=True)

container_metricas = st.container()
container_downloads = st.container()
container_grafico = st.container()
container_tabela = st.container()

if botao_processar or st.session_state.get('processado', False):
    st.session_state['processado'] = True

    if not arquivo:
        if 'arquivo_salvo' not in st.session_state and not arquivo: st.warning("⚠️ Por favor, envie o arquivo."); st.stop()

    if arquivo:
        arquivo.seek(0)
        caminho_entrada = "upload_temp.csv"
        caminho_csv_separado = "ciclos_colunas_separadas.csv"
        with open(caminho_entrada, "wb") as f: shutil.copyfileobj(arquivo, f)

        try:
            # 1. Processamento
            analisador.Planilha(caminho_entrada, data_inicio, data_fim)
            extrato_dados = analisador.extrato(caminho_csv_separado, modelo)
            analisador.Grafico(modelo, caminho_csv_separado)

            # 2. Renderizar Tabela
            df_tabela = pd.read_csv(caminho_csv_separado)
            df_tabela.insert(0, 'Item', range(1, len(df_tabela) + 1))
            ids_para_destaque = None

            with container_tabela:
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("<h3 style='text-align: center;'>📋 Detalhamento dos Dados</h3>", unsafe_allow_html=True)

                opcoes_status = df_tabela['Status Final'].unique().tolist()
                col_filt1, col_filt2 = st.columns([1, 6])
                with col_filt1: st.markdown("**Filtrar Status:**")
                with col_filt2:
                    status_selecionados = st.multiselect("Selecione status", options=opcoes_status, default=opcoes_status, label_visibility="collapsed")

                df_view = df_tabela[df_tabela['Status Final'].isin(status_selecionados)]

                event = st.dataframe(df_view, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="multi-row", key="tabela_dados")
                if event.selection.rows:
                    ids_para_destaque = df_view.iloc[event.selection.rows]['Rotor ID'].tolist()

            # 3. Renderizar Gráfico
            with container_grafico:
                fig_interativo = analisador.GraficoInterativo(modelo, caminho_csv_separado, ids_destaque=ids_para_destaque)
                col_img1, col_img2, col_img3 = st.columns([1, 10, 1])
                with col_img2:
                    # --- AQUI ESTÁ A CORREÇÃO DE ALINHAMENTO ---
                    # Título agora é gerado pelo Streamlit, fora do Plotly
                    st.markdown(f"<h3 style='text-align: center; margin-bottom: 0px;'>Gráfico Interativo - Modelo {modelo}</h3>", unsafe_allow_html=True)

                    if fig_interativo:
                        st.plotly_chart(fig_interativo, use_container_width=True)
                    else:
                        st.image(f"grafico_desbalanceamento_{modelo}.png", caption="Gráfico Estático")

            with container_metricas:
                st.markdown("<br>", unsafe_allow_html=True)
                st.success("✅ Processamento finalizado com sucesso!")
                st.markdown(f"<h5 style='text-align: center; color: gray;'>🕒 Período: {data_inicio.strftime('%d/%m/%Y %H:%M')} até {data_fim.strftime('%d/%m/%Y %H:%M')}</h5>", unsafe_allow_html=True)
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Peças OK", extrato_dados['contador_ok'])
                m2.metric("Peças NOK", extrato_dados['contador_nok'])
                m3.metric("Média Desb. 1", f"{extrato_dados['mediaE1']:.2f}")
                m4.metric("Média Desb. Final", f"{extrato_dados['mediaEF']:.2f}")

            with container_downloads:
                excel_data = gerar_excel_com_grafico(extrato_dados, caminho_csv_separado)
                with open(caminho_csv_separado, "rb") as f: csv_data = f.read()
                st.markdown("---")
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    st.download_button("📥 BAIXAR RELATÓRIO EXCEL", excel_data, f"Relatorio de balanceamento {modelo}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with col_btn2:
                     st.download_button("📄 BAIXAR CSV SEPARADO", csv_data, f"Ciclos com colunas separadas {modelo}.csv", "text/csv", use_container_width=True)

        except Exception as e:
            st.error(f"❌ Erro: {e}")