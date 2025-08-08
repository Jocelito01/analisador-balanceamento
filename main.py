import streamlit as st
import shutil
import analisador
from analisador import gerar_excel_com_grafico

st.title("⚙️ Analisador de Balanceamento")

arquivo = st.file_uploader("📁 Faça upload do arquivo CSV", type=["csv"])
modelo = st.selectbox("Selecione o modelo:", ["1119", "1121", "1141","1144","4203", "4147","4203","4282","MB01","MB02","MB03", "TB01-1200", "TB01-1205"])

if st.button("🔍 Processar"):
  if not arquivo or not modelo:
      st.warning("Por favor, envie o arquivo e selecione um modelo.")
  else:
      caminho_entrada = "upload_temp.csv"
      with open(caminho_entrada, "wb") as f:
          shutil.copyfileobj(arquivo, f)

      try:
          analisador.Planilha(caminho_entrada)
          extrato_dados = analisador.extrato("ciclos_colunas_separadas.csv", modelo)
          analisador.Grafico(modelo, "ciclos_colunas_separadas.csv")

          # Mostrar dados no front
          st.success("✅ Processamento finalizado com sucesso!")
          st.write(f"Modelo: {extrato_dados['modelo']}")
          st.write(f"Peças boas: {extrato_dados['contador_ok']}")
          st.write(f"Reprovadas no balanceamento: {extrato_dados['contador_nok']}")
          st.write(f"Média de desbalanceamento primeira medição: {extrato_dados['mediaE1']:.2f}")
          st.write(f"Média do ângulo primeira medição: {extrato_dados['mediaA1']:.2f}")
          st.write(f"Média de desbalanceamento pós correção: {extrato_dados['mediaEF']:.2f}")
          st.write(f"Média do ângulo pós correção: {extrato_dados['mediaAF']:.2f}")

          # Gerar e disponibilizar Excel para download

          excel_data = gerar_excel_com_grafico(extrato_dados, "ciclos_colunas_separadas.csv")
          st.download_button(
              label="📥 Baixar relatório Excel",
              data=excel_data,
              file_name=f"relatorio_balanceamento_{modelo}.xlsx",
              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          )

          # Mostrar imagem do gráfico
          st.image(f"grafico_desbalanceamento_{modelo}.png", caption="Gráfico de Desbalanceamento")

      except Exception as e:
          st.error(f"❌ Erro durante o processamento: {e}")