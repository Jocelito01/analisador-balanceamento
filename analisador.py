import csv
import math
import matplotlib.pyplot as plt
import numpy as np
from collections import defaultdict
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
import streamlit as st
from datetime import datetime

# Tenta importar o Plotly de forma segura
try:
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

def Planilha(arquivo_entrada, data_inicio=None, data_fim=None):
    arquivo_saida = 'ciclos_colunas_separadas.csv'
    grupos_por_rotor = defaultdict(list)

    with open(arquivo_entrada, encoding='utf-8') as f_in:
        header_line = f_in.readline()
        cleaned_header = [h.strip().replace('"', '') for h in header_line.split(',')]
        linhas_restantes = f_in.readlines()

    import io
    novo_arquivo = io.StringIO()
    novo_arquivo.write(','.join(cleaned_header) + '\n')
    novo_arquivo.writelines(linhas_restantes)
    novo_arquivo.seek(0)

    leitor = csv.DictReader(novo_arquivo, delimiter=',')

    for linha in leitor:
        if data_inicio and data_fim:
            data_str = linha.get('#Time', '').strip() 
            if data_str:
                try:
                    data_obj = datetime.strptime(data_str, '%m/%d/%Y %H:%M:%S')
                    if not (data_inicio <= data_obj <= data_fim):
                        continue
                except ValueError:
                    pass

        rotor_id = linha.get('Rotor ID')
        if rotor_id:
            grupos_por_rotor[rotor_id].append(linha)

    if not grupos_por_rotor:
         if data_inicio and data_fim:
             raise ValueError(f"Nenhum dado encontrado para o período entre {data_inicio} e {data_fim}.")
         else:
             raise ValueError("Nenhum dado válido encontrado para 'Rotor ID'.")

    max_leituras = max(len(linhas) for linhas in grupos_por_rotor.values())

    cabecalho = ['Rotor ID', 'Status Final']
    for i in range(1, max_leituras + 1):
        cabecalho.append(f'Static [gmm] {i}')
        cabecalho.append(f'Angle {i}')

    with open(arquivo_saida, 'w', newline='', encoding='utf-8') as f_out:
        escritor = csv.writer(f_out)
        escritor.writerow(cabecalho)

        for rotor_id, linhas in grupos_por_rotor.items():
            ultima = linhas[-1]
            status = 'OK' if ultima.get('Tolerance', '').strip().upper() == 'Y' else 'NOK'

            linha_saida = [rotor_id, status]
            for l in linhas:
                static = l.get('Static [gmm]') or l.get('Amount 1 [gmm]') or ''
                angle = l.get('Angle') or l.get('Angle"') or l.get('Angle 1') or ''
                linha_saida.append(static)
                linha_saida.append(angle)

            while len(linha_saida) < len(cabecalho):
                linha_saida.append('')

            escritor.writerow(linha_saida)
    print(f"✅ Novo arquivo CSV salvo como '{arquivo_saida}' com colunas separadas.")

def extrato(arquivo_csv, modelo):
    contador_ok = 0
    contador_nok = 0
    valoresE1 = []
    valoresA1 = []
    valoresEF = []
    valoresAF = []

    with open(arquivo_csv, newline='', encoding='utf-8') as f:
        leitor = list(csv.DictReader(f))
        for linha in leitor:
            estatico = linha.get('Static [gmm] 1', '').strip()
            if estatico:
                try:
                    valoresE1.append(float(estatico))
                except ValueError:
                    pass

        for linha in leitor:
            angulo = linha.get('Angle 1', '').strip()
            if angulo:
                try:
                    valoresA1.append(float(angulo))
                except ValueError:
                    pass
        for linha in leitor:
            estatico = linha.get('Static [gmm] 2', '').strip()
            if estatico:
                try:
                    valoresEF.append(float(estatico))
                except ValueError:
                    pass

        for linha in leitor:
            angulo = linha.get('Angle 2', '').strip()
            if angulo:
                try:
                    valoresAF.append(float(angulo))
                except ValueError:
                    pass
        for linha in leitor:
            status = linha.get('Status Final', '').strip().upper()
            if status == 'OK':
                contador_ok += 1
            elif status == 'NOK':
                contador_nok += 1

    mediaE1 = sum(valoresE1) / len(valoresE1) if valoresE1 else 0
    mediaA1 = sum(valoresA1) / len(valoresA1) if valoresA1 else 0
    mediaEF = sum(valoresEF) / len(valoresEF) if valoresEF else 0
    mediaAF = sum(valoresAF) / len(valoresAF) if valoresAF else 0

    return {
        "modelo": modelo,
        "contador_ok": contador_ok,
        "contador_nok": contador_nok,
        "mediaE1": mediaE1,
        "mediaA1": mediaA1,
        "mediaEF": mediaEF,
        "mediaAF": mediaAF
    }

def Grafico(modelo, arquivo):
    fundo_cor = '#ffffff'
    if modelo == '4147': theta = np.linspace(0, 2*np.pi, 360); raio_inferior = 0; raio_superior = 40
    elif modelo == 'TB01-1200': theta = np.linspace(3.83972, 4.01426, 360); raio_inferior = 330; raio_superior = 370
    elif modelo == 'TB01-1205': theta = np.linspace(4.01426, 4.18879, 360); raio_inferior = 420; raio_superior = 460
    elif modelo == '1121': theta = np.linspace(3.26377, 4.31096, 360); raio_inferior = 128; raio_superior = 172
    elif modelo == '1141': theta = np.linspace(3.76991, 4.81711, 360); raio_inferior = 32; raio_superior = 72
    else: theta = np.linspace(0, 2*np.pi, 360); raio_inferior = 0; raio_superior = 30

    valores_por_par = []
    with open(arquivo, newline='', encoding='utf-8') as arquivocsv:
        leitor = csv.DictReader(arquivocsv)
        static_cols = [c for c in leitor.fieldnames if c.startswith('Static')]
        angle_cols = [c for c in leitor.fieldnames if c.startswith('Angle')]
        for _ in range(len(static_cols)): valores_por_par.append(([], []))
        for linha in leitor:
            for i, (sc, ac) in enumerate(zip(static_cols, angle_cols)):
                try:
                    raio = float(linha[sc])
                    angulo_graus = float(linha[ac])
                    angulo_rad = math.radians(angulo_graus)
                    valores_por_par[i][0].append(raio)
                    valores_por_par[i][1].append(angulo_rad)
                except (ValueError, KeyError): continue

    plt.figure(figsize=(8, 8), facecolor=fundo_cor)
    ax = plt.subplot(111, polar=True, facecolor=fundo_cor)
    ax.fill_between(theta, raio_inferior, raio_superior, color='lightgreen', alpha=0.9)
    if modelo == 'TB01-1200': ax.set_ylim(0, 400)
    elif modelo == 'TB01-1205': ax.set_ylim(0, 500)
    elif modelo == '1121': ax.set_ylim(0, 250)
    else: ax.set_ylim(0, 100)
    cores = ['blue', 'red', 'green', 'orange', 'purple', 'brown', 'cyan']
    for i, (raios, thetas) in enumerate(valores_por_par):
        if raios and thetas:
            cor = cores[i % len(cores)]
            ax.scatter(thetas, raios, color=cor, label=f'Medição {i+1}', s=5, alpha=0.7)
    plt.title('Gráfico de Desbalanceamento')
    plt.legend(loc='upper right')
    plt.savefig(f'grafico_desbalanceamento_{modelo}.png')
    plt.close()
    print(f"✅ Imagem PNG salva para Excel.")

# --- ATUALIZAÇÃO: AGORA RECEBE IDs ---
def GraficoInterativo(modelo, arquivo, ids_destaque=None):
    if not PLOTLY_AVAILABLE: return None

    if ids_destaque is not None:
        ids_destaque = set(str(x) for x in ids_destaque)

    raio_max_chart = 100
    green_r_min, green_r_max = 0, 30

    if modelo == '4147': green_r_min, green_r_max, raio_max_chart = 0, 40, 100
    elif modelo == 'TB01-1200': green_r_min, green_r_max, raio_max_chart = 330, 370, 400
    elif modelo == 'TB01-1205': green_r_min, green_r_max, raio_max_chart = 420, 460, 500
    elif modelo == '1121': green_r_min, green_r_max, raio_max_chart = 128, 172, 250
    elif modelo == '1141': green_r_min, green_r_max, raio_max_chart = 32, 72, 100

    valores_por_par = []

    with open(arquivo, newline='', encoding='utf-8') as arquivocsv:
        leitor = list(csv.DictReader(arquivocsv))
        if not leitor: return go.Figure()

        static_cols = [c for c in leitor[0].keys() if c.startswith('Static')]
        angle_cols = [c for c in leitor[0].keys() if c.startswith('Angle')]

        for _ in range(len(static_cols)):
            valores_por_par.append({'r': [], 'theta': [], 'ids': [], 'colors': [], 'sizes': [], 'opacities': []})

        base_colors = ['blue', 'red', 'green', 'orange', 'purple', 'brown', 'cyan']

        for linha in leitor:
            rid = linha.get('Rotor ID', 'N/A')
            rid_str = str(rid)

            is_selected = True
            if ids_destaque and len(ids_destaque) > 0:
                if rid_str not in ids_destaque:
                    is_selected = False

            for i, (sc, ac) in enumerate(zip(static_cols, angle_cols)):
                try:
                    raio = float(linha[sc])
                    angulo_graus = float(linha[ac])

                    valores_por_par[i]['r'].append(raio)
                    valores_por_par[i]['theta'].append(angulo_graus)
                    valores_por_par[i]['ids'].append(rid)

                    cor_base = base_colors[i % len(base_colors)]

                    if is_selected:
                        valores_por_par[i]['colors'].append(cor_base)
                        valores_por_par[i]['sizes'].append(10 if ids_destaque else 8)
                        valores_por_par[i]['opacities'].append(1.0)
                    else:
                        valores_por_par[i]['colors'].append('lightgray')
                        valores_por_par[i]['sizes'].append(5)
                        valores_por_par[i]['opacities'].append(0.3)

                except (ValueError, KeyError):
                    continue

    fig = go.Figure()

    if green_r_min == 0:
        fig.add_trace(go.Scatterpolar(
            r=[green_r_max] * 360, theta=list(range(360)), mode='lines', fill='toself',
            fillcolor='rgba(144, 238, 144, 0.5)', line=dict(color='green', width=1),
            name='Tolerância', hoverinfo='skip'
        ))
    else:
        fig.add_trace(go.Scatterpolar(r=[green_r_min]*360, theta=list(range(360)), mode='lines',
                                      line=dict(color='green', width=1, dash='dash'), name='Min Tol', hoverinfo='skip'))
        fig.add_trace(go.Scatterpolar(r=[green_r_max]*360, theta=list(range(360)), mode='lines',
                                      line=dict(color='green', width=1, dash='dash'), name='Max Tol', hoverinfo='skip'))

    for i, dados in enumerate(valores_por_par):
        if dados['r']:
            fig.add_trace(go.Scatterpolar(
                r=dados['r'], theta=dados['theta'], mode='markers',
                marker=dict(
                    color=dados['colors'], 
                    size=dados['sizes'], 
                    opacity=dados['opacities'],
                    line=dict(color='white', width=0.5)
                ),
                name=f'Medição {i+1}', text=dados['ids'],
                hovertemplate="<b>Rotor: %{text}</b><br>Raio: %{r:.2f}<br>Ângulo: %{theta:.2f}°"
            ))

    fig.update_layout(
        # --- REMOVIDO O TITLE DAQUI ---
        polar=dict(radialaxis=dict(visible=True, range=[0, raio_max_chart]), angularaxis=dict(direction="counterclockwise", rotation=0)),
        showlegend=True, 
        template="plotly_white", 
        height=600, 
        # Reduzi a margem superior (t) de 50 para 20
        margin=dict(l=50, r=50, t=20, b=50) 
    )
    return fig

def gerar_excel_com_grafico(dados_extrato, arquivo_csv):
    df = pd.read_csv(arquivo_csv)
    df.insert(0, 'Item', range(1, len(df) + 1))
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_extrato = pd.DataFrame([{
            "Modelo": dados_extrato["modelo"],
            "Peças OK": dados_extrato["contador_ok"],
            "Peças NOK": dados_extrato["contador_nok"],
            "Média Desbalanceamento 1": dados_extrato["mediaE1"],
            "Média Ângulo 1": dados_extrato["mediaA1"],
            "Média Desbalanceamento 2": dados_extrato['mediaEF'],
            "Média Ângulo 2": dados_extrato['mediaAF']
        }])
        df_extrato.to_excel(writer, sheet_name="Resumo", index=False)
        df.to_excel(writer, sheet_name="Dados", index=False)
        writer.book.save(output)

    output.seek(0)
    wb = load_workbook(output)
    ws = wb["Resumo"]
    try:
        img = Image.open(f"grafico_desbalanceamento_{dados_extrato['modelo']}.png")
        img_path = f"grafico_temp_{dados_extrato['modelo']}.png"
        img.save(img_path)
        xl_img = XLImage(img_path)
        xl_img.anchor = "A7"
        ws.add_image(xl_img)
    except Exception as e:
        print(f"⚠️ Erro ao inserir imagem no Excel: {e}")
    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output