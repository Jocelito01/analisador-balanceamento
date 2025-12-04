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


def Planilha(arquivo_entrada):
    arquivo_saida = 'ciclos_colunas_separadas.csv'
    grupos_por_rotor = defaultdict(list)

    with open(arquivo_entrada, encoding='utf-8') as f_in:
        # Ler a primeira linha (cabeçalho)
        header_line = f_in.readline()
        # Remove aspas e espaços extras
        cleaned_header = [h.strip().replace('"', '') for h in header_line.split(',')]

        # Leia o resto do arquivo
        linhas_restantes = f_in.readlines()

    # Agora reabra para leitura com DictReader manualmente criado
    import io
    novo_arquivo = io.StringIO()
    novo_arquivo.write(','.join(cleaned_header) + '\n')
    novo_arquivo.writelines(linhas_restantes)
    novo_arquivo.seek(0)

    leitor = csv.DictReader(novo_arquivo, delimiter=',')

    for linha in leitor:
        rotor_id = linha.get('Rotor ID')
        if rotor_id:
            grupos_por_rotor[rotor_id].append(linha)

    # Verificações
    if not grupos_por_rotor:
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
                static = l.get('Static [gmm]', '')
                angle = l.get('Angle') or l.get('Angle"') or ''
                linha_saida.append(static)
                linha_saida.append(angle)

            while len(linha_saida) < len(cabecalho):
                linha_saida.append('')

            escritor.writerow(linha_saida)

    print(f"✅ Novo arquivo CSV salvo como '{arquivo_saida}' com colunas separadas.")

def filtro(arquivo_csv):
    with open(arquivo_csv, newline='', encoding='utf-8') as fil:
        leitorData = list(csv.DictReader(fil))


def extrato(arquivo_csv, modelo):
    contador_ok = 0
    contador_nok = 0
    valoresE = []
    valoresA = []

    with open(arquivo_csv, newline='', encoding='utf-8') as f:
        leitor = list(csv.DictReader(f))
        for linha in leitor:
            estatico = linha.get('Static [gmm] 1', '').strip()
            if estatico:
                try:
                    valoresE.append(float(estatico))
                except ValueError:
                    pass

        for linha in leitor:
            angulo = linha.get('Angle 1', '').strip()
            if angulo:
                try:
                    valoresA.append(float(angulo))
                except ValueError:
                    pass

        for linha in leitor:
            status = linha.get('Status Final', '').strip().upper()
            if status == 'OK':
                contador_ok += 1
            elif status == 'NOK':
                contador_nok += 1

    mediaE = sum(valoresE) / len(valoresE) if valoresE else 0
    mediaA = sum(valoresA) / len(valoresA) if valoresA else 0

    return {
        "modelo": modelo,
        "contador_ok": contador_ok,
        "contador_nok": contador_nok,
        "mediaE": mediaE,
        "mediaA": mediaA
    }


def Grafico(modelo, arquivo):
    fundo_cor = '#ffffff'

    if modelo == '4147':
        theta = np.linspace(0, 2*np.pi, 360)
        raio_inferior = 0
        raio_superior = 40
    elif modelo == 'TB01-1200':
        theta = np.linspace(3.83972, 4.01426, 360)
        raio_inferior = 330
        raio_superior = 370
    elif modelo == 'TB01-1205':
        theta = np.linspace(4.01426, 4.18879, 360)
        raio_inferior = 420
        raio_superior = 460
    elif modelo == '1121':
        theta = np.linspace(3.26377, 4.31096, 360)
        raio_inferior = 128
        raio_superior = 172
    elif modelo == '1141':
        theta = np.linspace(3.76991, 4.81711, 360)
        raio_inferior = 32
        raio_superior = 72
    else:
        theta = np.linspace(0, 2*np.pi, 360)
        raio_inferior = 0
        raio_superior = 30

    valores_por_par = []

    with open(arquivo, newline='', encoding='utf-8') as arquivocsv:
        leitor = csv.DictReader(arquivocsv)

        static_cols = [c for c in leitor.fieldnames if c.startswith('Static')]
        angle_cols = [c for c in leitor.fieldnames if c.startswith('Angle')]

        for _ in range(len(static_cols)):
            valores_por_par.append(([], []))

        for linha in leitor:
            for i, (sc, ac) in enumerate(zip(static_cols, angle_cols)):
                try:
                    raio = float(linha[sc])
                    angulo_graus = float(linha[ac])
                    angulo_rad = math.radians(angulo_graus)
                    valores_por_par[i][0].append(raio)
                    valores_por_par[i][1].append(angulo_rad)
                except (ValueError, KeyError):
                    continue

    plt.figure(figsize=(8, 8), facecolor=fundo_cor)
    ax = plt.subplot(111, polar=True, facecolor=fundo_cor)

    ax.fill_between(theta, raio_inferior, raio_superior, color='lightgreen', alpha=0.9)

    if modelo == 'TB01-1200':
        ax.set_ylim(0, 400)
    elif modelo == 'TB01-1205':
        ax.set_ylim(0, 500)
    elif modelo == '1121':
        ax.set_ylim(0, 250)
    else:
        ax.set_ylim(0, 100)

    cores = ['blue', 'red', 'green', 'orange', 'purple', 'brown', 'cyan']

    for i, (raios, thetas) in enumerate(valores_por_par):
        if raios and thetas:
            cor = cores[i % len(cores)]
            ax.scatter(thetas, raios, color=cor, label=f'Static/Angle {i+1}', s=5, alpha=0.7)

    plt.title('Gráfico de Desbalanceamento')
    plt.legend(loc='upper right')
    plt.savefig(f'grafico_desbalanceamento_{modelo}.png')
    plt.close()
    print(f"✅ Gráfico salvo como 'grafico_desbalanceamento_{modelo}.png'.")

def gerar_excel_com_grafico(dados_extrato, arquivo_csv):
    # Lê os dados do CSV
    df = pd.read_csv(arquivo_csv)

    # Cria um Excel em memória
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Aba 1: Dados do extrato
        df_extrato = pd.DataFrame([{
            "Modelo": dados_extrato["modelo"],
            "Peças OK": dados_extrato["contador_ok"],
            "Peças NOK": dados_extrato["contador_nok"],
            "Média Static [gmm] 1": dados_extrato["mediaE"],
            "Média Ângulo 1": dados_extrato["mediaA"]
        }])
        df_extrato.to_excel(writer, sheet_name="Resumo", index=False)

        # Aba 2: Dados detalhados
        df.to_excel(writer, sheet_name="Dados", index=False)

        # Salva o Excel até aqui
        writer.book.save(output)

    # Reabre com openpyxl para inserir o gráfico como imagem
    output.seek(0)
    wb = load_workbook(output)
    ws = wb["Resumo"]

    # Carrega imagem do gráfico
    try:
        img = Image.open(f"grafico_desbalanceamento_{dados_extrato['modelo']}.png")
        img_path = f"grafico_temp_{dados_extrato['modelo']}.png"
        img.save(img_path)

        xl_img = XLImage(img_path)
        xl_img.anchor = "A7"  # Posição da imagem na planilha
        ws.add_image(xl_img)
    except Exception as e:
        print(f"⚠️ Erro ao inserir imagem no Excel: {e}")

    # Salva novamente no buffer
    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output