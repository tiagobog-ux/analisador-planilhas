import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from io import BytesIO

st.set_page_config(layout="wide")

st.title("Dashboard Inteligente de Planilhas")

st.write("Envie uma ou mais planilhas Excel com colunas 'Nome' e 'Valor'.")

def classificar_valor(valor):
    if valor > 800:
        return "Excelente"
    elif valor >= 500:
        return "Bom"
    elif valor >= 200:
        return "Médio"
    else:
        return "Baixo"

arquivos = st.file_uploader(
    "Escolha os arquivos Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if arquivos:

    dados_consolidados = []

    for arquivo in arquivos:
        df = pd.read_excel(arquivo)
        df["Arquivo"] = arquivo.name
        dados_consolidados.append(df)

    consolidado = pd.concat(dados_consolidados, ignore_index=True)

    # classificação automática
    consolidado["Classificação"] = consolidado["Valor"].apply(classificar_valor)

    # métricas
    total = len(consolidado)
    soma = consolidado["Valor"].sum()
    media = consolidado["Valor"].mean()
    maior = consolidado["Valor"].max()
    menor = consolidado["Valor"].min()

    # acima e abaixo da média
    acima_media = consolidado[consolidado["Valor"] > media]
    abaixo_media = consolidado[consolidado["Valor"] <= media]

    # top e bottom
    ranking = consolidado.sort_values(by="Valor", ascending=False)
    top10 = ranking.head(10)
    bottom10 = ranking.tail(10)

    # resumo por classificação
    resumo_classificacao = consolidado["Classificação"].value_counts().reset_index()
    resumo_classificacao.columns = ["Classificação", "Quantidade"]

    # métricas visuais
    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric("Total Registros", total)
    col2.metric("Soma Valores", round(soma, 2))
    col3.metric("Valor Médio", round(media, 2))
    col4.metric("Maior Valor", round(maior, 2))
    col5.metric("Menor Valor", round(menor, 2))

    st.divider()

    # insights automáticos
    qtd_excelente = len(consolidado[consolidado["Classificação"] == "Excelente"])
    qtd_bom = len(consolidado[consolidado["Classificação"] == "Bom"])
    qtd_medio = len(consolidado[consolidado["Classificação"] == "Médio"])
    qtd_baixo = len(consolidado[consolidado["Classificação"] == "Baixo"])

    st.subheader("Insights automáticos")

    st.write(f"A média geral dos valores é **{media:.2f}**.")
    st.write(f"Há **{len(acima_media)} registros acima da média** e **{len(abaixo_media)} abaixo ou iguais à média**.")
    st.write(f"Distribuição por classificação: **{qtd_excelente} Excelentes**, **{qtd_bom} Bons**, **{qtd_medio} Médios** e **{qtd_baixo} Baixos**.")

    st.divider()

    # filtro
    valor_minimo = st.slider(
        "Mostrar valores acima de:",
        int(consolidado["Valor"].min()),
        int(consolidado["Valor"].max()),
        int(consolidado["Valor"].min())
    )

    filtrado = consolidado[consolidado["Valor"] >= valor_minimo]

    st.subheader("Dados filtrados")
    st.dataframe(filtrado)

    st.divider()

    colA, colB = st.columns(2)

    with colA:
        st.subheader("Top 10")
        st.dataframe(top10)

    with colB:
        st.subheader("Bottom 10")
        st.dataframe(bottom10)

    st.divider()

    colC, colD = st.columns(2)

    with colC:
        st.subheader("Acima da média")
        st.dataframe(acima_media)

    with colD:
        st.subheader("Abaixo ou igual à média")
        st.dataframe(abaixo_media)

    st.divider()

    st.subheader("Resumo por classificação")
    st.dataframe(resumo_classificacao)

    st.divider()

    st.subheader("Gráfico Top 10")
    fig, ax = plt.subplots()
    ax.bar(top10["Nome"], top10["Valor"])
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(fig)

    st.subheader("Distribuição dos Valores")
    fig2, ax2 = plt.subplots()
    ax2.hist(consolidado["Valor"], bins=10)
    st.pyplot(fig2)

    # gerar relatório Excel
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        consolidado.to_excel(writer, sheet_name="Dados", index=False)
        top10.to_excel(writer, sheet_name="Top10", index=False)
        bottom10.to_excel(writer, sheet_name="Bottom10", index=False)
        acima_media.to_excel(writer, sheet_name="AcimaMedia", index=False)
        abaixo_media.to_excel(writer, sheet_name="AbaixoMedia", index=False)
        resumo_classificacao.to_excel(writer, sheet_name="Classificacao", index=False)

    output.seek(0)

    st.download_button(
        "Baixar relatório Excel",
        data=output,
        file_name="relatorio_dashboard_inteligente.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Envie um ou mais arquivos Excel para começar.")
