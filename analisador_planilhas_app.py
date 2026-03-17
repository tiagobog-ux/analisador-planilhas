import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from io import BytesIO

st.set_page_config(layout="wide")

st.title("Dashboard Automático de Planilhas")

st.write("Envie uma ou mais planilhas Excel com colunas 'Nome' e 'Valor'.")

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

    total = len(consolidado)
    soma = consolidado["Valor"].sum()
    media = consolidado["Valor"].mean()
    maior = consolidado["Valor"].max()

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Total Registros", total)
    col2.metric("Soma Valores", round(soma, 2))
    col3.metric("Valor Médio", round(media, 2))
    col4.metric("Maior Valor", round(maior, 2))

    st.divider()

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

    ranking = consolidado.sort_values(by="Valor", ascending=False)
    top10 = ranking.head(10)
    bottom10 = ranking.tail(10)

    colA, colB = st.columns(2)

    with colA:
        st.subheader("Top 10")
        st.dataframe(top10)

    with colB:
        st.subheader("Bottom 10")
        st.dataframe(bottom10)

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

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        consolidado.to_excel(writer, sheet_name="Dados", index=False)
        top10.to_excel(writer, sheet_name="Top10", index=False)
        bottom10.to_excel(writer, sheet_name="Bottom10", index=False)

    output.seek(0)

    st.download_button(
        "Baixar relatório Excel",
        data=output,
        file_name="relatorio_dashboard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Envie um ou mais arquivos Excel para começar.")