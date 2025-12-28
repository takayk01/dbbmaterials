import streamlit as st
import pandas as pd

# importa parâmetros
from parameters import (
    path_source,
    path_result,
    columns_final,
    limpar_df
)

st.set_page_config(page_title="Processar Arquivo Excel", layout="centered")

st.title("📄 Processamento de Arquivo Excel")

# --- Entrada do usuário ---
file = st.text_input("Informe o código do arquivo:", value="G10009")

# --- Função original ---
def ler_arquivo(path):
    df = pd.read_excel(path, dtype={"unique_id": str})
    df = limpar_df(df)

    df = df.map(
        lambda x: ' '.join(str(x).split()) if isinstance(x, str) else x
    )

    df["description"] = df["column_name_value"].str.replace(
        "Description ", "", regex=False
    )

    df[["column_name", "description_name"]] = (
        df["description"].str.split(":", expand=True)
    )

    df["description_name"] = df["description_name"].str.title()

    df = df[df["column_name_value"] != ""][columns_final].copy()

    return df


# --- Botão de execução ---
if st.button("▶ Executar processamento"):

    try:
        with st.spinner("Lendo arquivo..."):
            df = ler_arquivo(path_source.replace("xxxxx", file))
            total = df.shape[0]

        with st.spinner("Processando dados..."):
            df = df.drop_duplicates()
            df = df.sort_values(
                ["unique_id", "description_name"],
                ascending=True
            )

        with st.spinner("Salvando resultado..."):
            df.to_excel(
                path_result.replace("xxxxx", file),
                index=False
            )

        final = df.shape[0]
        df_unique = df["unique_id"].drop_duplicates()

        # --- Resultados ---
        st.success("✅ Processo finalizado com sucesso!")

        st.write(f"**{final} de {total}**")
        st.write(f"**Materiais únicos:** {df_unique.shape[0]}")

        st.dataframe(df.head(20))

    except Exception as e:
        st.error("❌ Erro ao processar o arquivo")
        st.exception(e)
