import streamlit as st
import pandas as pd
from io import BytesIO

# =========================================================
# Funções utilitárias
# =========================================================

def limpar_df(df):
    df = df.fillna("")
    df = df.replace("nan", "")
    df = df.astype(str)    
    return df

def ler_arquivo(uploaded_file):
    df = pd.read_excel(uploaded_file, dtype={"unique_id": str})
    df = limpar_df(df)

    df["description"] = df["column_name_value"].str.replace( "Description ", "", regex=False )
    df[["RESULT_COLUMN", "RESULT_DESCRIPTION"]] = ( df["description"].str.split(":", expand=True) )
    df["RESULT_DESCRIPTION"] = df["RESULT_DESCRIPTION"].str.title()
    df = df[df["column_name_value"] != ""].copy()    
    df = df.drop(columns=["description"])    
    
    # Aplica Regra do arquivo
    df["RESULT_DESCRIPTION"] = df["RESULT_DESCRIPTION"].apply(lambda x: " ".join(str(x).split()) if isinstance(x, str) else x)

    return df


# =========================================================
# Interface Streamlit
# =========================================================
file_name = 'GMM10081'

st.set_page_config(
    page_title=f"DBB Materiais - Arquivo {file_name}",
    layout="centered",
)

st.title(f"📄 Correção de Materiais")
st.title(f"Aplicação da Regra {file_name}")

st.markdown(
    """   
        - **Nome da Regra: Descrição do material: Caso misto - GMM10081**
        - __Rule Name: Material Description Is Mixed Case - GMM10081__
    """
)

st.markdown(
    """   
        - **Descrição da Regra: Não é permitido usar apenas letras maiúsculas na descrição do material.**
        - __Rule Description: It is not allowed to use only capital letters in Material Description__
    """
)

uploaded_file = st.file_uploader(
    "Selecione o arquivo (Excel)",
    type=["xlsx", "xls"],
)

# =========================================================
# Execução
# =========================================================

if uploaded_file is not None:

    if st.button("▶ Aplicar Regra"):
        try:
            with st.spinner("Lendo arquivo..."):
                df = ler_arquivo(uploaded_file)
                total = df.shape[0]

            with st.spinner("Processando dados..."):
                df = df.drop_duplicates()
                df = df.sort_values(
                    ["unique_id"],
                    ascending=True,
                )

            final = df.shape[0]
            df_unique = df["unique_id"].drop_duplicates()

            st.success("✅ Processo finalizado com sucesso!")

            st.write(f"**{final} de {total}**")
            st.write(f"**Materiais únicos:** {df_unique.shape[0]}")

            st.dataframe(df.head(20))

            # =================================================
            # Download do resultado (Excel em memória)
            # =================================================

            output_file = f"{file_name}_result.xlsx"

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df.to_excel(
                    writer,
                    index=False,
                    sheet_name=file_name,
                )

            buffer.seek(0)

            st.download_button(
                label="⬇️ Baixar resultado em Excel",
                data=buffer,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error("❌ Erro ao processar o arquivo")
            st.exception(e)

else:
    st.info("📎 Faça o upload de um arquivo Excel para iniciar.")







