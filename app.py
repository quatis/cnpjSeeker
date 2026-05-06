import streamlit as st
import pandas as pd
from google.cloud import bigquery
from google.oauth2 import service_account
import time
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="Consulta de Dados",
    layout="wide"
)

st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    .stButton>button {
        width: 100%;
        border-radius: 4px;
        height: 3.5em;
        font-weight: bold;
    }
    .status-box {
        padding: 20px;
        border-radius: 8px;
        background-color: #1e1e1e;
        color: #00ff00;
        font-family: 'Courier New', Courier, monospace;
        border: 1px solid #333;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.5);
        margin-bottom: 20px;
    }
    .time-card {
        padding: 15px;
        background-color: #262730;
        color: #ffffff !important;
        border-radius: 5px;
        border: 1px solid #464646;
        margin-bottom: 10px;
        font-size: 0.9em;
    }
    .time-card b { color: #00ff00; }
    </style>
    """, unsafe_allow_html=True)


def get_bq_client():
    info = st.secrets["gcp_service_account"]
    credentials = service_account.Credentials.from_service_account_info(info)
    return bigquery.Client(credentials=credentials, project=credentials.project_id)


def formatar_cnpj(c):
    c = str(c).zfill(14)
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"


def main():
    st.title("Consulta em Lote")
    st.caption("Interface de busca baseada em dados públicos")
    st.divider()

    col1, col2 = st.columns([1, 2])

    with col1:
        st.subheader("Entrada")
        input_texto = st.text_area("Insira os identificadores:", height=300, placeholder="00000000000000")
        placeholder_btn_download = st.empty()
        btn_processar = st.button("Executar Consulta")

    with col2:
        st.subheader("Monitoramento")

        if btn_processar:
            identificadores = list(
                set(re.findall(r"\d{14}", input_texto.replace(".", "").replace("/", "").replace("-", ""))))

            if not identificadores:
                st.error("Nenhum registro válido encontrado.")
            else:
                start_time = time.time()
                client = get_bq_client()

                raizes = list(set([c[:8] for c in identificadores]))
                lista_sql = ", ".join([f"'{r}'" for r in raizes])

                query = f"""
                    SELECT 
                        cnpj_basico,
                        IF(opcao_simples = 1, 'Sim', 'Não') as Status
                    FROM `basedosdados.br_me_cnpj.simples`
                    WHERE cnpj_basico IN ({lista_sql})
                """

                with st.spinner("Processando via Google Cloud..."):
                    try:
                        query_job = client.query(query)
                        df_nuvem = query_job.to_dataframe()

                        df_base = pd.DataFrame({'CNPJ': identificadores})
                        df_base['cnpj_basico'] = df_base['CNPJ'].str[:8]

                        df_nuvem['cnpj_basico'] = df_nuvem['cnpj_basico'].astype(str).str.zfill(8)

                        df_final = pd.merge(df_base, df_nuvem, on='cnpj_basico', how='left')
                        df_final['Status'] = df_final['Status'].fillna('Não encontrado')
                        df_final['CNPJ'] = df_final['CNPJ'].apply(formatar_cnpj)

                        df_final.drop(columns=['cnpj_basico'], inplace=True)

                        tempo_total = time.time() - start_time

                        st.markdown(
                            f"<div class='time-card'><b>Tempo de Resposta:</b><br>{round(tempo_total, 2)} segundos</div>",
                            unsafe_allow_html=True)
                        st.success(f"Concluído. {len(df_final)} registros processados.")

                        st.dataframe(df_final, use_container_width=True, hide_index=True)

                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_final.to_excel(writer, index=False, sheet_name='Resultados')
                        excel_data = output.getvalue()

                        with placeholder_btn_download:
                            st.download_button(
                                label="Gerar Excel (xlsx)",
                                data=excel_data,
                                file_name=f"consulta_{datetime.now().strftime('%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                        del df_final
                        output.close()

                    except Exception as e:
                        st.error(f"Erro na operação: {e}")


if __name__ == "__main__":
    main()