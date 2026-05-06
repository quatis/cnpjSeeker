import streamlit as st
import pandas as pd
import requests
import time
import re
from io import BytesIO
from datetime import datetime
from google.cloud import bigquery
from google.oauth2 import service_account

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
    .status-label { color: #aaaaaa; font-size: 0.9em; }
    .status-value { color: #ffffff; font-weight: bold; font-size: 1.1em; }
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
    .card-individual {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #1f4e79;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
    }
    </style>
    """, unsafe_allow_html=True)


# --- MOTORES DE CONSULTA ---

def get_bq_client():
    info = st.secrets["gcp_service_account"]
    credentials = service_account.Credentials.from_service_account_info(info)
    return bigquery.Client(credentials=credentials, project=credentials.project_id)


def consultar_brasilapi(cnpj, tentativas=3):
    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
    for i in range(tentativas):
        try:
            resp = requests.get(url, timeout=25)
            if resp.status_code == 200:
                d = resp.json()
                simples = d.get("opcao_pelo_simples")
                situacao = d.get("descricao_situacao_cadastral", "").upper()
                obs_parts = []
                if situacao == "BAIXADA": obs_parts.append("BAIXADA")
                status = "Sim" if simples is True else "Não"
                if simples is None: obs_parts.append("Sem registro")
                return {
                    "CNPJ": cnpj, "Status": status, "Razão Social": d.get("razao_social", "N/A"),
                    "Data Opção": d.get("data_opcao_pelo_simples", ""),
                    "Data Exclusão": d.get("data_exclusao_do_simples", ""), "OBS": " | ".join(obs_parts)
                }
            elif resp.status_code == 429:
                time.sleep(7)
                continue
            return {"CNPJ": cnpj, "Status": "Não", "Razão Social": "Não encontrado", "OBS": f"Erro {resp.status_code}"}
        except:
            time.sleep(2)
    return {"CNPJ": cnpj, "Status": "Erro", "Razão Social": "Falha Rede", "OBS": "Timeout"}


def formatar_cnpj(c):
    c = str(c).zfill(14)
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"


# --- ABAS DA INTERFACE ---

def aba_simples_nacional_lote():
    st.title("Simples Nacional (BigQuery)")
    st.caption("Consulta em massa via Google Cloud Platform")
    st.divider()

    col1, col2 = st.columns([1, 2])
    with col1:
        st.subheader("Entrada")
        input_texto = st.text_area("Insira os identificadores:", height=300, placeholder="00000000000000")
        placeholder_btn_download = st.empty()
        btn_processar = st.button("Executar Consulta BigQuery")

    with col2:
        st.subheader("Resultados")
        if btn_processar:
            identificadores = list(
                set(re.findall(r"\d{14}", input_texto.replace(".", "").replace("/", "").replace("-", ""))))
            if not identificadores:
                st.error("Nenhum identificador válido encontrado.")
            else:
                start_time = time.time()
                try:
                    client = get_bq_client()
                    raizes = list(set([c[:8] for c in identificadores]))
                    lista_sql = ", ".join([f"'{r}'" for r in raizes])

                    query = f"""
                        SELECT 
                            cnpj_basico,
                            IF(opcao_simples = 1, 'Sim', 'Não') as Simples_Nacional
                        FROM `basedosdados.br_me_cnpj.simples`
                        WHERE cnpj_basico IN ({lista_sql})
                    """

                    with st.spinner("Consultando nuvem..."):
                        query_job = client.query(query)
                        df_nuvem = query_job.to_dataframe()

                        df_base = pd.DataFrame({'CNPJ': identificadores})
                        df_base['cnpj_basico'] = df_base['CNPJ'].str[:8]
                        df_nuvem['cnpj_basico'] = df_nuvem['cnpj_basico'].astype(str).str.zfill(8)

                        df_final = pd.merge(df_base, df_nuvem, on='cnpj_basico', how='left')
                        df_final['OBS'] = ""
                        mask_not_found = df_final['Simples_Nacional'].isna()
                        df_final.loc[mask_not_found, 'OBS'] = "Não encontrado"
                        df_final['Simples_Nacional'] = df_final['Simples_Nacional'].fillna('Não')
                        df_final.rename(columns={'Simples_Nacional': 'Simples Nacional'}, inplace=True)
                        df_final['CNPJ'] = df_final['CNPJ'].apply(formatar_cnpj)
                        df_final.drop(columns=['cnpj_basico'], inplace=True)

                        tempo_total = time.time() - start_time
                        st.markdown(
                            f"<div class='time-card'><b>Tempo de Resposta:</b><br>{round(tempo_total, 2)} segundos</div>",
                            unsafe_allow_html=True)
                        st.dataframe(df_final, use_container_width=True, hide_index=True)

                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_final.to_excel(writer, index=False, sheet_name='SimplesNacional')

                        with placeholder_btn_download:
                            st.download_button(label="Gerar Excel (xlsx)", data=output.getvalue(),
                                               file_name=f"simples_{datetime.now().strftime('%H%M%S')}.xlsx",
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"Erro no BigQuery: {e}")


def aba_brasilapi_individual():
    st.title("Consulta Individual (BrasilAPI)")
    st.caption("Verificação detalhada de documento único")
    st.divider()

    col_input, _ = st.columns([1, 1])
    with col_input:
        cnpj_input = st.text_input("Digite o documento:", placeholder="00.000.000/0000-00")
        btn_individual = st.button("Consultar BrasilAPI")

    if btn_individual:
        cnpj_limpo = re.sub(r"\D", "", cnpj_input)
        if len(cnpj_limpo) != 14:
            st.error("Insira 14 dígitos.")
        else:
            with st.spinner("Buscando..."):
                res = consultar_brasilapi(cnpj_limpo)
                if res:
                    st.markdown(f"""
                    <div class="card-individual">
                        <p style="color: #666; margin-bottom: 5px;">Razão Social</p>
                        <h3 style="margin-top: 0; color: #1f4e79;">{res['Razão Social']}</h3>
                        <hr>
                        <p><b>CNPJ:</b> {res['CNPJ']}</p>
                        <p><b>Status Simples:</b> {res['Status']}</p>
                        <p><b>Data Opção:</b> {res['Data Opção'] or 'N/A'}</p>
                        <p><b>Data Exclusão:</b> {res['Data Exclusão'] or 'N/A'}</p>
                        <p><b>OBS:</b> {res['OBS']}</p>
                    </div>
                    """, unsafe_allow_html=True)


# --- MAIN ---

def main():
    st.sidebar.title("Navegação")
    opcao = st.sidebar.selectbox("Selecione a modalidade:", ["Lote (Simples Nacional)", "Individual (Detalhado)"])

    if opcao == "Lote (Simples Nacional)":
        aba_simples_nacional_lote()
    else:
        aba_brasilapi_individual()


if __name__ == "__main__":
    main()