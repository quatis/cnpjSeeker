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

# --- MOTORES ---

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
                    "CNPJ": cnpj, "Simples Nacional": status, "Razão Social": d.get("razao_social", "N/A"),
                    "Data Opção": d.get("data_opcao_pelo_simples", ""),
                    "Data Exclusão": d.get("data_exclusao_do_simples", ""), "OBS": " | ".join(obs_parts)
                }
            elif resp.status_code == 429:
                time.sleep(7)
                continue
            return {"CNPJ": cnpj, "Simples Nacional": "Não", "Razão Social": "Não encontrado", "OBS": f"Erro {resp.status_code}"}
        except:
            time.sleep(2)
    return {"CNPJ": cnpj, "Simples Nacional": "Erro", "Razão Social": "Falha Rede", "OBS": "Timeout"}

def formatar_cnpj(c):
    c = str(c).zfill(14)
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"

# --- ABAS ---

def aba_bigquery():
    st.title("BigQuery (Nuvem)")
    st.caption("Alta performance via Google Cloud")
    st.divider()
    col1, col2 = st.columns([1, 2])
    with col1:
        input_texto = st.text_area("Insira os CNPJs (BigQuery):", height=300)
        placeholder_btn = st.empty()
        btn = st.button("Executar BigQuery")
    with col2:
        if btn:
            cnpjs = list(set(re.findall(r"\d{14}", input_texto.replace(".", "").replace("/", "").replace("-", ""))))
            if not cnpjs: st.error("Nenhum CNPJ.")
            else:
                start = time.time()
                client = get_bq_client()
                lista_sql = ", ".join([f"'{c[:8]}'" for c in cnpjs])
                query = f"SELECT cnpj_basico, IF(opcao_simples = 1, 'Sim', 'Não') as Status FROM `basedosdados.br_me_cnpj.simples` WHERE cnpj_basico IN ({lista_sql})"
                with st.spinner("Consultando..."):
                    df_nuvem = client.query(query).to_dataframe()
                    df_base = pd.DataFrame({'CNPJ': cnpjs})
                    df_base['cnpj_basico'] = df_base['CNPJ'].str[:8]
                    df_nuvem['cnpj_basico'] = df_nuvem['cnpj_basico'].astype(str).str.zfill(8)
                    df_final = pd.merge(df_base, df_nuvem, on='cnpj_basico', how='left')
                    df_final['OBS'] = ""
                    df_final.loc[df_final['Status'].isna(), 'OBS'] = "Não encontrado"
                    df_final['Status'] = df_final['Status'].fillna('Não')
                    df_final.rename(columns={'Status': 'Simples Nacional'}, inplace=True)
                    df_final['CNPJ'] = df_final['CNPJ'].apply(formatar_cnpj)
                    df_final.drop(columns=['cnpj_basico'], inplace=True)
                    st.markdown(f"<div class='time-card'><b>Tempo Resposta:</b><br>{round(time.time()-start, 2)}s</div>", unsafe_allow_html=True)
                    st.dataframe(df_final, use_container_width=True, hide_index=True)
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer: df_final.to_excel(writer, index=False)
                    with placeholder_btn: st.download_button("Gerar Excel (xlsx)", output.getvalue(), f"bq_{datetime.now().strftime('%H%M%S')}.xlsx")

def aba_brasilapi_lote():
    st.title("BrasilAPI (Lote)")
    st.caption("Fallback detalhado com monitoramento vivo")
    st.divider()
    col1, col2 = st.columns([1, 2])
    with col1:
        input_texto = st.text_area("Insira os CNPJs (BrasilAPI):", height=300)
        delay_ref = 0.8
        placeholder_btn = st.empty()
        btn = st.button("Iniciar BrasilAPI")
    with col2:
        if btn:
            cnpjs = list(set(re.findall(r"\d{14}", input_texto.replace(".", "").replace("/", "").replace("-", ""))))
            if not cnpjs: st.error("Nenhum CNPJ.")
            else:
                total = len(cnpjs)
                col_t1, col_t2 = st.columns(2)
                with col_t1: container_est = st.empty()
                container_dec = col_t2.empty()
                barra = st.progress(0)
                container_log = st.empty()
                df_display = pd.DataFrame(columns=["CNPJ", "Simples Nacional", "Razão Social", "Data Opção", "Data Exclusão", "OBS"])
                tabela = st.dataframe(df_display, use_container_width=True, hide_index=True)
                start = time.time()
                for idx, cnpj in enumerate(cnpjs):
                    percent = int(((idx + 1) / total) * 100)
                    barra.progress((idx + 1) / total)
                    container_dec.markdown(f"<div class='time-card'><b>Tempo Decorrido:</b><br>{round(time.time()-start, 1)}s</div>", unsafe_allow_html=True)
                    res = consultar_brasilapi(cnpj)
                    df_display = pd.concat([df_display, pd.DataFrame([res])], ignore_index=True)
                    tabela.dataframe(df_display, use_container_width=True, hide_index=True)
                    container_log.markdown(f"<div class='status-box'><span class='status-label'>Lote:</span> <span class='status-value'>{idx+1}/{total}</span><br><span class='status-label'>Ativo:</span> <span class='status-value'>{cnpj}</span></div>", unsafe_allow_html=True)
                    time.sleep(delay_ref)
                container_log.empty()
                st.success("Concluído.")
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer: df_display.to_excel(writer, index=False)
                with placeholder_btn: st.download_button("Gerar Excel (xlsx)", output.getvalue(), f"api_{datetime.now().strftime('%H%M%S')}.xlsx")

def aba_individual():
    st.title("Consulta Individual")
    st.caption("Detalhes pontuais via BrasilAPI")
    st.divider()
    cnpj_in = st.text_input("Documento:")
    if st.button("Consultar Individual"):
        c_limpo = re.sub(r"\D", "", cnpj_in)
        if len(c_limpo) == 14:
            res = consultar_brasilapi(c_limpo)
            st.markdown(f"<div class='card-individual'><h3>{res['Razão Social']}</h3><hr><p><b>Status:</b> {res['Simples Nacional']}</p><p><b>OBS:</b> {res['OBS']}</p></div>", unsafe_allow_html=True)

def main():
    st.sidebar.title("Navegação")
    opcao = st.sidebar.selectbox("Modalidade:", ["BigQuery (Rápido)", "BrasilAPI (Detalhado/Lote)", "Individual"])
    if opcao == "BigQuery (Rápido)": aba_bigquery()
    elif opcao == "BrasilAPI (Detalhado/Lote)": aba_brasilapi_lote()
    else: aba_individual()

if __name__ == "__main__":
    main()