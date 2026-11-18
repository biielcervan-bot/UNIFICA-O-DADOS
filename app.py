import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Consolidador Diário por Agente - CPFL", layout="wide")

st.title("📊 Consolidador Diário de Produtividade e Horários (P e R)")
st.markdown(
    "Unifica as Unidades de Leitura por **Agente / Dia** considerando exclusivamente "
    "os registros de campo (**`P` e `R`**), mantendo a coluna **`IND_TIPO`** unificada (`P`, `R` ou `P / R`), "
    "somando os indicadores e calculando a **1ª Ação**, **Última Ação**, **Tempo em Campo** e **Média de Leituras por Hora**."
)

def formatar_timedelta(td):
    if pd.isna(td):
        return ""
    total_minutos = int(td.total_seconds() // 60)
    horas = total_minutos // 60
    minutos = total_minutos % 60
    return f"{horas:02d}:{minutos:02d}"

def processar_arquivos(arquivos):
    lista_dfs = []
    for arquivo in arquivos:
        if arquivo.name.endswith(".csv"):
            df_temp = pd.read_csv(arquivo, sep=None, engine="python")
        else:
            df_temp = pd.read_excel(arquivo)
        lista_dfs.append(df_temp)

    df = pd.concat(lista_dfs, ignore_index=True)

    # Padronização e filtro exclusivo para P e R (ignora C por completo)
    df["IND_TIPO"] = df["IND_TIPO"].astype(str).str.strip().str.upper()
    df["Data Realização"] = df["Data Realização"].astype(str).str.strip()
    df_campo = df[df["IND_TIPO"].isin(["P", "R"])].copy()

    # Converte HH:MM para timedelta
    df_campo["1a_td"] = pd.to_timedelta(df_campo["1ª Ação"].astype(str).str.strip() + ":00", errors="coerce")
    df_campo["ult_td"] = pd.to_timedelta(df_campo["Última Ação"].astype(str).str.strip() + ":00", errors="coerce")

    # Agrupamento: 1 linha por Agente por Dia
    grupo_cols = ["Data Realização", "Base Operacional", "Código Agente", "Nome Agente"]

    df_final = (
        df_campo.groupby(grupo_cols, as_index=False)
        .agg(
            Município=("Município", lambda x: " / ".join(sorted(x.dropna().astype(str).unique()))),
            Lote=("Lote", lambda x: " / ".join(sorted(x.dropna().astype(str).unique()))),
            Localização=("Localização", lambda x: " / ".join(sorted(x.dropna().astype(str).unique()))),
            IND_TIPO=("IND_TIPO", lambda x: " / ".join(sorted(x.dropna().astype(str).unique()))),
            Qtd_ULs=("Unidade de Leitura", "nunique"),
            Total_Registros=("Total Registros", "sum"),
            Limpos_Sucesso=("Limpos / Sucesso", "sum"),
            Impedimentos_G1=("Impedimentos G1", "sum"),
            Impedimentos_G2=("Impedimentos G2", "sum"),
            Total_Impedimentos=("Total Impedimentos", "sum"),
            primeira_td=("1a_td", "min"),
            ultima_td=("ult_td", "max"),
        )
    )

    # Cálculos de tempo e produtividade por hora
    df_final["duracao_td"] = df_final["ultima_td"] - df_final["primeira_td"]
    df_final["Horas Decimais"] = (df_final["duracao_td"].dt.total_seconds() / 3600.0).round(2)

    df_final["Média Leituras/Hora"] = df_final.apply(
        lambda r: round(r["Total_Registros"] / (r["duracao_td"].total_seconds() / 3600.0), 2)
        if pd.notna(r["duracao_td"]) and r["duracao_td"].total_seconds() > 0
        else 0.0,
        axis=1,
    )

    df_final["1ª Ação"] = df_final["primeira_td"].apply(formatar_timedelta)
    df_final["Última Ação"] = df_final["ultima_td"].apply(formatar_timedelta)
    df_final["Tempo Total Campo"] = df_final["duracao_td"].apply(formatar_timedelta)

    # Ordem final das colunas
    colunas_finais = [
        "Data Realização",
        "Base Operacional",
        "Município",
        "Lote",
        "Localização",
        "IND_TIPO",
        "Código Agente",
        "Nome Agente",
        "Qtd_ULs",
        "Total_Registros",
        "Limpos_Sucesso",
        "Impedimentos_G1",
        "Impedimentos_G2",
        "Total_Impedimentos",
        "1ª Ação",
        "Última Ação",
        "Tempo Total Campo",
        "Horas Decimais",
        "Média Leituras/Hora",
    ]

    df_final = df_final[colunas_finais].rename(
        columns={
            "Qtd_ULs": "Qtd Unidades de Leitura",
            "Total_Registros": "Total Registros",
            "Limpos_Sucesso": "Limpos / Sucesso",
            "Impedimentos_G1": "Impedimentos G1",
            "Impedimentos_G2": "Impedimentos G2",
            "Total_Impedimentos": "Total Impedimentos",
        }
    )

    df_final = df_final.sort_values(by=["Data Realização", "Base Operacional", "Nome Agente"]).reset_index(drop=True)
    return df_final

# Interface de Upload
arquivos_enviados = st.file_uploader(
    "Selecione uma ou mais planilhas (resumo_consolidado_operacional)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
)

if arquivos_enviados:
    try:
        df_resultado = processar_arquivos(arquivos_enviados)

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Agentes/Dia", len(df_resultado))
        col2.metric("Total Registros", f"{df_resultado['Total Registros'].sum():,}".replace(",", "."))
        col3.metric("Limpos / Sucesso", f"{df_resultado['Limpos / Sucesso'].sum():,}".replace(",", "."))
        col4.metric("Total Impedimentos", f"{df_resultado['Total Impedimentos'].sum():,}".replace(",", "."))
        media_geral = df_resultado[df_resultado["Média Leituras/Hora"] > 0]["Média Leituras/Hora"].mean()
        col5.metric("Média Geral Leit./Hora", f"{media_geral:.2f}".replace(".", ","))

        st.dataframe(df_resultado, use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_resultado.to_excel(writer, index=False, sheet_name="Consolidado Agente Dia")

        st.download_button(
            label="📥 Baixar Planilha Consolidada (.xlsx)",
            data=buffer.getvalue(),
            file_name="consolidado_agente_dia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Erro ao processar os arquivos: {e}")
