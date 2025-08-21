import basedosdados as bd
import pandas as pd

billing_id = 'primeiroprojeto-469222'

query = """
  SELECT
    dados.ano as ano,
    dados.id_municipio AS id_municipio,
    diretorio_id_municipio.nome AS id_municipio_nome,
    dados.localizacao as localizacao,
    dados.taxa_abandono_ef_anos_iniciais as taxa_abandono_ef_anos_iniciais
FROM `basedosdados.br_inep_indicadores_educacionais.municipio` AS dados
LEFT JOIN (SELECT DISTINCT id_municipio,nome  
           FROM `basedosdados.br_bd_diretorios_brasil.municipio`) AS diretorio_id_municipio
    ON dados.id_municipio = diretorio_id_municipio.id_municipio
"""

df = bd.read_sql(query=query, billing_project_id=billing_id)

# Renomear para manter consistência
df = df.rename(columns={'id_municipio': 'id_mundv'})

# Carregar planilha do Instituto Alpargatas
file = pd.read_excel(r'C:\Users\anton\Desktop\data_analise_dados\resultado_final_PB.xlsx')

# Selecionar colunas importantes
df_abandono = df[["id_mundv", "ano", "taxa_abandono_ef_anos_iniciais"]]

# Garantir mesmo tipo
file["id_mundv"] = file["id_mundv"].astype(str)
df_abandono["id_mundv"] = df_abandono["id_mundv"].astype(str)

# Agrupar para evitar duplicatas
df_abandono = df_abandono.groupby(["id_mundv", "ano"], as_index=False).mean()

# Pivot: cada ano vira coluna
df_abandono_wide = df_abandono.pivot(
    index="id_mundv", 
    columns="ano", 
    values="taxa_abandono_ef_anos_iniciais"
).reset_index()

# Renomear colunas
df_abandono_wide.columns = ["id_mundv"] + [f"taxa_abandono_{ano}" for ano in df_abandono_wide.columns[1:]]

# Merge final
df_final = pd.merge(file, df_abandono_wide, on="id_mundv", how="left")

print(df_final)

if "ds_rgi" in df_final.columns:
    df_final = df_final.drop(columns=["ds_rgi"])
# Caminho para salvar o arquivo final
caminho_saida = r'C:\Users\anton\Desktop\data_analise_dados\resultado_final_PB_com_abandono.xlsx'

# Salvar em Excel
df_final.to_excel(caminho_saida, index=False)

print(f"Arquivo salvo com sucesso em: {caminho_saida}")
