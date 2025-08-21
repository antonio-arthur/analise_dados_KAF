import pandas as pd
from os import path

# ======================
# 1. Configuração de caminhos e anos
# ======================
data_dir = r'C:\Users\anton\Desktop\data_analise_dados'
file = path.join(data_dir, 'OUTROS', 'Projetos_de_Atuac807a771o_-_IA_-_2020_a_2025.xlsx')
anos = [2020, 2021, 2022, 2023, 2024, 2025]

# ======================
# 2. Leitura de Dados IA (Projetos 2020-2025)
# ======================
dfs = []

for ano in anos:
    try:
        df = pd.read_excel(file, sheet_name=str(ano), skiprows=5)

        # Padroniza nomes das colunas de município e estado
        if 'CIDADES' in df.columns:
            df = df.rename(columns={'CIDADES': 'ds_mun'})
        else:
            df['ds_mun'] = None

        if 'ESTADO' in df.columns:
            df = df.rename(columns={'ESTADO': 'sg_uf'})
        elif 'UF' in df.columns:
            df = df.rename(columns={'UF': 'sg_uf'})
        else:
            df['sg_uf'] = None

        # Identifica colunas de projetos, instituições e beneficiados
        projeto_col = [c for c in df.columns if 'Projetos' in c][-1]
        instituicoes_col = [c for c in df.columns if 'Instituições' in c or 'Instituições.' in c][-1]
        beneficiados_col = [c for c in df.columns if 'Beneficiados' in c][-1]

        df = df.rename(columns={
            projeto_col: 'nprojetos',
            instituicoes_col: 'ninstituicoes',
            beneficiados_col: 'nbeneficiados'
        })

        df['ano'] = ano
        df = df[['ds_mun', 'sg_uf', 'nprojetos', 'ninstituicoes', 'nbeneficiados', 'ano']]
        dfs.append(df)

    except Exception as e:
        print(f"⚠️ Erro ao processar aba {ano}: {e}")

# Concatena todos os anos
df_final = pd.concat(dfs, ignore_index=True)

# ======================
# 3. Pivot para formato wide
# ======================
df_wide_valores = df_final.pivot_table(
    index=['ds_mun', 'sg_uf'],
    columns='ano',
    values=['nprojetos', 'ninstituicoes', 'nbeneficiados'],
    fill_value=0
)
df_wide_valores.columns = [f"{col}_{ano}" for col, ano in df_wide_valores.columns]
df_wide_valores = df_wide_valores.reset_index()

# ======================
# 4. Salva arquivos em Excel
# ======================
saida_wide = path.join(data_dir, 'df_wide_valoress.xlsx')
df_wide_valores.to_excel(saida_wide, index=False)

saida_long = path.join(data_dir, 'df_final.xlsx')
df_final.to_excel(saida_long, index=False)

print(f"✅ Arquivo wide salvo em: {saida_wide}")
print(f"✅ Arquivo longo salvo em: {saida_long}")
