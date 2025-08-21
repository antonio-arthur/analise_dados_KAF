from os import path
from utils import getdtb
from pandas import read_excel
import pandas as pd

# Configuração Inicial de Diretório
data_dir = r'C:\Users\anton\Desktop\data_analise_dados'

# ======================
# 1. Leitura de Dados DTB
# ======================
file = path.join(data_dir, 'IBGE', 'DTB_2024', 'RELATORIO_DTB_BRASIL_2024_MUNICIPIOS.xls')
print("Lendo o arquivo:", file)
df_dtb = getdtb(file)

# ======================
# 2. Leitura de Dados IA (Projetos 2020-2025)
# ======================
file = path.join(data_dir, 'OUTROS', 'Projetos_de_Atuac807a771o_-_IA_-_2020_a_2025.xlsx')
anos = [2020, 2021, 2022, 2023, 2024, 2025]
dfs = []

for ano in anos:
    try:
        df = pd.read_excel(file, sheet_name=str(ano), skiprows=5)

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

df_final = pd.concat(dfs, ignore_index=True)

# Pivot com métricas
df_wide_valores = df_final.pivot_table(
    index=['ds_mun', 'sg_uf'],
    columns='ano',
    values=['nprojetos', 'ninstituicoes', 'nbeneficiados'],
    fill_value=0
)
df_wide_valores.columns = [f"{col}_{ano}" for col, ano in df_wide_valores.columns]
df_wide_valores = df_wide_valores.reset_index()

# ======================
# 3. Padronização de Nomes
# ======================
def formatar_nome(df, coluna, novo_nome="ds_formatada"):
    df[novo_nome] = (
        df[coluna]
        .str.upper()
        .str.replace(r"[-.!?'`()]", "", regex=True)
        .str.replace("MIXING CENTER", "", regex=False)
        .str.strip()
        .str.replace(" ", "", regex=False)
    )
    return df

df_dtb = formatar_nome(df_dtb, coluna="ds_mun")
data = formatar_nome(df_wide_valores, coluna="ds_mun")

# Mapeia UF para nome completo
data['ds_uf'] = data['sg_uf'].map({
    "PB": "Paraíba", 
    "PE": "Pernambuco",
    "MG": "Minas Gerais", 
    "SP": "São Paulo"
})

# ======================
# 4. Merge IA + DTB
# ======================
data_m = data.merge(
    df_dtb, how="inner", 
    on=["ds_formatada", "ds_uf"],
    suffixes=["_ia", ""], 
    indicator="tipo_merge"
)
print("Merge IA+DTB", data_m["tipo_merge"].value_counts())

# ======================
# 5. Leitura IDEB
# ======================
lista_ideb = [f'VL_OBSERVADO_{x}' for x in range(2005,2025,2)]
nomes_ideb = [f'ideb_{x}' for x in range(2005,2025,2)]
notas_mate = [f'VL_NOTA_MATEMATICA_{x}' for x in range(2005,2025,2)]
notas_portugues = [f'VL_NOTA_PORTUGUES_{x}' for x in range(2005,2025,2)]

ideb = read_excel(
    path.join(data_dir, 'IBGE', 'divulgacao_anos_iniciais_municipios_2023',
              'divulgacao_anos_iniciais_municipios_2023.xlsx'),
    skiprows=9,
    usecols=['CO_MUNICIPIO', 'REDE'] + notas_mate + notas_portugues + lista_ideb,
    na_values=['-', '--']
)

ideb.columns = ['id_mundv', 'rede'] + notas_mate + notas_portugues + nomes_ideb

# ======================
# 6. Merge Final (IA+DTB+IDEB)
# ======================
# ======================
# 6. Merge Final (IA+DTB+IDEB)
# ======================
data_final = data_m.merge(ideb, how='left')

# Filtrar apenas Paraíba
data_final_pb = data_final.query("sg_uf == 'PB'")

# Salvar em Excel só PB
saida_arquivo = path.join(data_dir, 'resultado_final_PB.xlsx')
data_final_pb.to_excel(saida_arquivo, index=False)

print(f"✅ Planilha da Paraíba salva em: {saida_arquivo}")