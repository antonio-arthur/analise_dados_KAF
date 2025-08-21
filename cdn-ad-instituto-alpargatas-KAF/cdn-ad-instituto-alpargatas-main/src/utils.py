from pandas import read_excel

def getdtb(file: str):
    """
    Leitura de Dados de Divisão Geográfica do Brasil - Municípios.
    """
    data = read_excel(file, skiprows=6, 
                    usecols=['UF', 'Nome_UF', 'Nome Região Geográfica Imediata',
                            'Código Município Completo', 'Nome_Município'])
    data.columns = ['id_uf', 'ds_uf', 'ds_rgi', 
                    'id_mundv', 'ds_mun']
    
    # Garantir unicidade de municípios
    data = data.drop_duplicates(subset=['id_mundv'])

    print(f"INSPEÇÃO PRIMEIRAS LINHAS\n", "--"*10, "\n\n", data.head(5))
    print("INSPEÇÃO ULTIMAS LINHAS \n\n", data.tail(5))
    print("INFO \n\n", data.info())

    return data

# Defina o caminho do arquivo antes de chamar a função
arquivo = r'C:/Users/anton/Desktop/data_analise_dados/IBGE/DTB_2024/RELATORIO_DTB_BRASIL_2024_MUNICIPIOS.xls'

# Agora chame a função com o caminho
dados = getdtb(arquivo)

