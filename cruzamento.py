import os
import pandas as pd

def cruzar_dados(df_dados,
                 coluna_semelhante_dados,
                 coluna_dados_ceder,
                 df_receptor,
                 coluna_semelhante_receptor,
                 coluna_final_receptor):
    mapa_credor = (
        df_dados
        .drop_duplicates(coluna_semelhante_dados)
        .set_index(coluna_semelhante_dados)[coluna_dados_ceder]
    )

    df_receptor[coluna_final_receptor] = df_receptor[coluna_semelhante_receptor].map(mapa_credor)

def localizar_renomear(df, 
                       coluna_map,
                       condicao,
                       coluna_renomeada,
                       nomeacao):
    if type(condicao)==list:
        mask = (
            df[coluna_map].str.contains(
                condicao[0],
                case=False,
                na=False,
                regex=True
            )
            &
            df[coluna_map].str.contains(
                condicao[1],
                case=False,
                na=False,
                regex=True
            )
        )
    else:
        mask = df[coluna_map].str.contains(
            condicao,
            case=False,
            na=False,
            regex=True
        )

    df.loc[mask, coluna_renomeada] = nomeacao

def tornar_datetime(df, coluna):
    # Garantir datetime
    df[coluna] = pd.to_datetime(
        df[coluna],
        dayfirst=True,
        errors='coerce'
    )

def dados(diretorio):
    # Abre arquivos
    lista_arq=os.listdir(diretorio)
    
    for arq in lista_arq:
        if arq.startswith('Relatório_Lançamentos Contábeis'):
            df_lancamentos = pd.read_excel(f'{diretorio}/{arq}', header=3)
        if arq.startswith('Relatório_OB'):
            df_ob = pd.read_excel(f'{diretorio}/{arq}', header=3)
        if arq.startswith('Relatório_PDs'):
            df_pd = pd.read_excel(f'{diretorio}/{arq}', header=3)
        if arq.startswith('Prestadores de Serviços'):
            df_ps = pd.read_excel(f'{diretorio}/{arq}', header=3)
        if arq.startswith('OB _ Apenas data'):
            df_datas = pd.read_excel(f'{diretorio}/{arq}', header=3)
        if arq.startswith('Relatório NE'):
            df_ne = pd.read_excel(f'{diretorio}/{arq}', header=3)
        if arq.startswith('Relatório - NL'):
            df_nl = pd.read_excel(f'{diretorio}/{arq}', header=3)

    # Cruzar dados 
    cruzar_dados(df_lancamentos, 'Programação de Desembolso', 'Situação envio', df_pd, 'Ordem Bancária', 'Situação envio')  # Descobrir Situação envio
    cruzar_dados(df_lancamentos, 'Programação de Desembolso', 'Credor', df_pd, 'Item de PD', 'Credor')                      # Descobrir Credor da OB
    cruzar_dados(df_lancamentos, 'Programação de Desembolso', 'Situação envio', df_pd, 'Ordem Bancária', 'Situação envio')  # Descobrir Situação envio
    cruzar_dados(df_lancamentos, 'Programação de Desembolso', 'Data Emissão', df_pd, 'Ordem Bancária', 'Data Emissão')      # Descobrir Data Emissão OB
    cruzar_dados(df_lancamentos, 'Programação de Desembolso', 'Observação', df_pd, 'Item de PD', 'Observação')              # Descobrir Observação

    # Renomeação condicional
    localizar_renomear(df_pd,'Observação',r'grupo sbv empreendimentos e participa','Credor','GRUPO SBV')                    # Descobre PDs da SBV
    localizar_renomear(df_pd,'Observação',r'di[aá]ria','Credor','Diárias')                                                  # Descobre PDs de Diárias
    localizar_renomear(df_pd,'Observação',[r'multas|juros',r'iss|inss|pis|cofins|csll'],'Credor','Multas e Juros sobre retenções')                        # Descobre PDs de Multas e Juros
    localizar_renomear(df_pd,'Observação',r'taxa','Credor','Taxa da emissão da Guia do ISS')                                # Descobre PDs de Taxa da emissão                                      

    # Garantir datetime
    tornar_datetime(df_datas,'Data Retorno')
    tornar_datetime(df_pd,'Data Emissão')

    # Pegar a ÚLTIMA data de retorno por Ordem Bancária
    mapa_status = (
        df_datas
        .groupby('Ordem Bancária')['Data Retorno']
        .max()
    )

    # Aplicar no dataframe principal
    df_pd['Data Retorno'] = df_pd['Ordem Bancária'].map(mapa_status)

    # Normaliza nomes dos credores
    # Faz um dicionário para as empresas
    df_dict = pd.read_excel('config/De_para.xlsx')
    dicionario = dict(zip(df_dict['De'], df_dict['Para']))

    df_pd['Credor'] = df_pd['Credor'].replace(dicionario)

    # Verifica valor das GDs que serão abatidos nas PDs parcialmente pagas
    df_gd = df_lancamentos[df_lancamentos['Programação de Desembolso'].astype(str).str.lower().str.contains("gd")]
    df_gd['Credor'] = df_gd['Credor'].replace(dicionario)
    df_gd = df_gd[df_gd['Tipo Estorno'].astype(str).str.lower().str.contains("não")]
    gd_validas = df_gd[
        df_gd["2026"] > 0
    ].copy()

    mapa_gd = (
        gd_validas
        .groupby("Credor")["2026"]
        .sum()
    )

    mask_parcial = (
        df_pd["Situação envio"]
        .astype(str)
        .str.strip()
        .str.lower()
        .eq("parcialmente pago")
    )

    # Define 3 valores (Valor Original, Valor da GD e Valor Líquido)
    df_pd["Valor Original"] = df_pd["2026"]
    df_pd["Valor GD"] = 0.0

    df_pd.loc[mask_parcial, "Valor GD"] = (
        df_pd.loc[mask_parcial, "Credor"]
        .map(mapa_gd)
        .fillna(0)
    )

    df_pd["Valor Liquido"] = (
        df_pd["Valor Original"]
        - df_pd["Valor GD"]
    )

    # Garantir datetime
    tornar_datetime(df_pd,'Data Retorno')
    tornar_datetime(df_pd,'Data Emissão')

    # Máscara das linhas que devem ser ajustadas
    mask_irr = (
        df_pd["Tipo de Retenção"]
        .astype(str)
        .str.lower()
        .str.contains("irr", na=False)
    ) & (
        df_pd["Situação envio"]
        .astype(str)
        .str.strip()
        .str.lower()
        .eq("processado e pago")
    ) & (
        df_pd["Data Retorno"].isna()
    )

    mask_pagos_hoje = (
        ~(df_pd["Tipo de Retenção"]
        .astype(str)
        .str.lower()
        .str.contains("irr", na=False))
    ) & (
        df_pd["Situação envio"]
        .astype(str)
        .str.strip()
        .str.lower()
        .eq("processado e pago")
    ) & (
        df_pd["Data Retorno"].isna()
    )

    # Preencher Data Retorno
    # Data Emissão + 1 dia
    df_pd.loc[mask_irr, "Data Retorno"] = (
        df_pd.loc[mask_irr, "Data Emissão"] + pd.Timedelta(days=1)
    )

    # Data Hoje
    df_pd.loc[mask_pagos_hoje, "Data Retorno"] = (
        pd.Timestamp.today().normalize() + pd.Timedelta(days=0)
    )

    return df_pd, df_ne, df_nl, df_lancamentos