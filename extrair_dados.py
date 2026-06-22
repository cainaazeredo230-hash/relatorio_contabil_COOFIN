import pandas as pd

def relatorio(df_pd):
    df_dict = pd.read_excel('config/De_para.xlsx')
    dicionario = dict(zip(df_dict['De'], df_dict['Para']))

    # Hoje
    hoje = pd.Timestamp.today().normalize()

    # Últimos 7 dias
    dia_7 = hoje - pd.Timedelta(days=6)

    list_ultimos_7_dias=[
        hoje - pd.Timedelta(days=0),
        hoje - pd.Timedelta(days=1),
        hoje - pd.Timedelta(days=2),
        hoje - pd.Timedelta(days=3),
        hoje - pd.Timedelta(days=4),
        hoje - pd.Timedelta(days=5),
        hoje - pd.Timedelta(days=6)
    ]

    # Início do mês atual
    inicio_mes = hoje.replace(day=1)

    # =========================
    # 1. Últimos 7 dias
    # =========================

    ultimos_7_dias = df_pd[
        (df_pd["Data Retorno"] >= dia_7) &
        (df_pd["Data Retorno"] <= hoje)
    ]

    # =========================
    # 2. Mês atual (antes dos últimos 7 dias)
    # =========================

    mes_atual_antes_7 = df_pd[
        (df_pd["Data Retorno"] >= inicio_mes) &
        (df_pd["Data Retorno"] < dia_7)
    ]
    MENOS_7_DIAS_PP = mes_atual_antes_7[mes_atual_antes_7['Situação envio']=='Processado e Pago']
    MENOS_7_DIAS_PARC_P = mes_atual_antes_7[mes_atual_antes_7['Situação envio']=='Parcialmente Pago']

    # =========================
    # 3. Meses passados
    # =========================

    meses_passados = df_pd[
        df_pd["Data Retorno"] < inicio_mes
    ]
    MESES_PASSADOS_PP = meses_passados[meses_passados['Situação envio']=='Processado e Pago']
    MESES_PASSADOS_PARC_P = meses_passados[meses_passados['Situação envio']=='Parcialmente Pago']

    
    situacao_pd_list = ['Processado e Pago','Erro no Pagamento','Parcialmente Pago']
    status_ex_list = ['Disponível', 'Agendada']

    df_processar = df_pd[~(df_pd['Situação envio'].isin(situacao_pd_list))]
    df_processar_cp = df_processar[df_processar['Status execução'].isin(status_ex_list)]

    list_credor = df_pd['Credor'].unique()
    armazenar_dados=[]
    for credor in list_credor:
        # Programação de desembolso
        df_credor = df_pd[df_pd['Credor']==credor]
        df_credor_contabilizado = df_credor[~(df_credor['Status PD']=='2 - Anulado')]                                                   # Ignora anuladas
        df_credor_contabilizado_sem_erro = df_credor_contabilizado[~(df_credor_contabilizado['Situação envio']=='Erro no Pagamento')]   # Ignora Erros de PGTO
        
        df_executar = df_credor_contabilizado_sem_erro[df_credor_contabilizado_sem_erro['Status execução']=='Disponível']               # PDs que não foram executadas
        
        df_executada = df_credor_contabilizado_sem_erro[df_credor_contabilizado_sem_erro['Status execução']=='Paga']                    # REGISTRO DE ENVIO (RE) DAS ORDENS BANCÁRIAS (OB)
        
        df_enviado = df_executada[
            df_executada['Situação envio'].isin([
                'Enviado ao Banco',
                'Aguardando Envio',
                'Enviado Manualmente'
            ])
        ]
        df_pago = df_executada[df_executada['Situação envio']=='Processado e Pago']                                                     # OBs PROCESSADAS E PAGAS
        df_parc_pago = df_executada[df_executada['Situação envio']=='Parcialmente Pago']                                                # OBs PROCESSADAS E PAGAS

        PASSOU_CREDOR = MESES_PASSADOS_PP[MESES_PASSADOS_PP['Credor']==credor]
        PASSOU_CREDOR_PARC = MESES_PASSADOS_PARC_P[MESES_PASSADOS_PARC_P['Credor']==credor]
    
        MENOS_7_DIAS_CREDOR = MENOS_7_DIAS_PP[MENOS_7_DIAS_PP['Credor']==credor]
        MENOS_7_DIAS_CREDOR_PARC = MENOS_7_DIAS_PARC_P[MENOS_7_DIAS_PARC_P['Credor']==credor]

        armazenar_retorno={}
        for DIA_RETORNO in list_ultimos_7_dias:
            df_dia_retorno = df_credor_contabilizado_sem_erro[df_credor_contabilizado_sem_erro["Data Retorno"]==DIA_RETORNO]
            armazenar_retorno[DIA_RETORNO.strftime('%d/%m/%Y')] = sum(df_dia_retorno['Valor Liquido'])

        # Armazena dados na variável "armazenar_dados"
        armazenar_dados.append({
            'CREDOR':credor,
            'PRESTADORES DE SERVIÇOS QUE FALTAM EMITIR PD': 0, #sum(PD_EXECUTAR_ORC['Valor Liquido']),
            'RETENÇÕES QUE FALTAM EMITIR PD': 0, #sum(PD_EXECUTAR_RET['Valor Liquido']),
            'PROGRAMAÇÃO DE DESEMBOLSO':sum(df_credor_contabilizado_sem_erro['Valor Liquido']),
            'PDs QUE FALTAM SER EXECUTADAS':sum(df_executar['Valor Liquido']),
            'REGISTRO DE ENVIO (RE) DAS ORDENS BANCÁRIAS (OB)':sum(df_executada['Valor Liquido']),
            'OBs ENVIADAS AO BANCO A SEREM PROCESSADAS E PAGAS':sum(df_enviado['Valor Liquido']),
            'OBs PROCESSADAS E PAGAS':sum(df_pago['Valor Liquido'])+sum(df_parc_pago['Valor Liquido']),
            '>MES ATUAL': sum(PASSOU_CREDOR['Valor Liquido'])+sum(PASSOU_CREDOR_PARC['Valor Liquido']),
            '>7 DIAS': sum(MENOS_7_DIAS_CREDOR['Valor Liquido'])+sum(MENOS_7_DIAS_CREDOR_PARC['Valor Liquido']),
            list_ultimos_7_dias[6].strftime('%d/%m/%Y'): armazenar_retorno[list_ultimos_7_dias[6].strftime('%d/%m/%Y')],
            list_ultimos_7_dias[5].strftime('%d/%m/%Y'): armazenar_retorno[list_ultimos_7_dias[5].strftime('%d/%m/%Y')],
            list_ultimos_7_dias[4].strftime('%d/%m/%Y'): armazenar_retorno[list_ultimos_7_dias[4].strftime('%d/%m/%Y')],
            list_ultimos_7_dias[3].strftime('%d/%m/%Y'): armazenar_retorno[list_ultimos_7_dias[3].strftime('%d/%m/%Y')],
            list_ultimos_7_dias[2].strftime('%d/%m/%Y'): armazenar_retorno[list_ultimos_7_dias[2].strftime('%d/%m/%Y')],
            list_ultimos_7_dias[1].strftime('%d/%m/%Y'): armazenar_retorno[list_ultimos_7_dias[1].strftime('%d/%m/%Y')],
            list_ultimos_7_dias[0].strftime('%d/%m/%Y'): armazenar_retorno[list_ultimos_7_dias[0].strftime('%d/%m/%Y')],
            })
        
    df_geral = pd.DataFrame(armazenar_dados)
    lista_empresas = list(dicionario.values())[2:]

    df_filtrado = df_geral[df_geral['CREDOR'].isin(lista_empresas)]

    return df_geral, df_filtrado

def ne_nl(df_ne, df_nl, df_lancamentos):
    # Normaliza nomes dos credores

    # Faz um dicionário para as empresas
    df_dict = pd.read_excel('config/De_para.xlsx')
    dicionario = dict(zip(df_dict['De'], df_dict['Para']))

    df_ne['Credor'] = df_ne['Credor'].replace(dicionario)
    df_nl['Credor'] = df_nl['Credor'].replace(dicionario)

    amazena_empenhos = []
    for credor in df_ne['Credor'].unique():
        # Nota de Empenho - NE
        df_credor_ne = df_ne[df_ne['Credor']==credor]

        df_contabilizado = df_credor_ne[df_credor_ne['Status NE']=='Contabilizado']

        df_empenho = df_contabilizado[~(df_contabilizado['Tipo Alteração NE']=='Anulacão')]
        df_anulacao = df_contabilizado[df_contabilizado['Tipo Alteração NE']=='Anulacão']

        # Nota de Liquidação - NL
        df_credor_nl = df_nl[df_nl['Credor']==credor]

        df_contabilizado_nl = df_credor_nl[df_credor_nl['Status NL']=='Contabilizado']

        df_liquidacao = df_contabilizado_nl[~(df_contabilizado_nl['Tipo Alteração NL']=='Anulacão')]
        df_anulacao_nl = df_contabilizado_nl[df_contabilizado_nl['Tipo Alteração NL']=='Anulacão']

        # Verifica GD do credor
        # Verifica valor das GDs que serão abatidos nas PDs parcialmente pagas
        df_gd = df_lancamentos[df_lancamentos['Programação de Desembolso'].astype(str).str.lower().str.contains("gd")]
        df_gd['Credor'] = df_gd['Credor'].replace(dicionario)
        df_gd = df_gd[df_gd['Tipo Estorno'].astype(str).str.lower().str.contains("não")]
        gd_validas = df_gd[
            df_gd["2026"] > 0
        ].copy()

        gd_validas_credor = gd_validas[gd_validas['Credor']==credor]

        # Dados
        LIQUIDACAO = sum(df_liquidacao['2026'])-(sum(df_anulacao_nl['2026'])+sum(gd_validas_credor['2026']))
        EMPENHO = sum(df_empenho['2026'])-sum(df_anulacao['2026'])

        amazena_empenhos.append({
                                'Credor': credor,
                                'NE': EMPENHO,
                                'NL': LIQUIDACAO
                                })
        
    df_ne_nl = pd.DataFrame(amazena_empenhos)

    return df_ne_nl