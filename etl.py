import pandas as pd
import numpy as np

def processar_dados_base(arquivo_up, capital_social):
    """Lê, transforma as abas e gera os DataFrames de evidências."""
    xls = pd.ExcelFile(arquivo_up)

    df_outorga = pd.read_excel(xls, sheet_name='Dados da outorga')
    df_mov     = pd.read_excel(xls, sheet_name='Histórico de movimentações', header=1)
    df_prev    = pd.read_excel(xls, sheet_name='Previsão outorga 2026', header=None)
    df_membros = pd.read_excel(xls, sheet_name='Membros')

    # ==========================================
    # LIMPEZA E TRANSFORMAÇÃO GLOBAL
    # ==========================================
    df_mov.columns = df_mov.columns.str.replace('<br/>', ' ').str.replace('\n', ' ').str.strip()
    df_mov['Órgão Administrativo'] = df_mov['Órgão Administrativo'].replace(
        {'Diretoria': 'Diretoria Estatutária', 'Conselheiro': 'Conselho de Administração'})
    df_mov['Data'] = pd.to_datetime(df_mov['Data'], errors='coerce', dayfirst=True)
    df_mov = df_mov[df_mov['Data'].notna()].copy()
    df_mov['Ano'] = df_mov['Data'].dt.year.astype('Int64')

    # ─── Identificação das colunas chave do histórico ───────────────────────
    # Quantidade do evento: preferir "Ações Liberadas no Exercício" (qty exercida/entregue)
    col_qtd_mov = None
    for c in df_mov.columns:
        if 'Liberadas' in c and 'Exercício' in c:
            col_qtd_mov = c
            break
    if col_qtd_mov is None:
        for c in df_mov.columns:
            if 'Quantidade de Ações' in c:
                col_qtd_mov = c
                break
    if col_qtd_mov is None:
        raise ValueError("Coluna de quantidade de ações não encontrada no histórico.")

    # Preço da ação (mercado) no dia do evento
    col_preco_merc = next((c for c in df_mov.columns if 'Preço da Ação' in c), None)

    # Preço de exercício no evento (pode ser vazio/NaN para RSU)
    col_preco_ex = next((c for c in df_mov.columns if 'Preço de Exercício' in c), None)

    # ─── Normalização do df_outorga ─────────────────────────────────────────
    df_outorga = df_outorga.dropna(subset=['Programa']).copy()
    df_outorga['Órgão Administrativo'] = df_outorga['Órgão Administrativo'].replace(
        {'Diretoria': 'Diretoria Estatutária', 'Conselheiro': 'Conselho de Administração'})
    df_outorga['Tipo de Plano'] = df_outorga['Tipo de Plano'].fillna('')

    preco_col = 'Preço de Exercício na Outorga'
    df_outorga[preco_col] = pd.to_numeric(
        df_outorga[preco_col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

    # ─── SEPARAÇÃO SOP vs RSU ────────────────────────────────────────────────
    # Critério: Preço de Exercício na Outorga
    #   preço > 0 e < 1e8  → Stock Options (SOP)
    #   preço == 0 ou >= 1e8 (sentinel para RSU/PSU/phantom) → Ações Restritas
    df_sop   = df_outorga[(df_outorga[preco_col] > 0) & (df_outorga[preco_col] < 1e8)].copy()
    df_acoes = df_outorga[(df_outorga[preco_col] == 0) | (df_outorga[preco_col] >= 1e8)].copy()

    cols_num = ['Preço de Exercício Atual', 'Fair Value na Outorga',
                'Outorgado (original)', 'Fair Value Atualizado']
    for df_temp in [df_sop, df_acoes]:
        for col in cols_num:
            if col in df_temp.columns:
                df_temp[col] = pd.to_numeric(
                    df_temp[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df_temp['Data de Outorga']   = pd.to_datetime(df_temp.get('Data de Outorga'),   errors='coerce', dayfirst=True)
        df_temp['Data da Carência']  = pd.to_datetime(df_temp.get('Data da Carência'),  errors='coerce', dayfirst=True)
        df_temp['Data de Expiração'] = pd.to_datetime(df_temp.get('Data de Expiração'), errors='coerce', dayfirst=True)

    # ─── Membros ─────────────────────────────────────────────────────────────
    df_membros['DATA DE ENTRADA'] = pd.to_datetime(df_membros['DATA DE ENTRADA'], errors='coerce', dayfirst=True)
    df_membros['DATA DE SAÍDA']   = pd.to_datetime(df_membros['DATA DE SAÍDA'],   errors='coerce', dayfirst=True)
    df_membros['_nome_lower']     = df_membros['NOME COMPLETO'].str.lower().str.strip()
    df_mov['_nome_lower']         = df_mov['Nome'].str.lower().str.strip()
    df_sop['_nome_lower']         = df_sop['Nome'].str.lower().str.strip()
    df_acoes['_nome_lower']       = df_acoes['Nome'].str.lower().str.strip()

    def calcular_pro_rata(row, ano):
        inicio_ano  = pd.Timestamp(f'{ano}-01-01')
        fim_ano     = pd.Timestamp(f'{ano}-12-31')
        entrada     = row['DATA DE ENTRADA']
        saida       = row['DATA DE SAÍDA']
        inicio_real = max(entrada, inicio_ano) if pd.notnull(entrada) else inicio_ano
        fim_real    = min(saida,   fim_ano)    if pd.notnull(saida)   else fim_ano
        if fim_real < inicio_ano or inicio_real > fim_ano:
            return 0.0
        return max(0.0, (fim_real - inicio_real).days + 1) / 365.0

    df_membros['Pro_Rata_2025'] = df_membros.apply(lambda x: calcular_pro_rata(x, 2025), axis=1)
    df_membros['Pro_Rata_2026'] = df_membros.apply(lambda x: calcular_pro_rata(x, 2026), axis=1)

    # ==========================================
    # LEITURA DA PREVISÃO 2026
    # Formato: blocos "Novo Programa N" com campos
    # em coluna B (descrição) e coluna C (valor)
    # ==========================================
    def _ler_previsao_2026():
        mapa_orgao_texto = {
            'conselho administrativo':   'Conselho de Administração',
            'conselho de administração': 'Conselho de Administração',
            'diretoria estatutária':     'Diretoria Estatutária',
            'diretoria':                 'Diretoria Estatutária',
            'conselho fiscal':           'Conselho Fiscal',
        }

        programas = []
        prog_atual = None

        for _, row in df_prev.iterrows():
            # Extrair chave (col B = index 1) e valor (col C = index 2)
            n_cols = len(row)
            chave = str(row.iloc[1]).strip() if n_cols > 1 and pd.notnull(row.iloc[1]) else ''
            valor = str(row.iloc[2]).strip() if n_cols > 2 and pd.notnull(row.iloc[2]) else ''
            if chave == 'nan': chave = ''
            if valor == 'nan': valor = ''

            # Detecta início de novo bloco de programa pelo campo da coluna B
            # O cabeçalho "Novo Programa N" aparece na coluna B também
            col1 = str(row.iloc[1]).strip() if n_cols > 1 and pd.notnull(row.iloc[1]) else ''
            if col1 == 'nan': col1 = ''
            if col1.startswith('Novo Programa') and (not valor or valor == ''):
                prog_atual = {'nome': col1, 'tipo': '', 'preco': 0.0, 'qtds': {}}
                programas.append(prog_atual)
                continue

            if prog_atual is None or not chave:
                continue

            chave_l = chave.lower()
            valor_num = pd.to_numeric(str(valor).replace(',', '.'), errors='coerce')

            if 'nome do programa' in chave_l:
                prog_atual['nome'] = valor if valor else prog_atual['nome']

            elif 'tipo de programa' in chave_l:
                prog_atual['tipo'] = valor

            elif 'estimativa do preço' in chave_l or ('preço da ação' in chave_l and 'outorga' in chave_l):
                if pd.notnull(valor_num):
                    prog_atual['preco'] = float(valor_num)

            elif 'quantidade a ser outorgada' in chave_l and pd.notnull(valor_num) and valor_num > 0:
                orgao_encontrado = None
                for chave_map, orgao_nome in mapa_orgao_texto.items():
                    if chave_map in chave_l:
                        orgao_encontrado = orgao_nome
                        break
                if orgao_encontrado:
                    prog_atual['qtds'][orgao_encontrado] = float(valor_num)

        return programas

    programas_prev_2026 = _ler_previsao_2026()

    # ==========================================
    # EVIDÊNCIAS INICIALIZADAS
    # ==========================================
    evid_85, evid_86, evid_87, evid_88 = [], [], [], []
    evid_89, evid_810, evid_811 = [], [], []
    limite_2025 = pd.Timestamp('2025-12-31')

    programas_sop = set(df_sop['Programa'].unique())
    programas_rsu = set(df_acoes['Programa'].unique())

    def _to_num(val):
        return pd.to_numeric(str(val).replace(',', '.'), errors='coerce')

    def _sum_col(df_slice, col):
        if col is None or col not in df_slice.columns:
            return 0.0
        return df_slice[col].apply(_to_num).fillna(0).sum()

    # ==========================================
    # SOP — ITENS 8.5 / 8.6 / 8.7 / 8.8
    # ==========================================
    for _, lote in df_sop.iterrows():
        nome      = lote['Nome']
        nlow      = lote['_nome_lower']
        orgao     = lote['Órgão Administrativo']
        prog      = lote['Programa']
        lote_num  = lote.get('Lote', 1)
        qtd       = lote['Outorgado (original)']
        data_out  = lote['Data de Outorga']
        fv_out    = lote.get('Fair Value na Outorga', 0)
        fv_atu    = lote.get('Fair Value Atualizado', fv_out)

        preco_inicial = float(lote['Preço de Exercício na Outorga'])
        preco_atual   = float(_to_num(lote.get('Preço de Exercício Atual', preco_inicial)) or preco_inicial)

        movs = df_mov[
            (df_mov['_nome_lower'] == nlow) &
            (df_mov['Programa']    == prog) &
            (df_mov['Lote']        == lote_num)
        ]

        exercidas_pre = _sum_col(movs[(movs['Ano'] <= 2024) & movs['Status'].str.contains('Exercido', case=False, na=False)], col_qtd_mov)
        perdidas_pre  = _sum_col(movs[(movs['Ano'] <= 2024) & movs['Status'].str.contains('Cancelado|Prescrito|Abandonado', case=False, na=False)], col_qtd_mov)
        saldo_ini = qtd - exercidas_pre - perdidas_pre

        if saldo_ini <= 0:
            continue
        if pd.notnull(data_out) and data_out.year > 2025:
            continue

        evid_85.append([nome, orgao, prog, lote_num, preco_inicial, saldo_ini, 'Inicial 2025'])

        movs_ex25  = movs[(movs['Ano'] == 2025) & movs['Status'].str.contains('Exercido', case=False, na=False)]
        movs_p25   = movs[(movs['Ano'] == 2025) & movs['Status'].str.contains('Cancelado|Prescrito|Abandonado', case=False, na=False)]
        exercidas_25 = _sum_col(movs_ex25, col_qtd_mov)
        perdidas_25  = _sum_col(movs_p25, col_qtd_mov)

        if exercidas_25 > 0:
            if col_preco_ex and not movs_ex25.empty:
                preco_ex = movs_ex25[col_preco_ex].apply(_to_num).mean()
                preco_ex = float(preco_ex) if pd.notnull(preco_ex) and preco_ex < 1e8 else preco_inicial
            else:
                preco_ex = preco_inicial
            evid_85.append([nome, orgao, prog, lote_num, preco_ex, exercidas_25, 'Exercidas 2025'])

        if perdidas_25 > 0:
            evid_85.append([nome, orgao, prog, lote_num, preco_inicial, perdidas_25, 'Perdidas 2025'])

        saldo_fim = saldo_ini - exercidas_25 - perdidas_25

        if saldo_fim > 0:
            evid_85.append([nome, orgao, prog, lote_num, preco_inicial, saldo_fim, 'Final 2025'])
            evid_85.append([nome, orgao, prog, lote_num, preco_atual,   saldo_fim, 'Inicial 2026'])
            evid_85.append([nome, orgao, prog, lote_num, preco_atual,   saldo_fim, 'Final 2026'])

            st_v = ('Exercível'
                    if (pd.notnull(lote['Data da Carência']) and lote['Data da Carência'] <= limite_2025)
                    else 'Não exercível')
            evid_87.append([
                orgao, nome, prog, lote_num, st_v, saldo_fim, preco_inicial,
                fv_atu, data_out, lote['Data da Carência'], lote['Data de Expiração']
            ])

        if pd.notnull(data_out) and data_out.year == 2025 and qtd > 0:
            evid_86.append([
                f"{prog} ({orgao})", orgao, nome, prog, lote_num,
                data_out, lote['Data da Carência'], lote['Data de Expiração'],
                qtd, fv_out
            ])

    # ── Previsões 2026 (classificadas por tipo) ──────────────────────────────
    for prog_info in programas_prev_2026:
        nome_prog  = prog_info.get('nome', 'Novo Programa 2026')
        tipo_prog  = prog_info.get('tipo', '').lower()
        preco_prev = prog_info.get('preco', 0.0)

        is_sop_prev = 'stock option' in tipo_prog or 'sop' in tipo_prog
        # PSU, RSU, "PSU + RSU", etc → RSU
        is_rsu_prev = not is_sop_prev  # tudo que não for SOP explícito

        for orgao_prev, qtd_prev in prog_info.get('qtds', {}).items():
            if qtd_prev > 0:
                if is_sop_prev:
                    evid_85.extend([
                        [nome_prog, orgao_prev, 'Previsão 2026', None, preco_prev, qtd_prev, 'Novas 2026'],
                        [nome_prog, orgao_prev, 'Previsão 2026', None, preco_prev, qtd_prev, 'Final 2026'],
                    ])
                    evid_86.append([
                        f"{nome_prog} ({orgao_prev})", orgao_prev, 'Previsão', 'Previsão 2026', None,
                        pd.NaT, pd.NaT, pd.NaT, qtd_prev, preco_prev
                    ])
                else:
                    evid_89.extend([
                        [nome_prog, orgao_prev, 'Previsão 2026', None, preco_prev, qtd_prev, 'Novas 2026'],
                        [nome_prog, orgao_prev, 'Previsão 2026', None, preco_prev, qtd_prev, 'Final 2026'],
                    ])
                    evid_810.append([
                        f"{nome_prog} ({orgao_prev})", orgao_prev, 'Previsão', 'Previsão 2026', None,
                        pd.NaT, pd.NaT, qtd_prev, preco_prev
                    ])

    # 8.8 — exercidas 2025 (SOP) por lote
    df_ex_25_sop = df_mov[
        (df_mov['Ano'] == 2025) &
        (df_mov['Status'].str.contains('Exercido', case=False, na=False)) &
        (df_mov['Programa'].isin(programas_sop))
    ].copy()

    for _, row in df_ex_25_sop.iterrows():
        qtd_ex = _to_num(row[col_qtd_mov])
        if pd.notnull(qtd_ex) and qtd_ex > 0:
            p_ex   = _to_num(row[col_preco_ex])   if col_preco_ex   else np.nan
            p_merc = _to_num(row[col_preco_merc]) if col_preco_merc else np.nan
            evid_88.append([
                row['Órgão Administrativo'], row['Nome'], row['Programa'], row.get('Lote', 1),
                row['Data'].strftime('%d/%m/%Y'), qtd_ex, p_ex, p_merc
            ])

    # ==========================================
    # RSU — ITENS 8.9 / 8.10 / 8.11
    # ==========================================
    termos_entregue = 'Exercido|Liberado|Entregue|Resgatado'
    termos_perdida  = 'Cancelado|Prescrito|Abandonado'

    for _, lote in df_acoes.iterrows():
        nome      = lote['Nome']
        nlow      = lote['_nome_lower']
        orgao     = lote['Órgão Administrativo']
        prog      = lote['Programa']
        lote_num  = lote.get('Lote', 1)
        qtd       = lote['Outorgado (original)']
        data_out  = lote['Data de Outorga']
        fv_out    = lote.get('Fair Value na Outorga', 0)
        fv_atu    = lote.get('Fair Value Atualizado', fv_out)
        ano_out   = data_out.year if pd.notnull(data_out) else 0

        movs = df_mov[
            (df_mov['_nome_lower'] == nlow) &
            (df_mov['Programa']    == prog) &
            (df_mov['Lote']        == lote_num)
        ]

        if ano_out == 2025:
            saldo_ini     = 0
            outorgadas_25 = qtd
        elif ano_out > 2025:
            continue
        else:
            outorgadas_25 = 0
            entregues_pre = _sum_col(movs[(movs['Ano'] <= 2024) & movs['Status'].str.contains(termos_entregue, case=False, na=False)], col_qtd_mov)
            perdidas_pre  = _sum_col(movs[(movs['Ano'] <= 2024) & movs['Status'].str.contains(termos_perdida,  case=False, na=False)], col_qtd_mov)
            saldo_ini = qtd - entregues_pre - perdidas_pre
            if saldo_ini <= 0:
                continue

        entregues_25 = _sum_col(movs[(movs['Ano'] == 2025) & movs['Status'].str.contains(termos_entregue, case=False, na=False)], col_qtd_mov)
        perdidas_25  = _sum_col(movs[(movs['Ano'] == 2025) & movs['Status'].str.contains(termos_perdida,  case=False, na=False)], col_qtd_mov)
        saldo_fim    = saldo_ini + outorgadas_25 - entregues_25 - perdidas_25

        if saldo_ini     > 0: evid_89.append([nome, orgao, prog, lote_num, fv_out, saldo_ini,     'Inicial 2025'])
        if outorgadas_25 > 0: evid_89.append([nome, orgao, prog, lote_num, fv_out, outorgadas_25, 'Outorgadas 2025'])
        if entregues_25  > 0: evid_89.append([nome, orgao, prog, lote_num, fv_out, entregues_25,  'Entregues 2025'])
        if perdidas_25   > 0: evid_89.append([nome, orgao, prog, lote_num, fv_out, perdidas_25,   'Perdidas 2025'])
        if saldo_fim     > 0: evid_89.append([nome, orgao, prog, lote_num, fv_atu, saldo_fim,     'Final 2025'])

        if ano_out == 2025 and qtd > 0:
            evid_810.append([
                f"{prog} ({orgao})", orgao, nome, prog, lote_num,
                data_out, lote['Data da Carência'], qtd, fv_out
            ])

    # 8.11 — ações entregues 2025 (RSU)
    df_ent_25 = df_mov[
        (df_mov['Ano'] == 2025) &
        (df_mov['Status'].str.contains(termos_entregue, case=False, na=False)) &
        (df_mov['Programa'].isin(programas_rsu))
    ].copy()

    for _, row_ent in df_ent_25.iterrows():
        qtd_ent = _to_num(row_ent[col_qtd_mov])
        if pd.notnull(qtd_ent) and qtd_ent > 0:
            p_aq   = _to_num(row_ent[col_preco_ex])   if col_preco_ex   else np.nan
            p_merc = _to_num(row_ent[col_preco_merc]) if col_preco_merc else np.nan
            data_mov = row_ent['Data']

            nlow_ent    = str(row_ent['Nome']).lower().strip()
            membro_info = df_membros[df_membros['_nome_lower'] == nlow_ent]
            era_estatutario = False
            if not membro_info.empty:
                entrada = membro_info.iloc[0]['DATA DE ENTRADA']
                saida   = membro_info.iloc[0]['DATA DE SAÍDA']
                era_estatutario = (
                    pd.notnull(entrada) and entrada <= data_mov and
                    (pd.isnull(saida) or saida >= data_mov)
                )

            evid_811.append([
                row_ent['Órgão Administrativo'], row_ent['Nome'],
                row_ent['Programa'], row_ent.get('Lote', 1),
                data_mov.strftime('%d/%m/%Y'), qtd_ent, p_aq, p_merc, era_estatutario
            ])

    # ==========================================
    # DATAFRAMES FINAIS
    # ==========================================
    df_evid_85  = pd.DataFrame(evid_85,  columns=['Nome', 'Órgão Administrativo', 'Programa', 'Lote', 'Preço', 'Qtd', 'Status'])
    df_evid_86  = pd.DataFrame(evid_86,  columns=['Coluna_Relatorio', 'Órgão Administrativo', 'Nome', 'Programa', 'Lote', 'Data Outorga', 'Data Carência', 'Data Expiração', 'Qtd_Outorgada', 'Fair_Value'])
    df_evid_87  = pd.DataFrame(evid_87,  columns=['Órgão Administrativo', 'Nome', 'Programa', 'Lote', 'Status_Vesting', 'Qtd_Saldo', 'Preço', 'Fair_Value', 'Data_Outorga', 'Data_Carência', 'Data_Expiração'])
    df_evid_88  = pd.DataFrame(evid_88,  columns=['Órgão Administrativo', 'Nome', 'Programa', 'Lote', 'Data', 'Qtd', 'Preço_Ex', 'Preço_Merc'])
    df_evid_89  = pd.DataFrame(evid_89,  columns=['Nome', 'Órgão Administrativo', 'Programa', 'Lote', 'Fair_Value', 'Qtd', 'Status'])
    df_evid_810 = pd.DataFrame(evid_810, columns=['Coluna_Relatorio', 'Órgão Administrativo', 'Nome', 'Programa', 'Lote', 'Data Outorga', 'Data Carência', 'Qtd_Outorgada', 'Fair_Value'])
    df_evid_811 = pd.DataFrame(evid_811, columns=['Órgão Administrativo', 'Nome', 'Programa', 'Lote', 'Data', 'Qtd', 'Preço_Aquisicao', 'Preço_Mercado', 'Era_Estatutario'])
    df_evid_811 = df_evid_811.drop_duplicates(subset=['Nome', 'Programa', 'Lote', 'Data', 'Qtd'])

    # ==========================================
    # MARCAÇÃO DE MEMBROS
    # ==========================================
    def nomes_lower(df, col):
        return set(df[col].str.lower().str.strip().unique()) if not df.empty and col in df.columns else set()

    nomes_85_2025 = nomes_lower(df_evid_85[df_evid_85['Status'].str.contains('2025', na=False)], 'Nome')
    nomes_85_2026 = nomes_lower(df_evid_85[df_evid_85['Status'].str.contains('2026', na=False)], 'Nome')

    df_membros['Tem_85_2025'] = df_membros['_nome_lower'].isin(nomes_85_2025).astype(int)
    df_membros['Tem_85_2026'] = df_membros['_nome_lower'].isin(nomes_85_2026).astype(int)
    df_membros['Tem_86']      = df_membros['_nome_lower'].isin(nomes_lower(df_evid_86,  'Nome')).astype(int)
    df_membros['Tem_87']      = df_membros['_nome_lower'].isin(nomes_lower(df_evid_87,  'Nome')).astype(int)
    df_membros['Tem_88']      = df_membros['_nome_lower'].isin(nomes_lower(df_evid_88,  'Nome')).astype(int)
    df_membros['Tem_89']      = df_membros['_nome_lower'].isin(nomes_lower(df_evid_89,  'Nome')).astype(int)
    df_membros['Tem_810']     = df_membros['_nome_lower'].isin(nomes_lower(df_evid_810, 'Nome')).astype(int)
    df_membros['Tem_811']     = df_membros['_nome_lower'].isin(nomes_lower(df_evid_811, 'Nome')).astype(int)

    df_membros = df_membros.drop(columns=['_nome_lower'])

    return {
        'membros':    df_membros,
        'df_outorga': df_outorga,
        '8.5':  df_evid_85,
        '8.6':  df_evid_86,
        '8.7':  df_evid_87,
        '8.8':  df_evid_88,
        '8.9':  df_evid_89,
        '8.10': df_evid_810,
        '8.11': df_evid_811,
    }
