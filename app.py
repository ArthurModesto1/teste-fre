import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from etl import processar_dados_base
from export import gerar_excel_final, MODEL_OPTIONS, MODEL_VOLATILITY

st.set_page_config(page_title="Gerador FRE - Item 8", layout="wide", page_icon="📊")

# ==========================================
# DICIONÁRIO DE INSTRUÇÕES DA CVM E MEMÓRIA DE CÁLCULO
# ==========================================
INFO_CVM = {
    "8.5": """
**📖 Exigência CVM:** O emissor deve apresentar informações quantitativas em relação à remuneração baseada em ações sob a forma de **opções de compra de ações** reconhecida no resultado dos 3 últimos exercícios sociais e à prevista para o exercício corrente.
**🧮 Memória de Cálculo:**
* **Preço Médio:** Calculado pela média ponderada das opções na base (Inicial, Perdidas, Exercidas).
* **Diluição Potencial (%):** `(Quantidade Total de Opções em Aberto no Final do Exercício / Capital Social) * 100`.
    """,
    "8.6": """
**📖 Exigência CVM:** Informações quantitativas em relação a **cada outorga** de opções de compra de ações realizadas nos 3 últimos exercícios e a prevista para o exercício social corrente.
**🧮 Memória de Cálculo:**
* **Multiplicação (Ganho):** `Quantidade de opções outorgadas * Valor justo das opções na data da outorga`. Representa o valor total da remuneração a ser reconhecida para aquela outorga ao longo do período de carência (vesting).
    """,
    "8.7": """
**📖 Exigência CVM:** Refere-se às opções **em aberto** ao final do último exercício social, segregadas entre opções "ainda não exercíveis" e "exercíveis".
**🧮 Memória de Cálculo:**
* **Valor Justo Total:** `Quantidade de opções * Valor justo das opções no último dia do exercício social`.
    """,
    "8.8": """
**📖 Exigência CVM:** Refere-se exclusivamente às opções **exercidas** (remuneração baseada em ações) nos 3 últimos exercícios sociais.
**🧮 Memória de Cálculo:**
* **Multiplicação do Ganho Total:** `Total de ações relativas às opções exercidas * (Preço médio de mercado das ações - Preço médio ponderado de exercício)`.
    """,
    "8.9": """
**📖 Exigência CVM:** Refere-se aos planos de **ações restritas** (tradicionais ou fantasmas) a serem entregues diretamente aos beneficiários, reconhecidas no resultado dos 3 últimos anos e prevista para o ano corrente.
**🧮 Memória de Cálculo:**
* **Diluição Potencial (%):** `(Quantidade Final de Ações em Aberto / Capital Social) * 100`.
    """,
    "8.10": """
**📖 Exigência CVM:** Detalhamento quantitativo para **cada outorga de ações restritas** realizada nos 3 últimos exercícios sociais e a prevista para o exercício social corrente.
**🧮 Memória de Cálculo:**
* **Multiplicação:** `Quantidade de ações outorgadas * Valor justo das ações na data da outorga`. 
* **Nota:** O prazo máximo para entrega especifica a data limite na qual as ações outorgadas serão entregues de forma irrevogável.
    """,
    "8.11": """
**📖 Exigência CVM:** Detalha as ações **entregues (adquiridas)** relativas à remuneração baseada em ações (planos de ações restritas) nos 3 últimos exercícios sociais.
**🧮 Memória de Cálculo:**
* **Multiplicação (Efetivo Dispêndio):** `Total de ações adquiridas * (Preço médio de mercado - Preço médio de aquisição)`. Representa o ganho/dispêndio efetivo da companhia nas ações entregues.
* **Validação Estatutário:** Indica se o beneficiário já era Estatutário na data em que o exercício foi realizado.
    """,
    "8.12": """
**📖 Exigência CVM:** O emissor deve certificar que os investidores compreendam os dados dos itens 8.5 a 8.11, detalhando o modelo de precificação do valor justo (ex: Black-Scholes, Binomial).
**🧮 Detalhamento:** Devem ser informados dados quantificados como volatilidade esperada, taxa livre de risco, dividendos esperados e preço médio ponderado. Os valores apresentados aqui refletem os programas que estavam **em aberto no início do exercício** (Data de Outorga anterior a 01/01/2025).
    """
}

# ==========================================
# GERADOR DE TEMPLATE ANONIMIZADO
# ==========================================
def gerar_template_anonimizado():
    """Gera um arquivo Excel em memória com a estrutura exata exigida pelo ETL, mas com dados fictícios."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:

        # 1. Dados da outorga
        df_outorga = pd.DataFrame({
            'Nome': ['Administrador A', 'Administrador A', 'Administrador B'],
            'Órgão Administrativo': ['Diretoria Estatutária', 'Diretoria Estatutária', 'Conselho de Administração'],
            'Desligado?': ['Não', 'Não', 'Não'],
            'Tipo de Plano': ['Stock Options', 'Stock Options', 'Ações Restritas'],
            'Programa': ['SOP 2022', 'SOP 2023', 'RSU 2024'],
            'Data de Outorga': ['01/07/2022', '10/05/2023', '01/01/2024'],
            'Preço de Exercício na Outorga': [15.50, 18.00, 0.00],
            'Preço de Exercício Atual': [15.50, 18.00, 0.00],
            'Lote': [1, 1, 1],
            'Fase': ['Encerrado', 'Carência', 'Carência'],
            'Status': ['Exercido', 'Em aberto', 'Em aberto'],
            'Data da Carência': ['01/01/2025', '01/01/2026', '01/01/2027'],
            'Data de Expiração': ['01/01/2030', '01/01/2031', '01/01/2034'],
            'Outorgado (original)': [50000, 40000, 20000],
            'Fair Value na Outorga': [5.20, 6.80, 20.00],
            'Fair Value Atualizado': [6.00, 7.50, 22.00],
            'Proporção de Exercício Antecipado': ['N/A', 'N/A', 'N/A'],
            'Preço da Ação / Opção': [20.00, 20.00, 25.00],
            'Taxa de Juros Livre de Risco': [0.10, 0.11, 0.10],
            'Dividendos Esperados': [0.03, 0.03, 0.03],
            'Volatilidade': [0.35, 0.35, 0.35],
            'Model Options': [1, 1, ''],
            'ModelVolatility': [1, 1, '']
        })
        df_outorga.to_excel(writer, sheet_name='Dados da outorga', index=False)

        # 2. Histórico de movimentações (header na linha 2)
        cols_mov = ['Nome', 'Órgão Administrativo', 'Empresa', 'Data', 'Plano', 'Programa',
                    'Lote', 'Data de Outorga', 'Preço de Ação na Data de Outorga',
                    'Status', 'Demitido', 'Motivo da Demissão',
                    'Quantidade de Ações na Base Atual', 'Ações Liberadas  no Exercício',
                    'Preço da Ação', 'Preço de Exercício na Base Atual',
                    'Receita*', 'Ganho*', 'Tipo de Recebimento', 'Ganho na Moeda Local']
        df_mov_data = [
            ['Administrador A', 'Diretoria Estatutária', 'Empresa Exemplo',
             '16/04/2025', 'Plano 2021', 'SOP 2022', 1, '01/07/2022', None,
             'Exercido', False, '-', 50000, 50000, 20.00, None, 260000, 225000, 'Ações', 0],
        ]
        df_mov_excel = pd.DataFrame(columns=range(len(cols_mov)))
        df_mov_excel.loc[0] = ['Histórico de movimentação'] + [''] * (len(cols_mov) - 1)
        df_mov_excel.loc[1] = cols_mov
        for i, row in enumerate(df_mov_data):
            df_mov_excel.loc[i+2] = row
        df_mov_excel.to_excel(writer, sheet_name='Histórico de movimentações', index=False, header=False)

        # 3. Previsão outorga 2026 (formato real)
        df_prev = pd.DataFrame([
            [None, 'Caso haja algum novo programa planejado para o ano de 2026...', None],
            [None, 'Preencha as informações nas células da coluna "B"', None],
            [None, 'Obs: O levantamento abaixo não precisa ser exato, pois é uma previsão para 2026.', None],
            [None, None, None],
            [None, 'Novo Programa 1', None],
            [None, 'Nome do Programa:', 'PROGRAMA 2026'],
            [None, 'Tipo de programa a ser outorgado', 'RSU'],
            [None, 'Quantidade a ser outorgada para o Conselho de Administração:', 0],
            [None, 'Quantidade a ser outorgada para a Diretoria Estatutária:', 300000],
            [None, 'Quantidade a ser outorgada para o Conselho Fiscal:', 0],
            [None, 'Estimativa do Preço da ação na data da outorga:', 25.50],
        ])
        df_prev.to_excel(writer, sheet_name='Previsão outorga 2026', index=False, header=False)

        # 4. Membros
        df_membros = pd.DataFrame({
            'Orgão': ['Diretoria Estatutária', 'Conselho de Administração', 'Conselho Fiscal'],
            'NOME COMPLETO': ['Administrador A', 'Administrador B', 'Administrador C'],
            'DATA DE ENTRADA': ['01/01/2020', '01/01/2021', '01/01/2022'],
            'DATA DE SAÍDA': ['31/12/2999', '31/12/2999', '31/12/2026'],
        })
        df_membros.to_excel(writer, sheet_name='Membros', index=False)

    output.seek(0)
    return output

# ==========================================
# HELPER: FORMATAÇÃO PADRÃO CVM (Traços e Moedas)
# ==========================================
def fmt(val, tipo="num"):
    """Formata valores. Retorna '-' se for nulo ou zero."""
    try:
        if pd.isna(val) or val == "" or float(val) == 0:
            return " - "
        v = float(val)
        if tipo == "moeda": return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        if tipo == "perc":  return f"{v*100:.2f}%".replace(".", ",") if v <= 1 else f"{v:.2f}%".replace(".", ",")
        if tipo == "int":   return f"{int(v)}"
        if tipo == "float": return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(val) if pd.notna(val) and val != "" else " - "

# ==========================================
# MOTOR DE CONSOLIDAÇÃO DO RESUMO
# ==========================================
def gerar_resumo_cvm(dados_completos, chave, cap_social):
    """Constrói o DataFrame de resumo quebrado por anos com os rótulos exatos da CVM."""
    df = dados_completos.get(chave, pd.DataFrame())
    df_membros = dados_completos.get('membros', pd.DataFrame())

    if df.empty and chave != "8.12":
        return {}

    col_org_membros = 'Orgão' if 'Orgão' in df_membros.columns else (
        'Órgão' if 'Órgão' in df_membros.columns else 'Órgão Administrativo')
    orgaos = ['Conselho de Administração', 'Diretoria Estatutária', 'Conselho Fiscal', 'Total']

    resumos_por_ano = {}

    # ── Determina os anos presentes nos dados ─────────────────────────────────
    if chave in ["8.5", "8.9"]:
        # Status tem formato "Inicial 2025", "Final 2026", etc.
        anos_raw = df['Status'].dropna().str.extract(r'(\d{4})')[0].dropna().unique()
        anos = sorted(anos_raw)
    elif chave in ["8.8", "8.11"]:
        df = df.copy()
        df['Ano_Ref'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce').dt.year
        anos = sorted(df['Ano_Ref'].dropna().astype(int).astype(str).unique())
    elif chave in ["8.6", "8.10"]:
        df = df.copy()
        df['Ano_Ref'] = pd.to_datetime(df['Data Outorga'], errors='coerce').dt.year
        anos = sorted(df['Ano_Ref'].dropna().astype(int).astype(str).unique())
    else:
        anos = ["Último Exercício"]

    if not anos:
        anos = ["Último Exercício"]

    for ano in anos:
        # Filtra registros do ano
        if chave in ["8.5", "8.9"]:
            df_ano = df[df['Status'].str.contains(str(ano), na=False)].copy()
        elif chave in ["8.8", "8.11", "8.6", "8.10"]:
            df_ano = df[df['Ano_Ref'].astype(str) == str(ano)].copy()
        elif chave in ["8.7", "8.12"]:
            df_ano = df.copy()
        else:
            df_ano = df.copy()

        if df_ano.empty and chave != "8.12":
            continue

        lbl_ano = (f"{ano} (Projeção)" if str(ano) == "2026"
                   else f"Exercício {ano}" if str(ano).isdigit()
                   else str(ano))

        # ── 8.6 / 8.10: por programa ─────────────────────────────────────────
        if chave in ["8.6", "8.10"]:
            if 'Coluna_Relatorio' not in df_ano.columns:
                continue
            programas = [p for p in df_ano['Coluna_Relatorio'].unique() if pd.notnull(p)]
            if not programas:
                continue

            linhas = {
                "Nº total de membros": {}, "N° de membros remunerados": {}, "Data de outorga": {},
                "Quantidade outorgada": {}, "Prazo para opções/ações se tornarem exercíveis": {},
                "Prazo máximo": {}, "Prazo de restrição": {}, "Valor justo na data": {}, "Multiplicação (Ganho)": {}
            }

            for prog in programas:
                d_p = df_ano[df_ano['Coluna_Relatorio'] == prog]
                if d_p.empty:
                    continue
                org_prog = d_p['Órgão Administrativo'].iloc[0]

                linhas["Nº total de membros"][prog] = fmt(
                    len(df_membros[df_membros[col_org_membros].astype(str) == org_prog]), "int"
                )
                linhas["N° de membros remunerados"][prog] = fmt(len(d_p['Nome'].unique()), "int")
                linhas["Data de outorga"][prog] = (
                    d_p['Data Outorga'].iloc[0].strftime('%d/%m/%Y')
                    if pd.notnull(d_p['Data Outorga'].iloc[0]) else " - "
                )
                linhas["Quantidade outorgada"][prog] = fmt(d_p['Qtd_Outorgada'].sum(), "int")
                linhas["Prazo para opções/ações se tornarem exercíveis"][prog] = (
                    d_p['Data Carência'].iloc[0].strftime('%d/%m/%Y')
                    if pd.notnull(d_p['Data Carência'].iloc[0]) else " - "
                )
                linhas["Prazo máximo"][prog] = (
                    d_p['Data Expiração'].iloc[0].strftime('%d/%m/%Y')
                    if chave == '8.6' and 'Data Expiração' in d_p.columns and pd.notnull(d_p['Data Expiração'].iloc[0])
                    else " - "
                )
                linhas["Prazo de restrição"][prog] = "N/A"
                linhas["Valor justo na data"][prog] = fmt(d_p['Fair_Value'].mean(), "moeda")
                linhas["Multiplicação (Ganho)"][prog] = fmt(
                    (d_p['Qtd_Outorgada'] * d_p['Fair_Value']).sum(), "moeda"
                )

            resumos_por_ano[lbl_ano] = pd.DataFrame(linhas).T
            continue

        # ── 8.12: premissas de precificação ─────────────────────────────────
        if chave == "8.12":
            df_out = dados_completos.get('df_outorga', pd.DataFrame())
            df_out = df_out.copy()
            df_out['Data de Outorga'] = pd.to_datetime(df_out.get('Data de Outorga'), errors='coerce', dayfirst=True)
            df_25 = df_out[df_out['Data de Outorga'].dt.year < 2025].copy()

            if df_25.empty:
                resumos_por_ano["Exercício Atual"] = pd.DataFrame(
                    {"Aviso": ["Nenhum programa em aberto no início de 2025"]})
                continue

            df_25['Chave_Coluna'] = df_25['Programa'] + " (" + df_25['Órgão Administrativo'] + ")"
            col_progs = df_25['Chave_Coluna'].unique()

            linhas = {
                "Modelo de Precificação": {}, "Preço Médio Ponderado das Ações (R$)": {},
                "Preço de Exercício (R$)": {}, "Volatilidade Esperada (%)": {},
                "Prazo de vida da opção": {}, "Dividendos Esperados (%)": {},
                "Taxa de juros livre de riscos (%)": {}, "Exercício antecipado": {},
                "Determinação da volatilidade": {}, "Outras características": {}
            }

            for prog in col_progs:
                d_p = df_25[df_25['Chave_Coluna'] == prog].iloc[0]

                modelo_raw = pd.to_numeric(str(d_p.get('Model Options', '')).replace(',', '.'), errors='coerce')
                linhas["Modelo de Precificação"][prog] = (
                    MODEL_OPTIONS.get(int(modelo_raw), "A preencher") if pd.notnull(modelo_raw) else "A preencher"
                )
                linhas["Preço Médio Ponderado das Ações (R$)"][prog] = fmt(
                    pd.to_numeric(str(d_p.get('Preço da Ação / Opção', 0)).replace(',', '.'), errors='coerce'), "moeda"
                )
                linhas["Preço de Exercício (R$)"][prog] = fmt(
                    pd.to_numeric(str(d_p.get('Preço de Exercício na Outorga', 0)).replace(',', '.'), errors='coerce'), "moeda"
                )
                linhas["Volatilidade Esperada (%)"][prog] = fmt(
                    pd.to_numeric(str(d_p.get('Volatilidade', 0)).replace(',', '.'), errors='coerce'), "perc"
                )

                dout = d_p.get('Data de Outorga')
                dexp = pd.to_datetime(d_p.get('Data de Expiração'), errors='coerce')
                if pd.notnull(dout) and pd.notnull(dexp):
                    if dexp.year >= 9999:
                        dcar = pd.to_datetime(d_p.get('Data da Carência'), errors='coerce')
                        linhas["Prazo de vida da opção"][prog] = (
                            f"{(dcar - dout).days / 365.25:.1f} anos" if pd.notnull(dcar) else " - "
                        )
                    else:
                        linhas["Prazo de vida da opção"][prog] = f"{(dexp - dout).days / 365.25:.1f} anos"
                else:
                    linhas["Prazo de vida da opção"][prog] = " - "

                linhas["Dividendos Esperados (%)"][prog] = fmt(
                    pd.to_numeric(str(d_p.get('Dividendos Esperados', 0)).replace(',', '.'), errors='coerce'), "perc"
                )
                linhas["Taxa de juros livre de riscos (%)"][prog] = fmt(
                    pd.to_numeric(str(d_p.get('Taxa de Juros Livre de Risco', 0)).replace(',', '.'), errors='coerce'), "perc"
                )
                linhas["Exercício antecipado"][prog] = str(d_p.get('Proporção de Exercício Antecipado', ' - '))

                vol_raw = pd.to_numeric(str(d_p.get('ModelVolatility', '')).replace(',', '.'), errors='coerce')
                linhas["Determinação da volatilidade"][prog] = (
                    MODEL_VOLATILITY.get(int(vol_raw), "A preencher") if pd.notnull(vol_raw) else "A preencher"
                )
                linhas["Outras características"][prog] = "A preencher"

            resumos_por_ano["Exercício Atual"] = pd.DataFrame(linhas).T
            continue

        # ── 8.5 / 8.7 / 8.8 / 8.9 / 8.11: por órgão ───────────────────────
        resumo = {org: {} for org in orgaos}

        for org in orgaos:
            d = df_ano if org == 'Total' else df_ano[df_ano['Órgão Administrativo'] == org]

            # Total de membros — Pro_Rata correto por ano
            col_pr = 'Pro_Rata_2026' if str(ano) == '2026' else 'Pro_Rata_2025'
            if org == 'Total':
                n_total = df_membros[col_pr].sum() if col_pr in df_membros.columns else len(df_membros)
            else:
                mask = df_membros[col_org_membros].astype(str) == org
                n_total = df_membros.loc[mask, col_pr].sum() if col_pr in df_membros.columns else mask.sum()
            resumo[org]["Nº total de membros"] = fmt(n_total, "float")

            # N° de membros remunerados — somente quem tem registro no ano
            if chave == "8.5":
                flag_col = 'Tem_85_2026' if str(ano) == '2026' else 'Tem_85_2025'
                if flag_col in df_membros.columns:
                    if org == 'Total':
                        n_rem = (df_membros[col_pr] * df_membros[flag_col]).sum()
                    else:
                        mask = df_membros[col_org_membros].astype(str) == org
                        n_rem = (df_membros.loc[mask, col_pr] * df_membros.loc[mask, flag_col]).sum()
                else:
                    n_rem = len(d['Nome'].unique()) if not d.empty else 0
            elif chave == "8.7":
                flag_col = 'Tem_87'
                if flag_col in df_membros.columns:
                    if org == 'Total':
                        n_rem = (df_membros[col_pr] * df_membros[flag_col]).sum()
                    else:
                        mask = df_membros[col_org_membros].astype(str) == org
                        n_rem = (df_membros.loc[mask, col_pr] * df_membros.loc[mask, flag_col]).sum()
                else:
                    n_rem = len(d['Nome'].unique()) if not d.empty else 0
            elif chave == "8.9":
                flag_col = 'Tem_89'
                if flag_col in df_membros.columns:
                    if org == 'Total':
                        n_rem = (df_membros[col_pr] * df_membros[flag_col]).sum()
                    else:
                        mask = df_membros[col_org_membros].astype(str) == org
                        n_rem = (df_membros.loc[mask, col_pr] * df_membros.loc[mask, flag_col]).sum()
                else:
                    n_rem = len(d['Nome'].unique()) if not d.empty else 0
            else:
                n_rem = len(d['Nome'].unique()) if not d.empty else 0
            resumo[org]["N° de membros remunerados"] = fmt(n_rem, "float")

            if chave == "8.5":
                resumo[org]["Preço médio ponderado de exercício:"] = ""
                resumo[org]["  Em aberto no início do exercício"] = fmt(
                    d[d['Status'].str.contains(f'Inicial {ano}', na=False)]['Preço'].mean(), "moeda"
                )
                resumo[org]["  Perdidas e expiradas"] = fmt(
                    d[d['Status'].str.contains(f'Perdidas {ano}', na=False)]['Preço'].mean(), "moeda"
                )
                resumo[org]["  Exercidas durante o exercício"] = fmt(
                    d[d['Status'].str.contains(f'Exercidas {ano}', na=False)]['Preço'].mean(), "moeda"
                )
                qtd_fim = d[d['Status'].str.contains(f'Final {ano}', na=False)]['Qtd'].sum()
                resumo[org]["Diluição potencial em caso de exercício"] = fmt(
                    (qtd_fim / cap_social) if cap_social else 0, "perc"
                )

            elif chave == "8.7":
                d_nao = d[d['Status_Vesting'] == 'Não exercível']
                d_sim = d[d['Status_Vesting'] == 'Exercível']

                resumo[org]["Opções ainda não exercíveis"] = ""
                resumo[org]["  i. Quantidade"]              = fmt(d_nao['Qtd_Saldo'].sum(), "int")
                resumo[org]["  ii. Preço médio ponderado"]  = fmt(d_nao['Preço'].mean(), "moeda")
                resumo[org]["  iii. Valor justo"]           = fmt(d_nao['Fair_Value'].mean(), "moeda")

                resumo[org]["Opções exercíveis"] = ""
                resumo[org]["  i. Quantidade"]              = fmt(d_sim['Qtd_Saldo'].sum(), "int")
                resumo[org]["  ii. Preço médio ponderado"]  = fmt(d_sim['Preço'].mean(), "moeda")
                resumo[org]["  iii. Valor justo"]           = fmt(d_sim['Fair_Value'].mean(), "moeda")
                resumo[org]["Valor justo do TOTAL"]         = fmt((d['Qtd_Saldo'] * d['Fair_Value']).sum(), "moeda")

            elif chave == "8.8":
                resumo[org]["Número de ações"]          = fmt(d['Qtd'].sum(), "int")
                resumo[org]["Preço médio de exercício"] = fmt(d['Preço_Ex'].mean(), "moeda")
                resumo[org]["Preço médio de mercado"]   = fmt(d['Preço_Merc'].mean(), "moeda")
                resumo[org]["Ganho Total (Multiplicação)"] = fmt(
                    (d['Qtd'] * (d['Preço_Merc'] - d['Preço_Ex'])).sum(), "moeda"
                )

            elif chave == "8.9":
                qtd_fim = d[d['Status'].str.contains(f'Final {ano}', na=False)]['Qtd'].sum()
                resumo[org]["Número de ações"]    = fmt(qtd_fim, "int")
                resumo[org]["Diluição potencial"] = fmt((qtd_fim / cap_social) if cap_social else 0, "perc")

            elif chave == "8.11":
                resumo[org]["Número de ações"]           = fmt(d['Qtd'].sum(), "int")
                resumo[org]["Preço médio de aquisição"]  = fmt(d['Preço_Aquisicao'].mean(), "moeda")
                resumo[org]["Preço médio de mercado"]    = fmt(d['Preço_Mercado'].mean(), "moeda")
                resumo[org]["Ganho Total (Multiplicação)"] = fmt(
                    (d['Qtd'] * (d['Preço_Mercado'] - d['Preço_Aquisicao'])).sum(), "moeda"
                )
                if 'Era_Estatutario' in d.columns:
                    resumo[org]["Membros que eram Estatutários na data"] = fmt(
                        d[d['Era_Estatutario'] == True]['Nome'].nunique(), "int"
                    )

        df_resumo = pd.DataFrame(resumo)
        for col in df_resumo.columns:
            if all(val == " - " or val == "" for val in df_resumo[col]):
                df_resumo[col] = " - "

        resumos_por_ano[lbl_ano] = df_resumo

    return resumos_por_ano

# ==========================================
# BARRA LATERAL (CONFIGS E NAVEGAÇÃO)
# ==========================================
with st.sidebar:
    st.title("📊 Gerador FRE")

    st.header("1. Configurações")
    capital_social = st.number_input(
        "Capital Social (Ações)",
        min_value=0,
        value=0,
        step=1000,
        help="Informe o total de ações do capital social para cálculo de diluição."
    )
    if capital_social == 0:
        st.warning("⚠️ Informe o Capital Social para calcular a Diluição Potencial.")

    st.header("2. Base de Dados")
    arquivo_up = st.file_uploader("Suba a Base de dados.xlsx", type=["xlsx"])

    st.divider()
    st.header("3. Exportação")
    btn_exportar = st.empty()

# ==========================================
# ÁREA PRINCIPAL
# ==========================================
if arquivo_up is not None:
    xls = pd.ExcelFile(arquivo_up)
    abas_esperadas = ['Dados da outorga', 'Histórico de movimentações', 'Previsão outorga 2026', 'Membros']
    abas_faltantes = [aba for aba in abas_esperadas if aba not in xls.sheet_names]

    if abas_faltantes:
        st.error(f"⚠️ Erro: Faltam as abas: **{', '.join(abas_faltantes)}**.")
    else:
        if "dados_processados" not in st.session_state:
            with st.spinner("Lendo e processando base de dados..."):
                try:
                    st.session_state.dados_processados = processar_dados_base(arquivo_up, capital_social)
                    st.toast('Processamento concluído!', icon='✅')
                except Exception as e:
                    st.error(f"Erro ao processar a base: {e}")
                    st.stop()

        dados = st.session_state.dados_processados
        if "8.12" not in dados:
            dados["8.12"] = dados["df_outorga"]

        chaves_quadros = ["8.5", "8.6", "8.7", "8.8", "8.9", "8.10", "8.11", "8.12"]
        titulos_abas = [f"Item {k}" for k in chaves_quadros]

        abas_principais = st.tabs(titulos_abas)

        for aba_principal, quadro_ativo in zip(abas_principais, chaves_quadros):
            with aba_principal:
                st.markdown(f"### 📈 Resumo Consolidado Padrão CVM - Quadro {quadro_ativo}")

                with st.expander(f"📖 Entenda o Item {quadro_ativo}"):
                    st.markdown(INFO_CVM.get(quadro_ativo, "Instrução não encontrada."))

                resumo_placeholder = st.empty()
                st.markdown("---")
                st.markdown(f"### 📝 Tabela Base (Editável)")

                df_completo = dados[quadro_ativo]

                if df_completo.empty and quadro_ativo != "8.12":
                    st.info("Nenhum dado registrado para este quadro.")
                    resumo_placeholder.info("Sem dados para consolidar.")
                else:
                    # --- 1. ÁREA DE EDIÇÃO ---
                    if quadro_ativo in ["8.6", "8.10", "8.12"]:
                        dados[quadro_ativo] = st.data_editor(
                            df_completo,
                            key=f"ed_{quadro_ativo}",
                            num_rows="dynamic",
                            use_container_width=True,
                            hide_index=True
                        )
                    else:
                        orgaos_presentes = [str(org) for org in df_completo['Órgão Administrativo'].dropna().unique()
                                            if str(org).strip() not in ['nan', '']]
                        if not orgaos_presentes:
                            orgaos_presentes = ["Geral"]

                        abas_orgaos = st.tabs(orgaos_presentes)
                        df_partes_editadas = []

                        for i, org in enumerate(orgaos_presentes):
                            with abas_orgaos[i]:
                                df_fatia = (df_completo if org == "Geral"
                                            else df_completo[df_completo['Órgão Administrativo'] == org])
                                editado = st.data_editor(
                                    df_fatia, key=f"ed_{quadro_ativo}_{i}", num_rows="dynamic",
                                    use_container_width=True, hide_index=True
                                )
                                df_partes_editadas.append(editado)

                        dados[quadro_ativo] = pd.concat(df_partes_editadas, ignore_index=True)

                    # --- 2. GERAÇÃO DO RESUMO ---
                    dict_resumos_por_ano = gerar_resumo_cvm(dados, quadro_ativo, capital_social)

                    with resumo_placeholder.container():
                        if dict_resumos_por_ano:
                            anos_keys = list(dict_resumos_por_ano.keys())
                            if len(anos_keys) > 1:
                                abas_anos = st.tabs([str(k) for k in anos_keys])
                                for idx_ano, ano_key in enumerate(anos_keys):
                                    with abas_anos[idx_ano]:
                                        st.dataframe(dict_resumos_por_ano[ano_key], use_container_width=True)
                            else:
                                st.markdown(f"**{anos_keys[0]}**")
                                st.dataframe(dict_resumos_por_ano[anos_keys[0]], use_container_width=True)
                        else:
                            st.info("Altere os dados abaixo para gerar o resumo.")

        with btn_exportar.container():
            if capital_social == 0:
                st.warning("Defina o Capital Social antes de exportar.")
            elif st.button("Gerar Excel Auditável", type="primary", use_container_width=True):
                with st.spinner("Montando arquivo Excel..."):
                    arquivo_saida = gerar_excel_final(dados, capital_social)
                    st.success("Planilha Pronta!")
                    st.download_button(
                        label="⬇️ Baixar Arquivo", data=arquivo_saida,
                        file_name="FRE_Item_8_Completo_Auditavel.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
else:
    # ==========================================
    # LANDING PAGE (ANTES DO UPLOAD)
    # ==========================================
    st.markdown("<h1 style='text-align: center;'>📊 Gerador Automático do FRE - Item 8</h1>", unsafe_allow_html=True)
    st.markdown("---")

    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown(
            """
            ### O que é esta ferramenta?
            Este sistema foi projetado para automatizar, consolidar e auditar a geração dos quadros da **Seção 8 (Remuneração dos Administradores)** do Formulário de Referência, em estrita conformidade com as exigências da CVM. Ele elimina o trabalho manual e previne erros de consolidação de planilhas complexas.

            ### 🚀 Como funciona o processo?
            1. **Upload da Base de Dados:** Utilize a barra lateral para definir o Capital Social atual da empresa e suba a sua planilha contendo os dados brutos de outorgas, histórico de movimentações e listagem de membros.
            2. **Revisão e Edição:** O sistema processará os dados e separará as tabelas e memórias de cálculo em abas (8.5 a 8.12). Você poderá visualizar os resumos nos layouts exigidos e editar a base de dados em tempo real.
            3. **Exportação Auditável:** Baixe o arquivo de saída. Ele é um Excel completo contendo as fórmulas (como `SOMASES` e `SOMARPRODUTO`) amarradas à base de dados para garantir a rastreabilidade e a aprovação pela auditoria.
            """
        )

    with col2:
        st.info(
            """
            **📋 Quadros Atendidos:**
            * **8.5 e 8.9:** Posição, Outorgas e Cancelamentos.
            * **8.6 e 8.10:** Detalhamento de Outorgas.
            * **8.7:** Opções em Aberto.
            * **8.8 e 8.11:** Opções Exercidas e Ações Entregues.
            * **8.12:** Premissas de Precificação.
            """
        )
        st.warning("👈 **Para iniciar, acesse a barra lateral à esquerda, configure o Capital Social e anexe a Base de Dados.**")

        st.markdown("---")
        st.markdown("### 📥 Teste o Sistema")
        st.markdown("Ainda não tem a planilha preenchida? Baixe nosso modelo com dados anonimizados para testar a ferramenta.")

        template_file = gerar_template_anonimizado()
        st.download_button(
            label="⬇️ Baixar Template Anonimizado",
            data=template_file,
            file_name="Template_Premissas_FRE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
