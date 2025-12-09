# app/utils.py
import pandas as pd
import os
from datetime import datetime

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, '..', 'data', 'DADOS CONSOLIDADOS.xlsx')


def limpar_numero(x):
    if pd.isna(x) or x == '' or x is None:
        return 0.0
    texto = str(x).replace('R$', '').replace(' ', '').strip()
    # Adiciona a substituição da vírgula por ponto para garantir float
    texto = texto.replace(',', '.')

    # Remove tudo que não for dígito ou ponto
    texto = ''.join(c for c in texto if c.isdigit() or c == '.')

    if not texto:
        return 0.0
    try:
        return float(texto)
    except ValueError:
        return 0.0


def get_absolute_file_url(relative_path):
    """
    Converte um caminho relativo do Flask (/static/...) para um URL absoluto 'file:///'
    para ser usado pelo wkhtmltopdf.
    """
    path_segment = relative_path.lstrip('/')
    full_path = os.path.join(BASE_DIR, path_segment)
    return 'file:///' + full_path.replace('\\', '/')


def carregar_fotos_melhorias():
    """
    Lê imagens 'antes' e 'depois' da pasta /static/fotos_melhorias.
    """
    fotos_dir = os.path.join(BASE_DIR, "static", "fotos_melhorias")

    if not os.path.exists(fotos_dir):
        return []

    arquivos = sorted(os.listdir(fotos_dir))
    registros = {}

    for file in arquivos:
        nome = file.lower()
        caminho = f"/static/fotos_melhorias/{file}"

        base = ''.join([c for c in nome if c.isdigit()]) or nome.replace("antes","").replace("depois","")

        if base not in registros:
            registros[base] = {"antes": "", "depois": "", "tipo": "registro"}

        if "antes" in nome:
            registros[base]["antes"] = caminho
        elif "depois" in nome:
            registros[base]["depois"] = caminho

    lista = []
    for base, item in registros.items():
        if item["antes"] and item["depois"]:
            tipo = "melhoria"
        else:
            tipo = "registro"

        titulo = f"Reg. {base}"

        lista.append({
            "antes": item["antes"],
            "depois": item["depois"],
            "tipo": tipo,
            "titulo": titulo
        })

    return lista


def ler_dados():
    if not os.path.exists(EXCEL_FILE):
        return {"erro": f"Arquivo não encontrado:<br><b>{EXCEL_FILE}</b><br>Coloque na pasta data/"}

    try:
        # Leitura das planilhas
        vendas = pd.read_excel(EXCEL_FILE, sheet_name="VENDAS_DIARIAS", header=None, dtype=str)
        ticket = pd.read_excel(EXCEL_FILE, sheet_name="TICKET_MEDIO", dtype=str)
        formas = pd.read_excel(EXCEL_FILE, sheet_name="FORMAS_PAGAMENTO", dtype=str)
        produtos = pd.read_excel(EXCEL_FILE, sheet_name="PRODUTOS_VENDIDOS", dtype=str)
        # CORREÇÃO: Pula 2 linhas de metadados
        contas_df = pd.read_excel(EXCEL_FILE, sheet_name="REGISTRO DE CONTAS", header=2, dtype=str)
        resumo = pd.read_excel(EXCEL_FILE, sheet_name="RESUMO GERAL", dtype=str)
        sangria = pd.read_excel(EXCEL_FILE, sheet_name="SANGRIA", dtype=str)
        despesas = pd.read_excel(EXCEL_FILE, sheet_name="DESPESAS EXTRAS", dtype=str)
        problemas = pd.read_excel(EXCEL_FILE, sheet_name="Problemas", dtype=str)

        faturamento_total = limpar_numero(vendas.iloc[1, 8])

        valores_diarios = [limpar_numero(vendas.iloc[1, i]) for i in range(7)]
        datas_diarias = ['01/12', '02/12', '03/12', '04/12', '05/12', '06/12', '07/12']
        melhor_dia_valor = max(valores_diarios)
        melhor_dia = datas_diarias[valores_diarios.index(melhor_dia_valor)]

        ticket['Pessoas Atendidas'] = pd.to_numeric(ticket['Pessoas Atendidas'], errors='coerce').fillna(0)
        clientes_total = int(ticket['Pessoas Atendidas'].sum())
        ticket_medio = round(faturamento_total / clientes_total, 2) if clientes_total else 0

        formas['Valor'] = formas['Valor Pago (R$)'].apply(limpar_numero)
        formas_map = {
            'Cartão Crédito': 'Crédito', 'Cartão Débito': 'Débito', 'Dinheiro': 'Dinheiro',
            'PIX': 'PIX', 'Vale Refeição': 'Vale', '_Delivery Online': 'Delivery'
        }
        formas['Tipo'] = formas['Forma de Pagamento'].map(formas_map).fillna('Outros')
        resumo_formas = formas.groupby('Tipo')['Valor'].sum().round(2)
        formas_data = [{"nome": k, "valor": float(v)} for k, v in resumo_formas.items()]

        produtos['Quantidade'] = pd.to_numeric(produtos['Quantidade'], errors='coerce').fillna(0)
        top10 = produtos.groupby('Produto')['Quantidade'].sum().sort_values(ascending=False).head(10)
        top_produtos = [(idx, int(qtd)) for idx, qtd in top10.items()]

        sangria_total = sangria['Valor R$'].apply(limpar_numero).sum()
        despesas_total = despesas['Valor (R$)'].apply(limpar_numero).sum()
        saidas_total = round(sangria_total + despesas_total, 2)
        saldo_caixa = round(faturamento_total - saidas_total, 2)

        # CONTAS A PAGAR
        contas = []
        for _, row in contas_df.iterrows():
            if pd.notna(row.get('ID')) and str(row['ID']).strip():

                # --- CORREÇÃO E FORMATAÇÃO DA DATA ---
                data_vencimento = str(row.get('DATA VENCIMENTO', ''))
                try:
                    # Converte de 'YYYY-MM-DD HH:MM:SS' para objeto datetime
                    data_obj = datetime.strptime(data_vencimento[:10], '%Y-%m-%d')
                    # Formata para o padrão brasileiro
                    data_formatada = data_obj.strftime('%d/%m/%Y')
                except ValueError:
                    data_formatada = data_vencimento
                # -------------------------------------

                contas.append({
                    "id": int(row['ID']),
                    "fornecedor": str(row.get('FORNECEDOR', '')),
                    "descricao": str(row.get('DESCRIÇÃO', '')),
                    "valor": limpar_numero(row.get('VALOR', 0)),
                    "vencimento": data_formatada,
                    "status": str(row.get('STATUS', '')).strip().upper()
                })

        a_vencer = sum(c['valor'] for c in contas if c['status'] == 'A VENCER')
        vencido = sum(c['valor'] for c in contas if c['status'] == 'VENCIDO')
        pago = sum(c['valor'] for c in contas if c['status'] == 'PAGO')

        fotos_melhorias = carregar_fotos_melhorias()

        return {
            'faturamento_total': round(faturamento_total, 2),
            'clientes_total': clientes_total,
            'ticket_medio': ticket_medio,
            'melhor_dia': melhor_dia,
            'melhor_dia_valor': round(melhor_dia_valor, 2),
            'datas_diarias': datas_diarias,
            'valores_diarios': [round(x, 2) for x in valores_diarios],
            'formas_pagamento': formas_data,
            'top_produtos': top_produtos,
            'sangria_total': round(sangria_total, 2),
            'despesas_total': round(despesas_total, 2),
            'saidas_total': saidas_total,
            'saldo_caixa': saldo_caixa,
            'a_vencer': round(a_vencer, 2),
            'vencido': round(vencido, 2),
            'pago': round(pago, 2),
            'contas': contas,
            'fotos_melhorias': fotos_melhorias,
            'periodo': '01 a 07 de dezembro de 2025',
            'data_hoje': datetime.now().strftime('%d/%m/%Y')
        }

    except Exception as e:
        import traceback
        return {"erro": f"Erro:<pre>{traceback.format_exc()}</pre>"}