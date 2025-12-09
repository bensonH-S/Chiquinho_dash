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
    texto = texto.replace(',', '.')
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
    Classifica como 'MELHORIA' (se tiver Antes e Depois) ou 'REGISTRO'.
    """
    fotos_dir = os.path.join(BASE_DIR, "static", "fotos_melhorias")

    if not os.path.exists(fotos_dir):
        return []

    arquivos = sorted(os.listdir(fotos_dir))
    registros = {}

    for file in arquivos:
        nome = file.lower()
        caminho = f"/static/fotos_melhorias/{file}"

        base = nome.replace(".jpg", "").replace(".jpeg", "").replace(".png", "").replace("antes", "").replace("depois", "").strip('-')

        if base not in registros:
            registros[base] = {"antes": "", "depois": "", "tipo": "registro", "titulo_base": base.replace("-", " ").title()}

        if "antes" in nome:
            registros[base]["antes"] = caminho
        elif "depois" in nome:
            registros[base]["depois"] = caminho
        elif not registros[base]["depois"]:
             registros[base]["depois"] = caminho

    lista = []
    for base, item in registros.items():
        if item["antes"] and item["depois"] and item["antes"] != item["depois"]:
            tipo = "MELHORIA"
            titulo_final = f"Melhoria: {item['titulo_base']}"
        elif item["depois"]:
            tipo = "REGISTRO"
            titulo_final = f"Registro: {item['titulo_base']}"
        else:
            continue

        lista.append({
            "antes": item["antes"],
            "depois": item["depois"],
            "tipo": tipo,
            "titulo": titulo_final
        })

    return lista


def analisar_despesas_extras(despesas_df):
    """
    Analisa as despesas extras por categoria e forma de pagamento.
    Retorna dicionário com totais, categorias e alertas.
    """
    despesas_df['Valor'] = despesas_df['Valor (R$)'].apply(limpar_numero)
    despesas_df['Categoria'] = despesas_df['Categoria'].fillna('Outros')
    despesas_df['Pago com'] = despesas_df['Pago com'].fillna('Não especificado')

    total_despesas = despesas_df['Valor'].sum()

    # Agrupa por categoria
    por_categoria = despesas_df.groupby('Categoria')['Valor'].sum().round(2).to_dict()

    # Identifica despesas pagas fora do caixa
    fora_caixa = despesas_df[
        despesas_df['Pago com'].str.contains('próprio|gerente|pessoal', case=False, na=False)
    ]['Valor'].sum()

    # Lista detalhada de despesas
    lista_despesas = []
    for _, row in despesas_df.iterrows():
        lista_despesas.append({
            'data': str(row.get('Data', '')),
            'descricao': str(row.get('Descrição', '')),
            'categoria': str(row.get('Categoria', 'Outros')),
            'valor': limpar_numero(row.get('Valor (R$)', 0)),
            'pago_com': str(row.get('Pago com', '')),
            'observacao': str(row.get('Observação', ''))
        })

    return {
        'total': round(total_despesas, 2),
        'por_categoria': por_categoria,
        'fora_caixa': round(fora_caixa, 2),
        'lista': lista_despesas
    }


def analisar_sangrias(sangria_df):
    """
    Analisa as sangrias por motivo e identifica padrões.
    """
    sangria_df['Valor'] = sangria_df['Valor R$'].apply(limpar_numero)
    sangria_df['Motivo'] = sangria_df['Motivo'].fillna('Não especificado')

    total_sangrias = sangria_df['Valor'].sum()

    # Agrupa por motivo
    por_motivo = sangria_df.groupby('Motivo')['Valor'].sum().round(2).to_dict()

    # Conta quantidade de sangrias
    qtd_sangrias = len(sangria_df)

    # Lista detalhada
    lista_sangrias = []
    for _, row in sangria_df.iterrows():
        lista_sangrias.append({
            'data': str(row.get('Data', '')),
            'motivo': str(row.get('Motivo', '')),
            'observacoes': str(row.get('Observações', '')),
            'valor': limpar_numero(row.get('Valor R$', 0))
        })

    return {
        'total': round(total_sangrias, 2),
        'por_motivo': por_motivo,
        'quantidade': qtd_sangrias,
        'lista': lista_sangrias
    }


def gerar_insights(dados):
    """
    Gera insights automáticos baseados nos dados do relatório.
    """
    insights = []
    alertas = []
    recomendacoes = []

    # Análise de despesas sobre faturamento
    perc_despesas = (dados['saidas_total'] / dados['faturamento_total'] * 100) if dados['faturamento_total'] > 0 else 0

    if perc_despesas > 15:
        alertas.append(f"⚠️ Despesas representam {perc_despesas:.1f}% do faturamento (ideal: abaixo de 15%)")
        recomendacoes.append("Revisar custos operacionais e identificar oportunidades de redução")

    # Despesas fora do caixa
    if dados['despesas_extras']['fora_caixa'] > 0:
        alertas.append(f"🔴 R$ {dados['despesas_extras']['fora_caixa']:.2f} pagos com recursos próprios/gerente")
        recomendacoes.append("Estabelecer fundo de caixa pequeno para despesas emergenciais")

    # Análise de sangrias
    if dados['sangrias']['quantidade'] > 5:
        alertas.append(f"📊 {dados['sangrias']['quantidade']} sangrias realizadas na semana")
        recomendacoes.append("Avaliar se há necessidade de ajustar o fluxo de caixa inicial")

    # Ticket médio
    if dados['ticket_medio'] < 20:
        insights.append(f"💡 Ticket médio de R$ {dados['ticket_medio']:.2f} - oportunidade de upsell")
        recomendacoes.append("Treinar equipe em técnicas de venda adicional (combos, upgrades)")

    # Contas vencidas
    if dados['vencido'] > 0:
        alertas.append(f"💰 R$ {dados['vencido']:.2f} em contas vencidas")
        recomendacoes.append("Priorizar regularização de contas vencidas para evitar juros")

    return {
        'insights': insights,
        'alertas': alertas,
        'recomendacoes': recomendacoes,
        'perc_despesas_faturamento': round(perc_despesas, 1)
    }


def ler_dados():
    if not os.path.exists(EXCEL_FILE):
        return {"erro": f"Arquivo não encontrado:<br><b>{EXCEL_FILE}</b><br>Coloque na pasta data/"}

    try:
        # Leitura das planilhas
        vendas = pd.read_excel(EXCEL_FILE, sheet_name="VENDAS_DIARIAS", header=None, dtype=str)
        ticket = pd.read_excel(EXCEL_FILE, sheet_name="TICKET_MEDIO", dtype=str)
        formas = pd.read_excel(EXCEL_FILE, sheet_name="FORMAS_PAGAMENTO", dtype=str)
        produtos = pd.read_excel(EXCEL_FILE, sheet_name="PRODUTOS_VENDIDOS", dtype=str)
        contas_df = pd.read_excel(EXCEL_FILE, sheet_name="REGISTRO DE CONTAS", header=2, dtype=str)
        resumo = pd.read_excel(EXCEL_FILE, sheet_name="RESUMO GERAL", dtype=str)
        sangria_df = pd.read_excel(EXCEL_FILE, sheet_name="SANGRIA", dtype=str)
        despesas_df = pd.read_excel(EXCEL_FILE, sheet_name="DESPESAS EXTRAS", dtype=str)
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

        # NOVA ANÁLISE DE DESPESAS E SANGRIAS
        despesas_extras = analisar_despesas_extras(despesas_df)
        sangrias = analisar_sangrias(sangria_df)

        saidas_total = round(sangrias['total'] + despesas_extras['total'], 2)
        saldo_caixa = round(faturamento_total - saidas_total, 2)

        # CONTAS A PAGAR
        contas = []
        for _, row in contas_df.iterrows():
            if pd.notna(row.get('ID')) and str(row['ID']).strip():
                data_vencimento = str(row.get('DATA VENCIMENTO', ''))
                try:
                    data_obj = datetime.strptime(data_vencimento[:10], '%Y-%m-%d')
                    data_formatada = data_obj.strftime('%d/%m/%Y')
                except ValueError:
                    data_formatada = data_vencimento

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

        # Monta dicionário de dados
        dados = {
            'faturamento_total': round(faturamento_total, 2),
            'clientes_total': clientes_total,
            'ticket_medio': ticket_medio,
            'melhor_dia': melhor_dia,
            'melhor_dia_valor': round(melhor_dia_valor, 2),
            'datas_diarias': datas_diarias,
            'valores_diarios': [round(x, 2) for x in valores_diarios],
            'formas_pagamento': formas_data,
            'top_produtos': top_produtos,
            'sangrias': sangrias,
            'despesas_extras': despesas_extras,
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

        # GERA INSIGHTS
        dados['insights'] = gerar_insights(dados)

        return dados

    except Exception as e:
        import traceback
        return {"erro": f"Erro:<pre>{traceback.format_exc()}</pre>"}