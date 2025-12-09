# app/app.py
import os
from flask import Flask, render_template, make_response
from utils import ler_dados, get_absolute_file_url
import pdfkit

app = Flask(__name__)

# Configuração do wkhtmltopdf (ajusta se precisar)
config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')

# --- NOVO FILTRO JINJA2 PARA FORMATAR MOEDA BRASILEIRA (R$ X.XXX,XX) ---
def formatar_brl(valor):
    """Formata um float para o padrão monetário brasileiro."""
    try:
        if valor is None:
            valor = 0.0
        # 1. Formata o número com 2 casas decimais e separador de milhar (,)
        # 2. Troca o ponto decimal para vírgula
        # 3. Troca a vírgula de milhar por ponto (padrão brasileiro)
        return f"{valor:,.2f}".replace(",", "_TEMP_").replace(".", ",").replace("_TEMP_", ".")
    except (TypeError, ValueError):
        # Retorna o valor bruto se a formatação falhar
        return f"{valor}"

app.jinja_env.filters['brl'] = formatar_brl
# -----------------------------------------------------------------------


@app.route('/')
def dashboard():
    dados = ler_dados()
    return render_template('dashboard.html', **dados)

@app.route('/pdf')
def pdf():
    dados = ler_dados()

    # CONVERSÃO DE CAMINHOS PARA PDF (file:///)
    for item in dados.get('fotos_melhorias', []):
        if item.get('antes'):
            item['antes'] = get_absolute_file_url(item['antes'])
        if item.get('depois'):
            item['depois'] = get_absolute_file_url(item['depois'])

    html = render_template('dashboard.html', **dados)

    options = {
        'page-size': 'A4',
        'margin-top': '0.5in',
        'margin-right': '0.5in',
        'margin-bottom': '0.5in',
        'margin-left': '0.5in',
        'encoding': "UTF-8",
        'no-outline': None,
        'enable-local-file-access': None,
        'quiet': ''
    }

    pdf = pdfkit.from_string(html, False, options=options, configuration=config)

    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'attachment; filename=Relatorio_Chiquinho_Sorvetes.pdf'
    return response

if __name__ == '__main__':
    app.run(debug=True, port=5000)