# app/app.py
import os
from flask import Flask, render_template, send_file, make_response
# Importa a função de conversão
from utils import ler_dados, get_absolute_file_url
import pdfkit

app = Flask(__name__)

# Configuração do wkhtmltopdf (ajusta se precisar)
config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')
# Se der erro, comenta a linha acima e deixa só: config = pdfkit.configuration()

@app.route('/')
def dashboard():
    dados = ler_dados()
    return render_template('dashboard.html', **dados)

@app.route('/pdf')
def pdf():
    dados = ler_dados()

    # --- BLOCO: CONVERSÃO DE CAMINHOS PARA PDF (file:///) ---
    for item in dados.get('fotos_melhorias', []):
        if item.get('antes'):
            item['antes'] = get_absolute_file_url(item['antes'])
        if item.get('depois'):
            item['depois'] = get_absolute_file_url(item['depois'])

    # Renderiza o HTML usando os caminhos ABSOLUTOS (file:///)
    html = render_template('dashboard.html', **dados)
    # ----------------------------------------------------

    # Opções para PDF ficar lindo
    options = {
        'page-size': 'A4',
        'margin-top': '0.5in',
        'margin-right': '0.5in',
        'margin-bottom': '0.5in',
        'margin-left': '0.5in',
        'encoding': "UTF-8",
        'no-outline': None,
        'enable-local-file-access': None, # Crucial para 'file:///'
        'quiet': ''
    }

    pdf = pdfkit.from_string(html, False, options=options, configuration=config)

    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'attachment; filename=Relatorio_Chiquinho_Sorvetes.pdf'
    return response

if __name__ == '__main__':
    app.run(debug=True, port=5000)