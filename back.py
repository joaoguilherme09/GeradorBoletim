import pandas as pd
from prettytable import PrettyTable
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))

CAMINHO_HTMLS = "F:\\4bi\\alberson\\pasta_html\\"
CAMINHO_PDFS = "F:\\4bi\\alberson\\pasta_pdf\\"


def ler_planilha():
    while True:
        caminho_planilha = input("Digite o caminho completo da planilha (.xlsx): ").strip().strip('"').strip("'")
        try:
            planilha = pd.read_excel(caminho_planilha)
            print("\n✅ Planilha carregada com sucesso!\n")
            return planilha
        except:
            print("❌ Caminho inválido ou erro ao abrir o arquivo. Tente novamente.")

# ------------------------------------------------------------
# FUNÇÃO: MOSTRAR E PROCESSAR ALUNOS
# ------------------------------------------------------------
def processar_alunos(planilha):
    colunas = planilha.columns
    total_linhas = len(planilha)
    total_colunas = len(colunas)

    for linha in range(total_linhas):
        tabela = PrettyTable()
        tabela.field_names = colunas

        dados_aluno = []
        for coluna in range(total_colunas):
            valor = planilha.iloc[linha, coluna]
            if pd.isna(valor):
                valor = "Não informada"
            dados_aluno.append(valor)

        tabela.add_row(dados_aluno)

        
        print(f"\n Aluno encontrado na linha {linha + 2} da planilha:")
        print(tabela)

        gerar_html(colunas, dados_aluno)
        gerar_pdf(colunas, dados_aluno)
        
        
# ------------------------------------------------------------
# FUNÇÃO: GERAR ARQUIVO HTML DO ALUNO
# ------------------------------------------------------------

def gerar_html(colunas, dados):
    codigo = str(dados[0])
    nome = dados[1]
    turma = dados[2]
    nome_pdf = f"../pasta_pdf/{codigo}.pdf"
    caminho_html = CAMINHO_HTMLS + codigo + ".html"

    # Gera as linhas da tabela colorindo conforme a nota
    linhas_html = ""
    for i in range(3, len(colunas)):
        valor = dados[i]
        if pd.isna(valor) or valor == "Não informada":
            valor_exibido = "Não informada"
            linhas_html += f"<tr><td>{colunas[i]}</td><td>{valor_exibido}</td></tr>\n"
        else:
            try:
                valor_num = float(valor)
                if valor_num < 6:
                    linhas_html += f"<tr><td>{colunas[i]}</td><td style='color:red;'>{valor_num}</td></tr>\n"
                else:
                    linhas_html += f"<tr><td>{colunas[i]}</td><td style='color:blue;'>{valor_num}</td></tr>\n"
            except:
                linhas_html += f"<tr><td>{colunas[i]}</td><td>{valor}</td></tr>\n"


    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8">
    <title>Boletim de {nome}</title>
    <style>
        :root {{
            --bg-color: #eef2f7;
            --text-color: #1a1a1a;
            --card-bg: rgba(255, 255, 255, 0.8);
            --accent-color: #007bff;
            --table-header: #007bff;
            --table-border: #ddd;
            --shadow-color: rgba(0,0,0,0.1);
        }}
        body.dark {{
            --bg-color: #0e1117;
            --text-color: #f1f1f1;
            --card-bg: rgba(30, 34, 45, 0.8);
            --accent-color: #00b4d8;
            --table-header: #00b4d8;
            --table-border: #2b2f3a;
            --shadow-color: rgba(0,0,0,0.6);
        }}
        body {{
            margin: 0;
            padding: 0;
            font-family: 'Poppins', 'Segoe UI', sans-serif;
            background: var(--bg-color);
            color: var(--text-color);
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: flex-start;
            min-height: 100vh;
            transition: background 0.5s, color 0.5s;
        }}
        .container {{
            margin-top: 40px;
            background: var(--card-bg);
            backdrop-filter: blur(15px);
            box-shadow: 0 8px 30px var(--shadow-color);
            border-radius: 20px;
            padding: 40px;
            width: 80%;
            max-width: 800px;
            transition: background 0.5s, box-shadow 0.5s;
            animation: fadeIn 1s ease-in-out;
        }}
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(30px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        .logos {{
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 30px;
            margin-bottom: 20px;
        }}
        .logos img {{
            width: 90px;
            height: auto;
            border-radius: 12px;
            box-shadow: 0 0 20px rgba(0,0,0,0.15);
        }}
        h1 {{
            font-size: 2.2em;
            margin-bottom: 10px;
            text-align: center;
            color: var(--accent-color);
        }}
        h2 {{
            text-align: center;
            margin-bottom: 30px;
            color: var(--text-color);
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            border-radius: 15px;
            overflow: hidden;
            box-shadow: 0 4px 20px var(--shadow-color);
        }}
        th {{
            background: var(--table-header);
            color: white;
            text-align: center;
            padding: 14px;
            font-size: 1em;
        }}
        td {{
            border-bottom: 1px solid var(--table-border);
            padding: 12px;
            text-align: center;
            font-size: 1em;
        }}
        .btn-theme {{
            position: fixed;
            top: 20px;
            right: 20px;
            background: transparent;
            border: 2px solid var(--accent-color);
            color: var(--accent-color);
            border-radius: 50%;
            width: 45px;
            height: 45px;
            cursor: pointer;
            font-size: 20px;
            transition: all 0.3s;
        }}
        .btn-theme:hover {{
            background: var(--accent-color);
            color: white;
        }}
        .btn {{
            margin-top: 25px;
            padding: 10px 20px;
            background: var(--accent-color);
            color: white;
            border: none;
            border-radius: 10px;
            cursor: pointer;
            font-weight: bold;
        }}
        .btn:hover {{
            background: #0056b3;
        }}
    </style>
</head>
<body>
    <button class="btn-theme" onclick="toggleTheme()">🌙</button>

    <div class="container">
    
        <h1>Boletim Escolar</h1>
        <h2>{nome} — Turma {turma}</h2>

        <table>
            <tr><th>Disciplina</th><th>Nota</th></tr>
            {linhas_html}
        </table>

        <div style="text-align:center;">
            <button class="btn" onclick="window.open('{nome_pdf}', '_blank')">📄 Abrir PDF</button>
        </div>
    </div>

    <script>
        function toggleTheme() {{
            document.body.classList.toggle('dark');
            const icon = document.querySelector('.btn-theme');
            icon.textContent = document.body.classList.contains('dark') ? '☀️' : '🌙';
        }}
    </script>
</body>
</html>"""

    with open(caminho_html, "w", encoding="utf-8") as arquivo:
        arquivo.write(html)

# ------------------------------------------------------------
# FUNÇÃO: GERAR ARQUIVO PDF DO ALUNO
# ------------------------------------------------------------
def gerar_pdf(colunas, dados):
    codigo = str(dados[0])
    nome = dados[1]
    turma = dados[2]
    caminho_pdf = CAMINHO_PDFS + codigo + ".pdf"

    pdf = canvas.Canvas(caminho_pdf)
    pdf.setTitle(f"Boletim - {nome}")
    pdf.setFont("Arial", 16)
    pdf.drawString(50, 800, f"Boletim do Aluno: {nome}")
    pdf.setFont("Arial", 12)
    pdf.drawString(50, 780, f"Código: {codigo}")
    pdf.drawString(50, 760, f"Turma: {turma}")

    y = 730
    pdf.drawString(50, y, "Notas e Média:")
    y -= 20

    for coluna in range(3, len(colunas)):
        nome_coluna = colunas[coluna]
        valor = dados[coluna]
        if valor == "Não informada":
            pdf.setFillColor("black")
            texto = f"{nome_coluna}: Não informada"
        elif isinstance(valor, (int, float)) and valor < 6:
            pdf.setFillColor("red")
            texto = f"{nome_coluna}: {valor}"
        else:
            pdf.setFillColor("blue")
            texto = f"{nome_coluna}: {valor}"

        pdf.drawString(70, y, texto)
        y -= 20

    pdf.save()


# ------------------------------------------------------------
# PROGRAMA PRINCIPAL
# ------------------------------------------------------------
def main():
    dados_planilha = ler_planilha()
    processar_alunos(dados_planilha)
    print("\n✅ Todos os HTMLs e PDFs foram gerados com sucesso!")

# ------------------------------------------------------------
# EXECUÇÃO
# ------------------------------------------------------------
main()
