# 🧾 Gerador de Boletins em HTML e PDF

Este projeto em *Python* lê uma planilha Excel contendo informações de alunos e suas notas, e gera automaticamente *boletins personalizados* em formato *HTML* e *PDF*.  
O sistema foi criado para facilitar a visualização dos resultados de cada aluno, com destaque visual para notas abaixo da média.

---

## 📋 Funcionalidades

- 📂 Leitura de planilhas .xlsx com dados dos alunos (usando *pandas*).  
- 🧮 Exibição formatada dos dados no terminal (usando *PrettyTable*).  
- 🌐 Geração automática de páginas *HTML* estilizadas com modo claro/escuro.  
- 📄 Criação de *PDFs* individuais com as notas de cada aluno (usando *ReportLab*).  
- 🎨 Destaque em *vermelho* para notas abaixo de 6 e *azul* para notas iguais ou acima de 6.  
- ✅ Substituição automática de campos vazios por “Não informada”.

---

## 🛠️ Tecnologias Utilizadas

- *Python 3.10+*
- *pandas* – para leitura da planilha Excel  
- *prettytable* – para exibição dos dados no terminal  
- *reportlab* – para geração dos arquivos PDF  

------------------------------------------------------------------------------------------------------------

## ▶️ Como Executar

1. Certifique-se de ter o Python e as bibliotecas instaladas:
   ```bash
   pip install pandas prettytable reportlab

2. Coloque a planilha Excel (.xlsx) em um local acessível no seu computador.

3. Execute o script:
python back.py

4. Informe o caminho completo da planilha quando solicitado.

5. Os arquivos HTML e PDF serão gerados automaticamente nas pastas configuradas.
