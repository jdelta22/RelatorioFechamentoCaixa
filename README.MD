# 📊 Relatório de Fechamento de Caixa — PDF → Excel

![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Formatting-orange)
![Tabula](https://img.shields.io/badge/Tabula-PDF--Extractor-red)
![Status](https://img.shields.io/badge/Status-Working-success)

---

## 🧾 Visão Geral

Este projeto automatiza o processo de **leitura de relatórios de vendas em PDF** e gera uma **planilha Excel pré-formatada** para fechamento de caixa.  
Ele combina o poder de **Pandas**, **Tabula**, **Pdfplumber**, **regex* e **OpenPyXL** para extrair dados, tratá-los e aplicar formatações monetárias, bordas e cores automaticamente.

---

## ⚙️ Funcionalidades

✅ Extrai automaticamente tabelas de relatórios PDF  
✅ Limpa e organiza os dados de forma estruturada  
✅ Insere o cabeçalho **“CAIXA DIA DD/MM/YYYY”** com a data atual  
✅ Aplica formatação monetária (R$) e bordas personalizadas  
✅ Permite configurar bordas **finas**, **grossas** ou **sem borda**  
✅ Gera um arquivo Excel formatado, pronto para uso contábil  

---

## 🧰 Tecnologias Utilizadas

| Biblioteca | Função |
|-------------|--------|
| **tabula-py** | Extração de tabelas do PDF |
| **pandas** | Limpeza e manipulação de dados |
| **openpyxl** | Edição e formatação do Excel |
| **jpype** | Interface entre Python e Java para o Tabula |
| **pdfplumber** | Extração de textos do PDF |
| **regex** | manipulção de textos |


---

## 📦 Instalação
Atenção!!!!!

Necessario instalação do Python(https://www.python.org/) e alguma distribuição JDK para Java
neste projeto eu tinha instalado a Temurin® JDK(Temurin® JDK) que pode ser adiquirida em https://adoptium.net/pt-BR/temurin/releases

Clone o repositório e configure o ambiente virtual:

```bash
git clone https://github.com/seuusuario/RelatorioFechamentoCaixa.git
cd RelatorioFechamentoCaixa
python -m venv venv
# Ative o ambiente
venv\Scripts\activate  # Windows
source venv/bin/activate  # Linux/macOS
Instale as dependências:

bash
Copiar código
pip install -r requirements.txt
⚠️ É necessário ter o Java instalado para que o tabula-py funcione corretamente.

📁 Estrutura do Projeto
bash
Copiar código
RelatorioFechamentoCaixa/
│
├── main.py                # Script principal
├── modelo_caixa.xlsx      # Modelo de planilha Excel (pré-formatado)
├── VendasPorProdutos.pdf  # Relatório PDF de exemplo
├── CAIXA_fechado.xlsx     # Planilha gerada automaticamente
├── requirements.txt
└── README.md
🚀 Como Usar
Coloque o arquivo VendasPorProdutos.pdf na raiz do projeto

Certifique-se de ter o arquivo modelo modelo_caixa.xlsx

Execute o script:

bash
Copiar código
python main.py
O terminal exibirá o progresso:

objectivec
Copiar código
Lendo o arquivo VendasPorProdutos.pdf...
✅ Arquivo PDF lido com sucesso!
Lendo o arquivo `RelatorioFechamentoCaixaBobina.pdf`...
✅ Arquivo PDF lido com sucesso!
✅ Arquivo Excel gerado: CAIXA_fechado.xlsx
O resultado será salvo automaticamente como CAIXA_fechado.xlsx

🎨 Formatação Aplicada no Excel
Elemento	Descrição
Cabeçalho	“CAIXA DIA DD/MM/YYYY” mesclado de A1:H1, negrito, fonte 16
Células monetárias	Formato “R$ #,##0.00”
Bordas	Finas ou grossas, conforme parâmetro da função
Cores de destaque	Fundo amarelo em células de totalização

🧠 Exemplo de Função de Estilo
python
Copiar código
def aplicar_estilo(faixa, cor=False, borda="fina"):
    intervalo = caixa[faixa]
    if not isinstance(intervalo, tuple):
        intervalo = [[intervalo]]

    borda_style = openpyxl.styles.Border(
        left=openpyxl.styles.Side(style='medium' if borda == "grossa" else 'thin'),
        right=openpyxl.styles.Side(style='medium' if borda == "grossa" else 'thin'),
        top=openpyxl.styles.Side(style='medium' if borda == "grossa" else 'thin'),
        bottom=openpyxl.styles.Side(style='medium' if borda == "grossa" else 'thin')
    )

    for row in intervalo:
        for cell in row:
            cell.style = moeda_style
            if cor:
                cell.fill = fundo
            cell.border = borda_style

👨‍💻 Autor
João Pedro Calaça Costa
💼 Desenvolvedor Python / Automação de Processos
📧 jotao22gm@gmail.com
🌐 https://www.linkedin.com/in/jo%C3%A3o-cala%C3%A7a-8aa190338/
📦 GitHub: https://github.com/jdelta22

🪪 Licença
Distribuído sob a licença MIT.
Sinta-se livre para usar, modificar e distribuir com os devidos créditos.

💬 "Automatizar o que é repetitivo é abrir espaço para pensar o que é estratégico."






