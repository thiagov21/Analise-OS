
# Analise-OS 📊🛠️

**Smart Analyzer** é uma aplicação Python com interface gráfica que automatiza o processo de análise e aprovação de Ordens de Serviço (OS), substituindo tarefas manuais demoradas por um fluxo mais rápido, preciso e eficiente.

Lembrando que os dados apresentados nesse git são ficticios, para inserir esse projeto na sua empresa, me contate via linkedin!

---

## 🚀 Objetivo

Reduzir significativamente o tempo e o esforço necessários para analisar ordens de serviço, minimizando erros e retrabalho, e aumentando a produtividade.

---

## ⚙️ Funcionalidades

- 🔐 Login com autenticação
- 🔍 Busca de OS individual com extração de dados via web scraping
- 📊 Análise em lote de planilhas (.xlsx) com múltiplas OS
- 🧾 Exportação de resultados para Excel
- 📈 Barra de progresso e feedback em tempo real
- 💡 Interface gráfica amigável com Tkinter

---

## 📉 Resultados

- 🔧 Antes: ~6h35min para analisar 526 OS manualmente
- ⚡ Depois: ~19min para analisar as mesmas 526 OS automaticamente
- ⏱️ Redução de tempo: **95,1%**

---

## 🛠️ Tecnologias Usadas

- Python 3
- Tkinter + TtkThemes (GUI)
- BeautifulSoup (Web Scraping)
- Pandas (manipulação de dados)
- PIL (para imagens)
- JSON / Excel
- Threading (para manter a GUI fluida)

---

## 💻 Como Rodar o Projeto

1. Clone o repositório:
   ```bash
   git clone https://github.com/thiagov21/Analise-OS.git
   cd Analise-OS
   ```

2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

3. Rode o projeto:
   ```bash
   python main.py
   ```

---

## 🧠 Observações

- Certifique-se de ter uma conexão com a internet, pois a aplicação realiza scraping em um sistema online.
- A estabilidade do sistema depende da estrutura do site de origem. Mudanças podem exigir manutenção no código.

---

## 👨‍💻 Autor

Desenvolvido por **Thiago Vieira**

---

## 🔒 Licença

Este projeto é de uso **proprietário**. Todos os direitos reservados.
