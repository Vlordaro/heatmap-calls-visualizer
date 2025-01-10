# Heatmap Calls Visualizer

Ferramenta para análise e visualização de chamados através de heatmap com integração Google Sheets.

## Features
- Visualização em heatmap por mês
- Integração com Google Sheets
- Exportação PNG/PDF/Excel
- Detalhes por horário
- Modo tela cheia
- Atualização automática
- Estatísticas de atendimento

## Setup

1. Google Cloud:
   - Crie projeto no Console
   - Ative Google Sheets API
   - Crie credencial Service Account
   - Baixe credentials.json
   - Coloque na pasta raiz

2. Google Sheets:
   - Crie planilha
   - Compartilhe com email do Service Account
   - Copie ID da URL (entre /d/ e /edit)
   - Substitua 'your-spreadsheet-id' no código

3. Estrutura Planilha:
   - Created (data/hora início)
   - Closed (data/hora fim)
   - Customer (ID chamado)
   - Category (tipo chamado)

4. Instalação:
  - git clone https://github.com/seu-usuario/heatmap-calls-visualizer
  - cd heatmap-calls-visualizer
  - pip install -r requirements.txt
  - python heatmap.py

   
## Uso
- Selecione ano/mês para visualizar
- F11: Alterna tela cheia
- Clique: Mostra detalhes do horário
- Botão Exportar: Salva em PNG/PDF/Excel

## Requisitos
- Python 3.x
- Google Cloud
- Bibliotecas em requirements.txt

## Licença
MIT
