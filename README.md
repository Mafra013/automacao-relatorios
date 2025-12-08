# AUTOMACAO DE RELATORIOS 

<img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"/>
<img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"/>
<img src="https://img.shields.io/badge/Microsoft_Word-2B579A?style=for-the-badge&logo=microsoft-word&logoColor=white" alt="Word"/>
<img src="https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" alt="Excel"/>
<img src="https://img.shields.io/badge/Automação-FF2D20?style=for-the-badge&logo=automation-anywhere&logoColor=white" alt="Automação"/>

---
## Sobre o projeto

Projeto desenvolvido com Python para automatizar a geração de relatórios profissionais em PDF com o histórico de revisões técnicas a partir de uma planilha Excel.
Ideal para empresas de transporte, fretamento, ônibus escolar ou qualquer frota que precise comprovar manutenção preventiva de forma rápida e padronizada.

---

## Tecnologias utilizadas
- **Python 3**
- **Pandas** → leitura e tratamento de dados
- **python-docx** → manipulação de arquivos Word
- **comtypes + pywin32** → conversão para PDF usando o próprio Microsoft Word
- **Excel** como fonte de dados

---
## FUNCIONALIDADES 
- Lê planilha Excel automaticamente 
- Agrupa revisões por **ID** (prefixo) do veículo 
- Preenche template Word com tabela dinâmica 
- Gera um **PDF individual** para cada veículo
- Insere cabeçalho (Cidade + data atual)
- Cria pasta com todos os relatórios prontos para envio ou impressão

---

## COMO USAR 

1.  Tenha uma planilha com o nome: dados.xlsx 
2.  Tenha o arquivo de modelo: template.docx 
3.  Execute o script: python automacao_relatorios.py

Todos os PDFs serão salvos na pasta: relatorios_gerados/

### EXEMPLO DE ESTRUTURA DA PLANILHA (aba "revisoes")
```
PREFIXO |  EMPRESA   | DATA REALIZADA | PROXIMA DATA (formula para 60 dias após a última) 

  100   | Fretamento |   15/03/2025   |   14/05/2025 

  100   | Fretamento |   16/05/2025   |   15/07/2025 

  101   | Fretamento |   20/03/2025   |   19/05/2025 

  102   | Fretamento |   22/03/2025   |   21/05/2025 

  ...   |   ...      |      ...       | ...
```
### SAIDA GERADA
```
├── relatorios_gerados/  ← (pasta criada automaticamente)
│ ├── Relatorio_100.pdf
│ ├── Relatorio_101.pdf
│ ├── Relatorio_102.pdf
│ └── ...
```

## DOCUMENTAÇÃO TECNICA
**POR QUE USEI O WORD PARA GERAR PDF?**
- Mantem 100% do layout (logos, tabelas com bordas, fontes especiais)
- Resultado identico ao que seria feito manualmente
  
COMO FUNCIONA O SCRIPT (passo a passo)
1.  Lê a planilha com pandas 
2.  Converte datas automaticamente (aceita número serial do Excel) 
3.  Agrupa por PREFIXO 
4.  Para cada veículo: 
- Abre o template.docx
- Escreve cidade + data por extenso no cabecalho
- Preenche a tabela com todas as revisoes daquele veículo
- Salva como PDF usando o Microsoft Word

