# DocFlow

**Automação documental para fluxos de prestação de contas.**

DocFlow é uma aplicação web construída com **Python e Streamlit** para organizar lotes de PDFs, localizar documentos a partir de listas ou planilhas Excel e combinar comprovantes com notas fiscais. O objetivo é reduzir tarefas manuais repetitivas e facilitar a conferência dos resultados.

> Projeto de portfólio em desenvolvimento. Não é um sistema fiscal homologado.

## O que o projeto faz

| Ferramenta | Funcionalidade |
| --- | --- |
| Localizar e separar | Pesquisa PDFs pelo nome ou texto interno, usando termos digitados ou uma coluna do Excel. Gera ZIP com documentos encontrados e relatório. |
| Juntar documentos | Coloca um comprovante antes de cada nota fiscal, criando PDFs separados. |
| Analisar e sugerir nomes | Extrai campos de PDFs textuais e sugere nomes, sem modificar os originais. |
| Mesclar por planilha | Relaciona notas e comprovantes por número de NF no nome do arquivo, sinalizando ambiguidades. |

### Interface

A interface utiliza uma identidade visual própria com degradês em vermelho e grafite, ilustrações em CSS com profundidade, navegação lateral e transições suaves. O layout é responsivo e respeita a preferência do sistema por movimento reduzido.

## Como usar

### Localizar documentos por planilha

1. Compacte a pasta com os PDFs em um arquivo `.zip`.
2. Abra **Localizar e separar PDFs**.
3. Envie o ZIP ou selecione vários PDFs.
4. Escolha **Planilha Excel** e envie um arquivo `.xlsx` com uma coluna `NF` (ou informe outro nome de coluna).
5. Clique em **Localizar e separar**.
6. Baixe o ZIP de saída com a pasta `encontrados/` e o arquivo `relatorio.xlsx`.

Para uma busca simples, escolha **Nome / lista** e digite um nome ou número por linha. A busca por conteúdo é opcional e pode ser mais lenta.

### Juntar comprovantes e notas

Envie um comprovante e as notas desejadas. O DocFlow gera um PDF para cada nota, mantendo o comprovante primeiro e preservando os arquivos originais.

### Mesclar por planilha

Prepare um Excel com a coluna `NF` e arquivos cujos nomes contenham os números correspondentes. Envie notas e comprovantes separadamente. Correspondências ausentes ou ambíguas são marcadas para revisão.

## Tecnologias

- Python 3.10+
- Streamlit — interface web
- pypdf — leitura e combinação de PDFs
- openpyxl — leitura e exportação de planilhas
- pytest — testes automatizados
- ReportLab — criação de PDFs sintéticos nos testes de desenvolvimento

## Rodar localmente

```bash
git clone https://github.com/thlozz/docflow.git
cd docflow
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Para executar os testes:

```bash
python -m pip install -e '.[dev]'
python -m pytest -q
```

Os testes incluem cenários sintéticos com comprovantes e notas de múltiplas páginas, números de NF semelhantes, planilhas incompletas, correspondências ambíguas e arquivos corrompidos. **A presença dos testes não implica que todos já tenham sido executados com sucesso no ambiente de produção.**

## Estrutura

```text
docflow/
├── app.py                 # Interface Streamlit
├── src/docflow/core.py    # Regras de processamento
├── tests/                 # Testes automatizados
├── examples/              # Exemplos, quando disponíveis
├── requirements.txt
└── pyproject.toml
```

## Limitações e segurança

- PDFs digitalizados sem texto selecionável ainda não contam com OCR.
- A extração de campos utiliza padrões textuais e pode não funcionar em todos os layouts.
- A mesclagem por planilha depende dos números presentes nos nomes dos arquivos; não valida pagamentos ou autenticidade fiscal.
- O processamento ocorre temporariamente no servidor da aplicação; **não envie documentos confidenciais para uma demonstração pública**.
- O envio de uma pasta pelo navegador é realizado por ZIP, não por acesso direto ao sistema de arquivos do computador.

## Próximos passos

Melhorar a cobertura de testes, adicionar pré-visualização e revisão manual de resultados e estudar OCR opcional para PDFs digitalizados.

---

Desenvolvido como projeto de portfólio por [@thlozz](https://github.com/thlozz). Licença MIT.
