# DocFlow — Automação de documentos para prestação de contas

Aplicativo em **Python + Streamlit** para organizar PDFs, relacionar notas fiscais e comprovantes de pagamento e gerar relatórios Excel.

> Projeto de portfólio reconstruído em 2026 a partir de automações administrativas desenvolvidas em 2025. Não é um produto fiscal homologado, nem realiza validação tributária de notas fiscais.

## Funcionalidades (v0.1)

- **Juntar PDFs em lote:** adiciona um comprovante antes de cada nota fiscal selecionada, sem alterar os originais.
- **Buscar PDFs:** pesquisa por nome, ano e texto selecionável contido nos arquivos.
- **Analisar / sugerir renomeação:** tenta identificar número de nota, CNPJ e valor via padrões textuais. Exige conferência humana; não altera os arquivos originais.
- **Mesclar por planilha:** lê a coluna `NF` de um Excel, associa notas e comprovantes cujos nomes contenham o mesmo número, e marca casos ambíguos para revisão.
- **Relatórios:** exporta Excel com status, detalhes e caminhos de saída; os PDFs gerados são baixados em ZIP.

## Instalação

Requer Python 3.10+.

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
# source .venv/bin/activate
python -m pip install -r requirements.txt
python -m pip install -e .
python -m streamlit run app.py
```

Acesse a URL local exibida no terminal.

## Testes

```bash
python -m pip install -e '.[dev]'
python -m pytest -q
```

## Estrutura

```text
app.py                   # interface web
src/docflow/core.py      # regras de processamento, independentes da interface
tests/test_core.py       # testes de operações e ambiguidades
examples/                # dados sintéticos para demonstração
```

## Limitações conhecidas

- PDFs digitalizados sem camada de texto **não são lidos por OCR** nesta versão.
- A extração de valores e números usa expressões regulares; diferentes layouts podem produzir erros ou campos vazios.
- A mesclagem por planilha é baseada em **número presente no nome do arquivo**, não em verificação fiscal completa. Quando houver múltiplos candidatos, não faz combinação automática.
- A busca por ano considera o ano no nome/caminho (ou no texto quando a busca por conteúdo estiver ativa); não infere a competência fiscal.
- A interface trabalha com uploads e downloads temporários, não acessa pastas locais arbitrárias no servidor.
- Para dados sensíveis, utilize ambiente local e siga as políticas de privacidade aplicáveis.

## Próximos passos

1. OCR opcional com Tesseract para PDFs digitalizados.
2. Identificação de notas e comprovantes com múltiplos critérios (CNPJ, valor, data, número).
3. Pré-visualização e confirmação manual de renomeações.
4. Autenticação e armazenamento seguro, se houver necessidade de uso multiusuário.
5. Demonstração em vídeo e capturas de tela com dados fictícios.

## Origem do projeto

Inspirado em scripts de automação usados para reduzir tarefas repetitivas de prestação de contas em 2025. A implementação pública usa apenas exemplos sintéticos e código reconstruído, sem arquivos ou informações de terceiros.

## Licença

MIT — consulte `LICENSE`.
