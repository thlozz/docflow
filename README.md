# DocFlow — automação de documentos

Aplicação web em **Python + Streamlit**, reconstruída a partir de automações de prestação de contas. Versão **0.2**.

## Ferramentas

- **Localizar e separar PDFs:** envie uma pasta compactada em ZIP (com subpastas, se quiser) ou vários PDFs. Busque por uma lista de nomes/números ou pela coluna de uma planilha Excel. Baixe um ZIP contendo **encontrados/** e **relatorio.xlsx**, inclusive os termos não encontrados.
- **Juntar PDFs:** coloca um comprovante antes de cada nota selecionada, criando arquivos separados.
- **Analisar e renomear:** sugere nomes com número da NF e valor extraídos de PDFs com texto selecionável. Não renomeia originais.
- **Mesclar por planilha:** relaciona PDFs de notas e comprovantes pelo número da NF nos nomes. Ambiguidades ficam para revisão.

## Como usar a separação por planilha

1. Compacte a pasta de PDFs em um arquivo `.zip` (no Chromebook, pelo app Arquivos).
2. Na aba **Localizar e separar PDFs**, envie o ZIP.
3. Escolha **Planilha Excel** e envie um `.xlsx` com a coluna `NF` (ou informe outro cabeçalho).
4. Clique **Localizar e separar**.
5. Baixe o ZIP com os PDFs encontrados e o relatório dos não encontrados.

Alternativamente, selecione **Nome / lista** e digite os nomes ou números, um por linha.

**Importante:** em uma aplicação hospedada na nuvem não é possível acessar livremente as pastas locais do usuário. O ZIP funciona como a pasta de origem e o ZIP de resultado como a pasta separada.

## Executar localmente

Requer Python 3.10+.

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Testes

```bash
python -m pip install -e '.[dev]'
python -m pytest -q
```

## Estrutura

- `app.py`: interface e fluxo de upload/exportação
- `src/docflow/core.py`: funções de processamento
- `tests/`: testes automatizados
- `examples/`: documentos fictícios de demonstração

## Limitações e privacidade

- Não há OCR para PDFs escaneados sem texto selecionável.
- Busca por nome é aproximada para texto e delimitada por números para evitar coincidências como 1001 dentro de 11001.
- A busca no conteúdo pode encontrar um número em qualquer contexto; confira os resultados.
- Mesclagem por planilha exige identificação da NF nos nomes dos arquivos e não verifica pagamentos.
- A extração de campos usa padrões e pode falhar com diferentes layouts.
- Arquivos enviados são processados temporariamente no servidor; **não envie documentos empresariais sigilosos à demonstração pública**.
- Esta aplicação não é um produto fiscal homologado.

## Próximas melhorias

OCR opcional, correspondência por fornecedor/valor/data, prévia e confirmação manual das renomeações, testes de integração da interface e melhorias de acessibilidade.

MIT License.
