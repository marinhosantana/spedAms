# Documentacao da estrutura do sistema

Este documento resume como o projeto esta organizado, onde fica cada responsabilidade e quais arquivos devem ser alterados para cada tipo de manutencao.

## Visao geral

O sistema e uma aplicacao desktop em Tkinter para leitura, analise, conciliacao, comparacao, ajuste e exportacao de dados SPED/XML/planilhas.

Fluxo geral:

1. `main.py` chama `Revisor.py`.
2. `Revisor.py` cria a janela Tkinter e instancia `SpedApp`.
3. `app/ui/app.py` monta a interface, recebe a acao do usuario e chama servicos.
4. `app/services/` executa as regras de negocio.
5. `app/parsers/` le arquivos SPED, XML e Excel.
6. `app/exporters/` grava Excel, CSV e Word.
7. `app/repositories/` acessa banco de dados.

## Entrada da aplicacao

- `Sped/main.py`
  - Ponto de entrada simples.
  - Importa `main` de `Revisor.py`.

- `Sped/Revisor.py`
  - Cria `Tk()`.
  - Instancia `SpedApp`.
  - Inicia `root.mainloop()`.

## UI

- `Sped/app/ui/app.py`
  - Arquivo principal da interface.
  - Contem a classe `SpedApp`.
  - Responsavel por telas, botoes, grids, menus, filtros, popups, mensagens e chamadas aos servicos.
  - Deve orquestrar fluxos, mas nao deve concentrar regra de negocio nova quando ela puder morar em `services/`.

- `Sped/app/ui/progress_dialog.py`
  - Contem `ProgressDialogHandle`.
  - Centraliza atualizar, resetar e fechar progress dialogs.
  - O `app.py` monta a janela em `open_progress_dialog()` e usa `progress_dialog()` como context manager.
  - Fluxos que reutilizam progresso entre etapas devem passar `external_progress`.

- `Sped/app/ui/app_config.json`
  - Configuracao visual/local da aplicacao, como titulos.

- `Sped/app/ui/mysql_config.json`
  - Configuracao local de conexao MySQL.

- `Sped/app/ui/logs/`
  - Pasta de logs/auditoria gerados em runtime.

## Configuracao e modelos

- `Sped/app/config.py`
  - Constantes globais.
  - Padroes de MySQL.
  - Marcadores visuais da comparacao.
  - Namespaces/padroes usados na leitura XML.

- `Sped/app/models.py`
  - Dataclasses compartilhadas.
  - Exemplos: `ProductRecord`, `CompareXmlInvoice`, `CompareXmlItem`, `CompareSheetInvoice`, `CompareSpedDocument`.
  - Use este arquivo quando precisar de estruturas tipadas reaproveitadas entre parser, service e UI.

## Parsers

Pasta: `Sped/app/parsers/`

Responsabilidade: transformar arquivos externos em estruturas Python.

- `sped_parser.py`
  - Utilitarios basicos de SPED.
  - Le linhas, normaliza campos, parseia decimais, CFOP, E110 e registros sumarizados.

- `sped_fiscal_parser.py`
  - Le SPED Fiscal detalhado.
  - Extrai itens, documentos, chaves, CST, CFOP, bases e valores.

- `compare_sped_reader.py`
  - Le documentos do SPED para fluxo de comparacao.
  - Coleta chaves/documentos e extrai CNPJ da empresa.

- `compare_xml.py`
  - Le XML NF-e/NFC-e/NFS-e para comparacao.
  - Monta `CompareXmlInvoice` e `CompareXmlItem`.
  - Tambem trata XMLs ignorados/cancelamentos.

- `compare_sheet.py`
  - Le planilhas de comparacao.
  - Converte `.xls` quando necessario.
  - Coleta notas/chaves de planilha.

- `excel_parser.py`
  - Le XLSX sem depender da UI.
  - Resolve abas, shared strings, cabecalhos e filtros.

## Services

Pasta: `Sped/app/services/`

Responsabilidade: regra de negocio, transformacoes, conciliacoes, calculos e preparacao de datasets para tela/exportacao.

- `analysis_reports.py`
  - Agrega varios relatorios e builders usados pela UI.
  - Atua como fachada para analises fiscais, periodos, diagnosticos e relatorios mensais.

- `period_comparisons.py`
  - Monta comparativos por periodo de entradas e saidas.
  - Funcoes principais: `build_entry_period_comparison_rows`, `build_sale_period_comparison_rows`.

- `xml_reconciliation.py`
  - Conciliacao SPED/XML.
  - Reconstrucao/ajuste de itens usando XML.
  - Comparativos PIS/COFINS por periodo.
  - Indices de itens fiscais XML.

- `product_import.py`
  - Importacao/candidatos de produtos por XML, SPED Fiscal, SPED Contribuicoes e SPED 0200.
  - Usado pelos fluxos de previa e importacao de produtos.

- `operation_summary.py`
  - Consolida operacoes, C190/C590, bases, reducoes, apuracao e detalhes por documento/produto.

- `monthly_reports.py`
  - Relatorios mensais por produto, operacao e aliquotas.
  - Inclui datasets lineares para exportacao.

- `credit_diagnostics.py`
  - Diagnosticos de credito/debito/base ICMS.
  - Comparativos por periodo e payloads de exportacao.

- `entry_exit_reports.py`
  - Analise consolidada de entradas e saidas.
  - Gera linhas e rodapes para relatorio.

- `adjusted_sped.py`
  - Regras de ajuste e geracao de SPED ajustado.
  - Recalculo de blocos, totais, E110/E210/E116 e logs de ajuste.

- `adjusted_sped_reports.py`
  - Exportacoes relacionadas ao SPED ajustado.
  - Exemplos: CST 061, CFOP 1252/1253, XML ST ausente.

- `compare_workflows.py`
  - Fluxos de comparacao SPED x XML e SPED x planilha.
  - Sumarios por CFOP, credito de entrada XML e coleta de XMLs.

- `compare_sped_launcher.py`
  - Insere notas XML no SPED de comparacao.
  - Garante registros 0150/0190/0200, C100/C170/C190 e recalcula blocos.

- `compare_operations.py`
  - Classifica operacao de XML/documento.
  - Normaliza escopo de comparacao e monta snapshot de nota.

- `compare_matching.py`
  - Regras auxiliares para casar XML/SPED por chave, numero e valores.

- `runtime_rules.py`
  - Parser e aplicacao de regras dinamicas.
  - Regras ICMS configuradas e padroes.
  - Matching e aplicacao em itens.

- `runtime_rule_history.py`
  - Carrega, salva, lembra e remove historico de regras dinamicas.

- `tax_rules.py`
  - Normalizacao fiscal comum.
  - CST, PIS/COFINS, aliquotas, formatacao decimal e operacao.

- `path_selection.py`
  - Normaliza selecao de caminhos.
  - Parse/format de multiplos arquivos, deduplicacao, limite, faltantes e colapso de XML.

- `app_paths.py`
  - Resolve caminhos base da aplicacao.
  - Caminhos de log/auditoria e historico de regras.

- `app_config_service.py`
  - Carrega e salva configuracao da aplicacao.

- `audit_utils.py`
  - Formata caminhos e resumos para auditoria/log.

- `analysis_utils.py`
  - Utilitarios compartilhados por analises.
  - Periodos, datas, filtros e curva ABC.

## Exporters

Pasta: `Sped/app/exporters/`

Responsabilidade: escrita de arquivos de saida.

- `excel_base.py`
  - Base comum para gerar XLSX.
  - Estilos, tipos de coluna, totalizadores, nomes de abas e XML interno do workbook.

- `workbook_exporter.py`
  - Exporta Excel e CSV simples.
  - Exporta workbook mensal colorido.
  - Serializa valores para clipboard/exportacao.

- `rules_report_exporter.py`
  - Gera relatorio Word das regras dinamicas.
  - Descreve condicoes, acoes e observacoes das regras.

## Repositories

Pasta: `Sped/app/repositories/`

- `mysql_cadastro.py`
  - Acesso ao MySQL para cadastro de produtos/empresas.
  - Mantem queries e operacoes de persistencia fora da UI.

## Arquivos auxiliares de raiz

- `Sped/mysql_schema.sql`
  - Estrutura SQL esperada para o banco.

- `Sped/explicacao_conciliacao_sped.md`
  - Documento explicativo existente sobre conciliacao SPED.

- `Sped/.gitignore`
  - Regras de ignorados do projeto.

## Fluxos principais e onde mexer

### Tela, botoes, menus e grids

Alterar:

- `app/ui/app.py`

Se a alteracao virar regra reaproveitavel, mover para `app/services/`.

### Progress dialog

Alterar:

- Estrutura/comportamento do handle: `app/ui/progress_dialog.py`
- Aparencia da janela: metodo `open_progress_dialog()` em `app/ui/app.py`
- Uso em fluxo: `with self.progress_dialog(...) as progress:`

Padrao atual:

```python
with self.progress_dialog("Titulo", "Mensagem inicial") as progress:
    servico(..., progress_callback=progress.update)
    progress.reset("Proxima etapa...")
```

Quando um fluxo chama outro e quer reaproveitar a mesma janela:

```python
self.outro_fluxo(external_progress=progress)
```

### Leitura de SPED/XML/Excel

Alterar:

- SPED geral: `app/parsers/sped_parser.py`
- SPED Fiscal detalhado: `app/parsers/sped_fiscal_parser.py`
- XML de comparacao: `app/parsers/compare_xml.py`
- Planilha de comparacao: `app/parsers/compare_sheet.py`
- XLSX generico/filtros: `app/parsers/excel_parser.py`

### Regra de negocio fiscal

Alterar:

- Regras fiscais comuns: `app/services/tax_rules.py`
- Regras dinamicas: `app/services/runtime_rules.py`
- Comparativos por periodo: `app/services/period_comparisons.py`
- Conciliacao SPED/XML: `app/services/xml_reconciliation.py`
- Ajuste de SPED: `app/services/adjusted_sped.py`

### Exportacao

Alterar:

- Excel/CSV: `app/exporters/workbook_exporter.py`
- Base XLSX/estilos: `app/exporters/excel_base.py`
- Word de regras: `app/exporters/rules_report_exporter.py`

### Banco de dados

Alterar:

- Queries e persistencia: `app/repositories/mysql_cadastro.py`
- Config padrao: `app/config.py`
- Config local: `app/ui/mysql_config.json`

### Historico e auditoria

Alterar:

- Caminhos: `app/services/app_paths.py`
- Auditoria: `app/services/audit_utils.py`
- Historico de regra dinamica: `app/services/runtime_rule_history.py`

## Convencao de responsabilidades

- `ui/`
  - Deve cuidar de tela, eventos, mensagens, selecao de arquivos e chamada dos servicos.

- `services/`
  - Deve conter calculo, regra fiscal, conciliacao, importacao e montagem de linhas/datasets.

- `parsers/`
  - Deve ler arquivos e devolver dados estruturados.

- `exporters/`
  - Deve gravar arquivos de saida.

- `repositories/`
  - Deve conversar com banco de dados.

- `models.py`
  - Deve conter dataclasses compartilhadas.

## Estado atual do refactor de progresso

O modelo de progress dialog ficou assim:

- `ProgressDialogHandle` mora em `app/ui/progress_dialog.py`.
- `SpedApp.open_progress_dialog()` cria os widgets Tkinter.
- `SpedApp.progress_dialog()` abre ou reutiliza uma janela e fecha automaticamente quando for dona dela.
- Fluxos usam `progress.update` como callback.
- Fluxos encadeados usam `progress.reset(...)` antes de chamar a proxima etapa.

Pontos ja convertidos:

- Previa/importacao de produtos.
- Consultas ICMS de entrada e saida.
- Consultas PIS/COFINS de entrada e saida.
- Atualizacao de telas apos regra dinamica.

As barras fixas embutidas nas telas de XML e comparacao continuam em `app.py`, porque nao sao dialogs reaproveitaveis.

