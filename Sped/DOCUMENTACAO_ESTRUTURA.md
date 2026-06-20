# Documentacao da estrutura do sistema

Este documento resume como o projeto esta organizado, onde fica cada responsabilidade e quais arquivos devem ser alterados para cada tipo de manutencao.

## Visao geral

O sistema e uma aplicacao desktop em Python com interface Qt/PySide6 para leitura, analise, conciliacao, comparacao, ajuste e exportacao de dados SPED/XML/planilhas.

A interface Tkinter antiga foi arquivada em `Sped/app/ui_legacy/`. Ela pode servir como referencia historica, mas manutencoes novas devem ser feitas na interface Qt e nas camadas de servico/parser/exportacao.

Fluxo geral:

1. `main.py` chama `Revisor.py`.
2. `Revisor.py` cria a aplicacao Qt e instancia `QtSpedApp`.
3. `app/ui_qt/app.py` monta a interface, recebe a acao do usuario e chama servicos.
4. `app/services/` executa as regras de negocio.
5. `app/parsers/` le arquivos SPED, XML e Excel.
6. `app/exporters/` grava Excel, CSV e Word.
7. `app/repositories/` acessa banco de dados.

## Entrada da aplicacao

- `Sped/main.py`
  - Ponto de entrada simples.
  - Importa `main` de `Revisor.py`.

- `Sped/NovoRevisorQt.py`
  - Ponto de entrada direto da interface Qt.
  - Cria `QApplication`.
  - Instancia `QtSpedApp`.
  - Inicia `app.exec()`.

- `Sped/Revisor.py`
  - Ponto de entrada compativel com os builds.
  - Cria ou reutiliza `QApplication`.
  - Instancia `QtSpedApp`.
  - Inicia `app.exec()`.

## UI

- `Sped/app/ui_qt/app.py`
  - Arquivo principal da interface.
  - Contem a classe `QtSpedApp`.
  - Responsavel por telas, botoes, grids, menus, filtros, popups, mensagens e chamadas aos servicos.
  - Deve orquestrar fluxos, mas nao deve concentrar regra de negocio nova quando ela puder morar em `services/`.

- `Sped/app/ui_qt/confronto_dialog.py`
  - Dialogos e telas auxiliares do confronto SPED/cadastro.

- `Sped/app/ui_qt/import_planilha.py`
  - Fluxos de importacao de planilhas na interface Qt.

- `Sped/app/ui_legacy/`
  - Interface Tkinter arquivada.
  - Nao adicionar funcionalidade nova aqui sem decisao explicita.

- Arquivos `app_config.*.json`
  - Configuracao visual/local da aplicacao, como titulos.

- Arquivos `mysql_config.*.json`
  - Configuracao local de conexao MySQL.

- Pastas `logs/` e `storage/`
  - Guardam logs/auditoria e arquivos preservados em runtime.
  - Ficam fora do Git.

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
  - Utilitarios de produtos/NCM/CEST usados por conciliacoes e relatorios fiscais.

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
  - Acesso ao MySQL para configuracao e arquivamento de SPEDs originais.
  - Mantem queries e operacoes de persistencia fora da UI.

## Arquivos auxiliares de raiz

- `Sped/mysql_schema.sql`
  - Estrutura SQL esperada para o banco.

- `Sped/explicacao_conciliacao_sped.md`
  - Documento explicativo existente sobre conciliacao SPED.

- `.gitignore` e `Sped/.gitignore`
  - Regras de ignorados do projeto, incluindo ambientes virtuais, builds, logs, storage, planilhas, dumps SQL e arquivos fiscais locais.

## Fluxos principais e onde mexer

### Tela, botoes, menus e grids

Alterar:

- `app/ui_qt/app.py`
- `app/ui_qt/confronto_dialog.py`
- `app/ui_qt/import_planilha.py`

Se a alteracao virar regra reaproveitavel, mover para `app/services/`.

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
- Config local/runtime: arquivos `mysql_config.*.json`

### Historico e auditoria

Alterar:

- Caminhos: `app/services/app_paths.py`
- Auditoria: `app/services/audit_utils.py`
- Historico de regra dinamica: `app/services/runtime_rule_history.py`

## Convencao de responsabilidades

- `ui_qt/`
  - Deve cuidar de tela, eventos, mensagens, selecao de arquivos e chamada dos servicos.

- `ui_legacy/`
  - Arquivo historico da interface Tkinter. Nao e a UI principal.

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

## Cuidados de governanca

- Nao versionar planilhas, XMLs, SPEDs, dumps SQL, backups ou exports de cliente.
- Antes de commit, revisar `git status --short` e confirmar que so ha codigo/documentacao intencional.
- Validar alteracoes com `python -m compileall -q Sped` e, quando aplicavel, testes manuais pela UI Qt em ambiente `dev`.
- Usar `dev` para desenvolvimento e `prod` apenas para validacao/uso estavel.

