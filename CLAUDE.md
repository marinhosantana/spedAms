# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Revisor SPED** is a Windows desktop application for reading, analyzing, reconciling, comparing, and exporting Brazilian tax document data (SPED/XML/spreadsheets). It is a Python application using **PySide6/Qt** as the UI framework.

> The legacy Tkinter UI was archived to `Sped/app/ui_legacy/` and is no longer active. Do not add new code there.

## Commands

### Run

```powershell
# Development
.\scripts\run-dev.ps1
# or
$env:SPED_ENV = "dev"; python .\Sped\NovoRevisorQt.py

# Production
.\scripts\run-prod.ps1
```

### Build EXE

```powershell
.\scripts\build-dev.ps1   # Output: dist\RevisorSPED_DEV\RevisorSPED_DEV.exe
.\scripts\build-prod.ps1  # Output: dist\RevisorSPED_PROD\RevisorSPED_PROD.exe
```

### Environment Setup

```powershell
.\scripts\setup-ambientes.ps1  # Creates .venv-dev/ and .venv-prod/
```

There is no test suite; testing is done manually via the UI.

## Architecture

The codebase follows a strict layered architecture under `Sped/app/`:

```
NovoRevisorQt.py → QtSpedApp (ui_qt/app.py)
Revisor.py       → QtSpedApp (ui_qt/app.py)   ← used by build_entries/
main.py          → Revisor.py
                        ↓
                  services/      ← business logic
                  parsers/       ← file readers (SPED, XML, Excel)
                  exporters/     ← file writers (Excel, CSV, Word)
                  repositories/  ← MySQL access
                  models.py      ← shared dataclasses
                  config.py      ← constants, namespaces, DB defaults
```

**Entry points:**
- Primary: `Sped/NovoRevisorQt.py` → `app/ui_qt/app.py`
- Via chain: `Sped/main.py` → `Sped/Revisor.py` → `app/ui_qt/app.py`
- PyInstaller builds: `Sped/build_entries/dev.py` and `Sped/build_entries/prod.py` → `Revisor.py`
- Legacy (archived, do not use): `Sped/app/ui_legacy/app.py`

### Layers

- **`ui_qt/`** — PySide6 event handlers, dialogs, data grids. Delegates all logic to services.
- **`ui_legacy/`** — Archived Tkinter UI. For reference only; not imported anywhere.
- **`services/`** — Pure business logic: fiscal reconciliation, comparisons, report generation, SPED adjustments, dynamic tax rules engine.
- **`parsers/`** — File readers: `sped_fiscal_parser.py` (SPED registers C100/C190/D100 etc.), `compare_xml.py` (NFe/NFSe/NFCe), `excel_parser.py` (XLSX).
- **`exporters/`** — File writers: `workbook_exporter.py` (Excel/CSV), `rules_report_exporter.py` (Word).
- **`repositories/`** — `mysql_cadastro.py`: all MySQL queries, no ORM (raw `mysql-connector-python`).
- **`models.py`** — Dataclasses shared across all layers (`ProductRecord`, `CompareXmlItem`, `CompareXmlInvoice`, etc.).
- **`config.py`** — Global constants, XML namespaces, MySQL defaults.

### Environment & Configuration

`SPED_ENV` controls which config files are loaded (`dev` or `prod`):
- `app/ui_legacy/app_config.{dev,prod}.json` — legacy; Qt reads config via `app_config_service.py`
- `app/ui_legacy/mysql_config.{dev,prod}.json` — MySQL connection (host, port, user, password, database)

Databases: `sped_icms_dev` (dev) and `sped_icms` (prod). Schema is in `Sped/mysql_schema.sql`.

Key tables: `sped_perfis` (company profiles), `sped_arquivos` (SPED archive with SHA256), `sped_produtos_0200` (product catalog with NCM/CEST/CST/rates).

## Domain Concepts

- **SPED** — Brazilian government tax file format (plain text registers: C100 invoice header, C190 tax summary by CST/CFOP/rate, D100/D590 service entries, E110/E116 apportionment).
- **NFe/NFSe/NFCe** — Electronic invoice XML formats.
- **CFOP** — Fiscal operation code identifying the nature of a transaction.
- **CST ICMS / PIS / COFINS** — Tax classification codes determining tax treatment.
- **Conciliation** — Cross-referencing SPED ↔ XML ↔ spreadsheets to detect discrepancies.
- **Adjustment** — Overwriting SPED registers with corrected data from XML or manual rules.
- **Dynamic rules engine** — Runtime-configurable rules (`runtime_rules.py`, `runtime_rule_history.py`) that define tax behavior per product/operation without code changes.

## Key Conventions

- Services never import UI code; UI imports services.
- Parsers return structured Python objects (dataclasses from `models.py`).
- MySQL queries live exclusively in `repositories/mysql_cadastro.py`.
- Long operations use `QThread` + worker objects with `progress` / `finished` / `failed` signals.
- Domain documentation: `Sped/explicacao_conciliacao_sped.md` explains the reconciliation logic in detail.
