# Como usar ambientes de desenvolvimento e producao

Este projeto agora suporta dois ambientes:

- `dev`: desenvolvimento e testes.
- `prod`: producao, para uso estavel.

O ambiente e escolhido pela variavel `SPED_ENV`.

## Arquivos de configuracao

Os arquivos ficam em `Sped/app/ui/`:

- `app_config.dev.json`: titulo/nome da aplicacao em desenvolvimento.
- `app_config.prod.json`: titulo/nome da aplicacao em producao.
- `mysql_config.dev.json`: conexao MySQL de desenvolvimento.
- `mysql_config.prod.json`: conexao MySQL de producao.

Por seguranca, o banco padrao do desenvolvimento e:

```text
sped_icms_dev
```

O banco padrao da producao e:

```text
sped_icms
```

Assim voce pode testar alteracoes sem mexer no banco real.

## Como rodar em desenvolvimento

No PowerShell, na pasta `c:\spedAms`, execute:

```powershell
.\run-dev.ps1
```

Esse script define:

```powershell
$env:SPED_ENV = "dev"
```

e abre o sistema usando os arquivos `*.dev.json`.

## Como rodar em producao

No PowerShell, na pasta `c:\spedAms`, execute:

```powershell
.\run-prod.ps1
```

Esse script define:

```powershell
$env:SPED_ENV = "prod"
```

e abre o sistema usando os arquivos `*.prod.json`.

## Como conferir o ambiente atual

Na aba de configuracao da aplicacao, o status mostra o ambiente atual:

```text
Ambiente atual: dev
```

ou:

```text
Ambiente atual: prod
```

No ambiente `dev`, o titulo da janela tambem vem com `[DEV]`.

## Ambientes virtuais Python

Crie um ambiente virtual para desenvolvimento:

```powershell
cd c:\spedAms
python -m venv .venv-dev
.\.venv-dev\Scripts\activate
pip install -r requirements.txt
```

Crie outro para producao:

```powershell
cd c:\spedAms
python -m venv .venv-prod
.\.venv-prod\Scripts\activate
pip install -r requirements.txt
```

Quando estiver trabalhando no codigo, use a `.venv-dev`.

Quando for validar uma versao estavel ou gerar executavel, use a `.venv-prod`.

## Fluxo recomendado com Git

Use a branch `develop` para trabalho do dia a dia:

```powershell
git checkout -b develop
```

Depois de testar tudo, envie para a branch de producao:

```powershell
git checkout main
git merge develop
```

Regra pratica:

- `develop`: alteracoes novas, testes, ajustes.
- `main`: versao estavel de producao.

## Configuracao manual sem script

Se nao quiser usar os scripts, rode manualmente:

```powershell
$env:SPED_ENV = "dev"
python .\Sped\main.py
```

ou:

```powershell
$env:SPED_ENV = "prod"
python .\Sped\main.py
```

Se `SPED_ENV` nao for informado, ao rodar pelo codigo-fonte o sistema usa `dev` por padrao.
