# Como usar ambientes de desenvolvimento e producao

Este projeto agora suporta dois ambientes:

- `dev`: desenvolvimento e testes.
- `prod`: producao, para uso estavel.

O ambiente e escolhido pela variavel `SPED_ENV`.

## Organizacao dos arquivos de ambiente

Os arquivos de apoio ficam separados para nao misturar com o codigo principal:

- `scripts/`: comandos PowerShell para rodar, preparar ambientes e gerar EXE.
- `requirements/`: dependencias Python separadas por finalidade.
- `Sped/build_entries/`: pontos de entrada usados apenas pelo PyInstaller.
- `dist/`: saida gerada dos executaveis.

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
.\scripts\run-dev.ps1
```

Esse script define:

```powershell
$env:SPED_ENV = "dev"
```

e abre o sistema usando os arquivos `*.dev.json`.

## Como rodar a nova interface em desenvolvimento

A interface atual continua existindo. A nova interface em reconstrucao abre separadamente:

```powershell
.\scripts\run-next-dev.ps1
```

Ela usa o mesmo ambiente `dev`, mas roda por:

```text
Sped/NovoRevisor.py
```

Os arquivos da nova interface ficam em:

```text
Sped/app/ui_next/
```

Nesta primeira etapa, a nova interface ja possui:

- Dashboard;
- SPEDs Arquivados;
- Consultas Fiscais;
- Configuracoes.

Em `Consultas Fiscais`, a tela funciona como um menu de consultas. Cada card abre uma tela propria para evitar misturar processamento, filtros e grades diferentes no mesmo espaco.

A tela `Consulta Entradas ICMS` ja traz a migracao inicial da antiga consulta de entradas em tela separada: selecao de SPEDs/XMLs, filtros por periodo, CST, CFOP, status e busca, grade por produto/periodo, totais no rodape, exportacao CSV e popups de Entradas, Diagnostico de Credito, Comparacao de Diagnostico, Curva ABC, Reducao BC, Apuracao e Espelho Docs.

## Como rodar em producao

No PowerShell, na pasta `c:\spedAms`, execute:

```powershell
.\scripts\run-prod.ps1
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
pip install -r requirements\base.txt
```

Crie outro para producao:

```powershell
cd c:\spedAms
python -m venv .venv-prod
.\.venv-prod\Scripts\activate
pip install -r requirements\base.txt
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

## Como gerar EXE de desenvolvimento

Use este comando na pasta `c:\spedAms`:

```powershell
.\scripts\build-dev.ps1
```

O executavel sera gerado em:

```text
dist\RevisorSPED_DEV\RevisorSPED_DEV.exe
```

Esse EXE abre sempre em ambiente `dev`.

## Como gerar EXE de producao

Use este comando na pasta `c:\spedAms`:

```powershell
.\scripts\build-prod.ps1
```

O executavel sera gerado em:

```text
dist\RevisorSPED_PROD\RevisorSPED_PROD.exe
```

Esse EXE abre sempre em ambiente `prod`.

## Diferenca pratica entre os dois EXEs

- `RevisorSPED_DEV.exe`: usa configuracoes `dev`, titulo com `[DEV]` e banco padrao `sped_icms_dev`.
- `RevisorSPED_PROD.exe`: usa configuracoes `prod` e banco padrao `sped_icms`.

Os scripts de build tambem copiam `mysql_schema.sql` para a pasta do executavel, porque o sistema usa esse arquivo para criar/atualizar o schema do MySQL.

## Onde configurar o MySQL do EXE

Ao abrir cada EXE pela primeira vez, ele cria o arquivo de configuracao JSON na mesma pasta do executavel, se ainda nao existir.

Exemplos:

```text
dist\RevisorSPED_DEV\mysql_config.dev.json
dist\RevisorSPED_PROD\mysql_config.prod.json
```

Voce tambem pode configurar pela tela `Configuracoes > Conexao MySQL`.

## Como deixar outro computador igual

No computador principal, primeiro envie as alteracoes para o Git:

```powershell
cd c:\spedAms
git status
git add .
git commit -m "Configura ambientes dev e prod"
git push origin develop
```

No outro computador, entre na pasta do projeto e baixe as alteracoes:

```powershell
cd c:\spedAms
git fetch origin
git checkout develop
git pull origin develop
```

Depois prepare os ambientes Python:

```powershell
.\scripts\setup-ambientes.ps1
```

Esse script cria, se ainda nao existirem:

```text
.venv-dev
.venv-prod
```

e instala as dependencias de uso e de build.

Depois teste:

```powershell
.\scripts\run-dev.ps1
```

e:

```powershell
.\scripts\run-prod.ps1
```

Se o outro computador ainda nao permitir executar scripts PowerShell, rode uma vez:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

## Arquivos que nao vao pelo Git

Alguns arquivos sao locais de cada computador e nao devem ser versionados:

- `.venv-dev`
- `.venv-prod`
- `build`
- `dist`
- arquivos `.spec`
- arquivos JSON de configuracao ignorados pelo `.gitignore` de `Sped`

Por isso, em cada computador voce deve configurar o MySQL pela tela `Configuracoes > Conexao MySQL`, ou deixar o sistema criar os JSON padrao na primeira abertura.
