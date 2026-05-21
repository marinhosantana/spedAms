# Arquivamento do SPED original

Este documento descreve o primeiro passo implementado para preservar o SPED original sem alteracoes.

## Objetivo

Quando um arquivo SPED e importado/selecionado no sistema, o app registra uma copia imutavel para auditoria futura.

Esse registro guarda:

- arquivo original copiado para uma pasta controlada;
- hash SHA-256 do arquivo;
- nome original;
- nome da empresa identificado no registro `0000`, quando possivel;
- tamanho do arquivo;
- ambiente (`dev` ou `prod`);
- tipo de SPED (`fiscal` ou `contribuicoes`);
- CNPJ/CPF identificado no registro `0000`, quando possivel;
- periodo inicial e final identificado no registro `0000`, quando possivel;
- perfil automatico de auditoria.

O fluxo recomendado agora e:

```text
Configuracoes > SPEDs Arquivados > Importar SPED
```

Essa deve ser a primeira acao para cadastrar um SPED original no sistema.

## Pasta de armazenamento

Os arquivos originais sao copiados para:

```text
Sped/storage/original_speds/
```

Essa pasta fica fora do Git.

## Tabelas criadas

O arquivo `mysql_schema.sql` agora cria:

```text
sped_perfis
sped_arquivos
sped_produtos_0200
sped_documentos
sped_itens_c170
sped_resumos_c190
```

`sped_perfis` guarda o agrupamento/perfil de auditoria por empresa.
O perfil automatico usa o nome e o CNPJ identificados no SPED, por exemplo:

```text
EMPRESA TESTE LTDA - 12345678000199
```

Dentro do mesmo perfil podem existir varios arquivos importados, por data e por tipo:

- SPED Fiscal de janeiro;
- SPED Fiscal de fevereiro;
- SPED Contribuicoes de janeiro;
- SPED Contribuicoes de fevereiro.

`sped_arquivos` guarda cada SPED original arquivado dentro do perfil.

As tabelas analiticas guardam os dados extraidos do SPED Fiscal:

- `sped_produtos_0200`: produtos do registro 0200.
- `sped_documentos`: documentos identificados a partir dos itens e resumos.
- `sped_itens_c170`: itens fiscais do registro C170.
- `sped_resumos_c190`: resumos fiscais do registro C190.

## Identificacao de duplicidade

O sistema calcula o hash SHA-256 do arquivo.

Se o mesmo arquivo for selecionado novamente no mesmo ambiente, ele nao duplica o cadastro. O sistema apenas registra no log que o SPED original ja estava cadastrado.

## Comportamento se o MySQL estiver indisponivel

Se o MySQL nao estiver disponivel, o sistema nao interrompe o uso normal.

Nesse caso, ele tenta manter a copia local do arquivo original e escreve o evento no log de auditoria.

## Pontos integrados nesta etapa

O arquivamento automatico foi ligado aos principais seletores de SPED:

- SPED Fiscal principal;
- SPED Contribuicoes principal;
- SPEDs de consulta;
- SPEDs de comparacao;
- SPED usado na tela XML;
- SPEDs da analise de entrada e saida.

## Persistencia dos dados extraidos

Depois de arquivar o SPED Fiscal original, o sistema le o arquivo e persiste os dados extraidos em tabelas proprias.

Esse processo esta ligado ao `sped_arquivos.id`, permitindo consultar no futuro:

- quais produtos existiam no SPED original;
- quais documentos estavam no arquivo;
- quais itens C170 foram lidos;
- quais resumos C190 faziam parte da apuracao.

Nesta etapa, a persistencia analitica foi aplicada ao SPED Fiscal. SPED Contribuicoes segue arquivado como arquivo original e pode ganhar tabelas proprias em uma etapa posterior.

## Consulta e reuso dos perfis

A fase 3 adicionou a tela de consulta dos arquivos arquivados.

Ela fica em:

```text
Configuracoes > SPEDs Arquivados
```

Nessa tela e possivel:

- importar SPEDs para criar/atualizar perfis de auditoria;
- atualizar a lista de perfis;
- ver os SPEDs originais vinculados ao perfil;
- conferir contagens de produtos 0200, documentos, itens C170 e resumos C190;
- abrir a pasta onde o arquivo original arquivado foi salvo;
- enviar um SPED Fiscal arquivado para a consulta fiscal.

A fase 4 adicionou o botao `Carregar Perfil`.

Ao carregar um perfil, o sistema:

- envia os SPEDs Fiscais do perfil para as consultas fiscais de entrada e saida;
- envia os SPEDs Contribuicoes para as consultas PIS/COFINS de entrada e saida;
- preenche a tela de comparacao quando existir exatamente um SPED Fiscal ou um SPED Contribuicoes no perfil;
- usa sempre o arquivo original arquivado, preservado na pasta `Sped/storage/original_speds/`.

## Proxima etapa sugerida

A proxima etapa natural e criar consultas historicas diretamente sobre as tabelas analiticas, sem reler o TXT quando o usuario quiser apenas visualizar dados ja arquivados.
