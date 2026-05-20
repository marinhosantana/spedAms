# Explicacao da Conciliacao SPED x XML

## Resumo

O status **CONCILIADO** significa que a diferenca final ficou zerada.

Neste caso:

```text
Diferenca XML x SPED:        1.058,21
Diferenca C100 x C190:      -1.379,32
Diferenca D500/D590:           321,11
Diferenca C500/C590:             0,00
--------------------------------------
Diferenca final:                0,00
```

Ou seja:

```text
1.058,21 - 1.379,32 + 321,11 = 0,00
```

A operacao esta batendo no total geral. Existem diferencas internas, mas elas se compensam.

## O Que Significa Cada Parte

### 1. XML x SPED: 1.058,21

Esse valor vem de notas em que o valor do XML nao bate com o valor do documento no SPED, ou seja, o `vNF` do XML esta diferente do `VL_DOC` do registro `C100`.

Dentro desse grupo:

```text
Provavel ST no SPED:                         1.094,33
Provavel IPI do XML nao levado ao VL_DOC:      -36,12
------------------------------------------------------
Total XML x SPED:                            1.058,21
```

Leitura:

- quando o SPED esta maior que o XML, pode haver ST lancada no SPED;
- quando o SPED esta menor que o XML, pode haver IPI no XML que nao foi levado ao `VL_DOC` do SPED.

### 2. C100 x C190: -1.379,32

Esse valor vem de diferencas dentro do proprio SPED.

O `C100` representa o valor do documento fiscal.

O `C190` representa o resumo fiscal por CST, CFOP e aliquota.

Eles podem ser diferentes porque o `C100` pode considerar desconto, ST, IPI, frete, seguro ou outras despesas, enquanto o `C190` pode resumir a operacao por outra base.

Composicao:

```text
Descontos / resumo maior que documento:          638,43
Acrescimos / impostos / resumo menor:         -2.017,75
-------------------------------------------------------
Total C100 x C190:                            -1.379,32
```

Leitura:

- valor positivo normalmente indica desconto no documento;
- valor negativo normalmente indica ST, IPI, outros acrescimos ou composicao fiscal que nao entra no `C190` da mesma forma.

### 3. D500/D590: 321,11

Esse valor vem de documentos do bloco D.

Esses registros entram no total fiscal da tela, mas nao fazem parte da comparacao principal de XML de notas `C100`.

Por isso eles aparecem como um componente separado da conciliacao.

### 4. C500/C590: 0,00

Neste caso, nao houve impacto de `C500/C590`.

Se houver valor nessa linha em outro arquivo, significa que documentos de energia/comunicacao do bloco C tambem estao entrando no total fiscal.

## Conclusao Para o Usuario

A diferenca final e **0,00**, portanto a operacao esta conciliada.

Existem diferencas internas entre XML, `C100`, `C190` e registros do bloco D, mas elas estao totalmente explicadas.

O ponto principal e:

```text
Diferenca original:    0,00
Diferenca explicada:   0,00
Diferenca pendente:    0,00
Status: CONCILIADO
```

Isso nao significa que todos os documentos sao identicos linha a linha. Significa que todas as diferencas encontradas foram explicadas e, no total geral, nao sobrou valor pendente.

