# gerador-cartas

Gerador de cartas e comunicados de RH em lote. Preenche um modelo do Word para
cada linha da planilha, salva em PDF individual, registra no log e prepara os
e-mails no Outlook.

Dados de demonstração vindos de
[hr-synthetic-data-br](../hr-synthetic-data-br).

## Começar

1. Abra `Gerador_Cartas.xlsx` e salve como `.xlsm`
2. `Alt + F11`, importe `vba/modUtil.bas`, `vba/modValidacao.bas`,
   `vba/modGerador.bas`
3. Aponte `PASTA_MODELOS` na aba Config para a pasta `modelos/`
4. `Alt + F8` > `ValidarDados`
5. `Alt + F8` > `GerarCartas` — gera 3 cartas, porque `MODO_TESTE` vem em SIM

Passo a passo completo em [`docs/GUIA.md`](docs/GUIA.md).

## A parte que importa não é gerar, é validar

Mala direta comum não valida nada: gera as 300 cartas e o problema aparece
quando alguém responde o e-mail. São nove verificações, e duas justificam o
projeto sozinhas.

**Percentual que não bate com os valores.** Se a linha diz 5% mas os salários
dizem 12%, alguém copiou de outra linha. Quem confere a carta pronta olha o
percentual, não refaz a conta — então esse erro chega assinado na mão da
pessoa.

**Marcador órfão no modelo.** O código abre cada `.docx`, extrai todo
`{{ALGUMA_COISA}}` e confere se existe coluna ou parâmetro correspondente. Sem
isso a carta sai com `{{GESTOR}}` impresso.

## Marcadores

Cada coluna da aba Dados vira um marcador, em maiúsculas entre chaves duplas.
`nome` vira `{{NOME}}`, `salario_novo` vira `{{SALARIO_NOVO}}`. Os parâmetros
de Config viram marcadores do mesmo jeito.

**Acrescentar uma coluna cria um marcador novo sem tocar em código.** É o que
faz o mesmo gerador servir para promoção, desligamento ou transferência.

A formatação segue convenção de nome de coluna: "salario" ou "valor" sai como
moeda, "pct" como percentual, "data" como `dd/mm/aaaa`.

## Decisões de desenho

**Uma instância do Word para o lote inteiro.** Abrir e fechar a cada carta
multiplica o tempo por dez.

**Substituição também em cabeçalho e rodapé.** `doc.Content` cobre o corpo e as
tabelas, mas não o timbre.

**`PrepararEmails` cria rascunho e para.** Nada é enviado. Envio automático de
comunicado de remuneração é o tipo de automação que só dá errado uma vez, e o
erro não tem desfazer.

**Coluna de verificação no log.** É um FNV-1a de 32 bits sobre o texto final —
verificador, não hash criptográfico, e não deve ser apresentado como tal. Serve
para provar, meses depois, que o PDF em mãos é o mesmo que o log registrou.

## Limites

Só roda em Windows com Word instalado. Automação COM não existe no Mac nem no
Excel Online, e não há contorno.

## Estrutura

```
Gerador_Cartas.xlsx   planilha de controle com 40 linhas de exemplo
modelos/              dois modelos do Word com marcadores
vba/                  três módulos para importar
build/                scripts que geraram a planilha e os modelos
docs/GUIA.md          montagem, criação de modelos, manutenção
```
