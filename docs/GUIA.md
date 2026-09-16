# Gerador de cartas e comunicados — guia

Preenche um modelo do Word para cada linha da planilha, salva em PDF
individual, registra no log e prepara os e-mails no Outlook.

Mesma situação do conciliador: vem `.xlsx` com os módulos separados, porque
macro em arquivo baixado da internet é bloqueada pelo Office desde 2022 e eu
não consigo testar um `vbaProject.bin` sem Excel aqui. A montagem leva dois
minutos.

---

## Montagem

1. Abra `Gerador_Cartas.xlsx` e salve como `.xlsm`.
2. `Alt + F11`, e importe os três módulos (`Ctrl + M`):
   `vba/modUtil.bas`, `vba/modValidacao.bas`, `vba/modGerador.bas`.
3. Copie a pasta `modelos/` para onde quiser e aponte `PASTA_MODELOS` na aba
   Config.
4. Salve.

## Primeira rodada

| Passo | Macro | O que esperar |
|---|---|---|
| 1 | `ValidarDados` | Lista de problemas na aba Validacao |
| 2 | `GerarCartas` | 3 PDFs, porque `MODO_TESTE` vem em SIM |
| 3 | conferir os PDFs | Nenhum `{{MARCADOR}}` sobrando |
| 4 | `MODO_TESTE` = NAO, `GerarCartas` | As 40 do exemplo |
| 5 | `PrepararEmails` | Rascunhos no Outlook, nada enviado |

Deixe o modo teste em SIM até ver um PDF pronto. Gerar 300 cartas erradas e
descobrir depois é o modo padrão de falhar nesse tipo de automação.

---

## Como os marcadores funcionam

Cada coluna da aba Dados vira um marcador, em maiúsculas entre chaves duplas.
A coluna `nome` vira `{{NOME}}`. `salario_novo` vira `{{SALARIO_NOVO}}`. Os
parâmetros de Config viram marcadores do mesmo jeito.

**Acrescentar uma coluna cria um marcador novo sem tocar em código.** É isso
que faz o gerador servir para carta de promoção, de desligamento, de
transferência ou de qualquer outra coisa: você escreve um `.docx` novo, usa os
marcadores que quiser, cadastra na aba Modelos e pronto.

Três marcadores são calculados e não vêm de coluna nenhuma:

| Marcador | Conteúdo |
|---|---|
| `{{DATA_CARTA}}` | data por extenso, para o corpo |
| `{{DATA_GERACAO}}` | data e hora, para o rodapé |
| `{{REFERENCIA}}` | código único da carta, usado no log |

A formatação é por convenção de nome de coluna: qualquer coluna com "salario"
ou "valor" no nome sai como moeda, com "pct" sai como percentual, com "data"
sai como `dd/mm/aaaa`. Sem isso a carta imprime `8734.5599999` e
`01/04/2026 00:00:00`.

---

## As validações

Nove verificações, e essa é a parte que separa a ferramenta de uma mala
direta comum.

**Bloqueiam a geração (Erro)**

| O quê | Por quê |
|---|---|
| Campo obrigatório vazio | Carta com nome em branco |
| Matrícula repetida | A pessoa recebe duas cartas diferentes |
| Modelo não cadastrado ou arquivo ausente | Falha no meio do lote |
| E-mail com formato inválido | Rascunho que não sai da caixa |
| Novo salário menor ou igual ao atual | Carta de aumento que não aumenta |
| Percentual que não bate com os valores | Ver abaixo |
| Data inválida | Carta com data impossível |
| Marcador no `.docx` sem origem em Config ou Dados | Ver abaixo |

**Só avisam**: aumento acima do limite configurado, data de efeito com mais
de seis meses, linha sem e-mail.

### As duas validações que valem o projeto inteiro

**Percentual que não confere com os valores.** Se a linha diz 5% mas os
salários dizem 12%, alguém copiou e colou de outra linha. Quem confere a carta
pronta olha o percentual, não refaz a conta — então esse erro passa direto e
chega assinado na mão da pessoa.

**Marcador órfão no modelo.** O código abre cada `.docx`, extrai todo
`{{ALGUMA_COISA}}` e confere se existe coluna ou parâmetro correspondente. Sem
isso a carta sai com `{{GESTOR}}` impresso, e o erro só aparece depois da
assinatura.

---

## Decisões de desenho

**Uma instância do Word para o lote inteiro.** Abrir e fechar o Word a cada
carta multiplica o tempo por dez e é o erro mais comum em mala direta caseira.

**Substituição também em cabeçalho e rodapé.** `doc.Content` cobre o corpo e
as tabelas, mas não o timbre. Faltar com isso deixa `{{EMPRESA}}` impresso no
topo de todas as cartas.

**Validação roda de novo dentro de `GerarCartas`.** Mesmo que você tenha
validado antes: entre uma coisa e outra alguém pode ter editado a planilha.

**Coluna de verificação no log.** É um FNV-1a de 32 bits sobre o texto final
do documento, e eu chamo de verificador, não de hash — não é criptográfico e
não deve ser apresentado como se fosse. Serve para uma coisa: provar, meses
depois, que o PDF em mãos é o mesmo que o log registrou. Numa discussão sobre
"não foi isso que me mandaram", é a única evidência que existe.

**`PrepararEmails` cria rascunho e para.** Nada é enviado. Envio automático de
comunicado de remuneração é o tipo de automação que só dá errado uma vez, e o
erro não tem desfazer.

---

## Criar um modelo novo

1. Copie `Modelo_Merito.docx`, renomeie e edite o texto no Word.
2. Use `{{MARCADOR}}` onde entra conteúdo variável. Maiúsculas.
3. Se precisar de um campo que não existe, crie a coluna na aba Dados.
4. Cadastre a linha na aba Modelos, com o padrão de nome do arquivo de saída.
5. Rode `ValidarDados`: ela avisa se algum marcador do modelo ficou órfão.

## Limites conhecidos

Só roda em Windows com Word instalado. Automação COM não existe no Mac nem no
Excel Online, e não há contorno.

Com mais de mil cartas, considere fechar e reabrir o Word a cada duzentas: ele
acumula memória em lotes longos.

O verificador trunca em 20 mil caracteres, o que caracteriza qualquer carta de
uma ou duas páginas. Para documento longo, suba o limite em `modUtil`.
