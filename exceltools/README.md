# Biblioteca ExcelTools
Para utilizar a biblioteca, copie o arquivo ExcelTools.ahk para dentro da pasta **Lib** do AutoHotkey (C:\Users\usuario\Documents\AutoHotkey\Lib) e inclua o seguinte comando na primeira linha do seu script antes de utilizar qualquer função:

**`#Include <ExcelTools>`**
___
# Funções
A biblioteca conta com 7 funções para serem usadas com o Excel. Dentre elas, considere obrigatório utilizar umas das 3 primeiras funções para "configurar a planilha a ser trabalhada", antes de utilizar as demais funções. Para apenas ver os exemplos na prática do uso da biblioteca, considere ir para a linha 135.

**Observação: Todas as funções devem ser utilizadas para trabalhar com apenas uma planilha de cada vez.**
___
## 1 - `abrir()`
É utilizada para trabalhar com uma planilha já existente no seu computador. Essa função recebe dois parâmetros separados por uma vírgula, dos quais apenas um é obrigatório. O primeiro parâmetro a ser passado, é o caminho completo da planilha que você deseja abrir (entre aspas duplas). Para pegar o caminho completo de uma planilha, basta clicar com o botão direito em cima dela, ir na opção "Propriedades", e em seguida clicar na aba "Segurança". Logo em cima você verá o caminho, do lado de "Nome do objeto". O segundo parâmetro da função (não obrigatório), é o valor 0 ou 1, onde 0 significa que a planilha que você vai trabalhar, ficará invisível (Exemplo 1) e 1 a planilha vai estar visível (Exemplo 2). Caso esse valor não seja passado para a função, ela assumirá o valor padrão de 1 (Exemplo 3) ou seja, ao ser aberta ficará visível.

#### Exemplo 1: Abrindo uma planilha invisível.

`abrir("C:\Users\usuario\Desktop\estoque.xlsx", 0)`

#### Exemplo 2: Abrindo uma planilha visível.

`abrir("C:\Users\usuario\Desktop\estoque.xlsx", 1)`

#### Exemplo 3: Abrindo uma planilha visível (Padrão).

`abrir("C:\Users\usuario\Desktop\estoque.xlsx")`
___
## 2 - `conectar()`

É utilizada para trabalhar com uma planilha ativa, ou seja, uma planilha que já está aberta na sua área de trabalho. Essa função não recebe nenhum parâmetro.

### Exemplo 1: Conectando em uma planilha aberta.

`conectar()`
___
## 3 - `criar()`

É utilizada para criar uma nova planilha. Essa função recebe apenas um parâmetro (não obrigatório). Se informado, o valor do parâmetro deve ser 0 ou 1, onde 0 significa que a planilha que você criar, ficará invisível (Exemplo 1) e 1 a planilha será visível (Exemplo 2). Caso esse valor não seja passado para a função, ela assumirá o valor padrão de 1 (Exemplo 3) ou seja, ao ser criada ficará visível.

#### Exemplo 1: Criando uma planilha invisível.

`criar(0)`

#### Exemplo 2: Criando uma planilha visível.

`criar(1)`

#### Exemplo 3: Criando uma planilha visível (Padrão).

`criar()`
___
## 4 - `escrever()`

É utilizada para adicionar um valor em uma célula da planilha. Essa função recebe dois parâmetros (obrigatórios), separados por vírgula. O primeiro parâmetro a ser informado (entre aspas duplas), é a célula que deseja inserir o valor. Você pode passar apenas uma célula (Exemplo 1) ou um intervalo de células (Exemplo 3). O segundo parâmetro é o valor que deseja inserir na célula. Caso o valor que deseja inserir seja numérico, não é necessário colocar entre aspas duplas (Exemplo 2).

#### Exemplo 1: Inserindo um valor comum na célula A1 (Primeira célula da primeira coluna).

`escrever("A1", "Gabriel")`

#### Exemplo 2: Inserindo um valor numérico na célula B1 (Primeira célula da segunda coluna).

`escrever("B2", 22)`

#### Exemplo 3: Inserindo um valor em um conjunto de células (O valor "SP" será adicionado ao mesmo tempo nas células "C1", "C2", "C3" e "C4").

`escrever("C1:C4", "SP")`
___
## 5 - `capturar()`

É utilizada para obter o valor de uma célula da planilha. Essa função recebe apenas um parâmetro (obrigatório). O parâmetro a ser informado (entre aspas duplas), é a célula que deseja obter o valor. Você pode exibir o valor de apenas uma célula (Exemplo 1 e 2) ou um conjunto de células (Exemplo 3 e 4). Ao informar um conjunto de células para ser exibido, é importante notar que as células e seus valores serão guardados em uma Matriz, isso é, um conjunto de Arrays. Portanto, antes de acessarmos qualquer desses valores, é preciso que salvemos esse conjunto capturado em uma variável. Em seguida, para acessar um valor específico do conjunto salvo, é preciso informar o nome da variável em que foi armazenada o conjunto, seguida dentro de colchetes, primeiro o número correspondente a célula, e depois o número correspondente a coluna. Tais valores devem ser separados por uma vírgula, como é mostrado nos Exemplos 3 e 4.

#### Exemplo 1: Exibindo o valor da célula A1 (Primeira célula da primeira coluna).

`MsgBox % capturar("A1")`

#### Exemplo 2: Salvando o valor da célula capturada em uma variável chamada "celula", para depois ser exibida.

```
celula := capturar("A1")
MsgBox % celula
```

#### Exemplo 3: Salvando um conjunto de células capturadas em uma variável e exibindo apenas o valor da célula A2 (Segunda célula da Primeira coluna).

```
conjunto := capturar("A1:B2")
MsgBox % conjunto[2, 1]
```

#### Exemplo 4: Salvando um conjunto de células capturadas em uma variável e exibindo apenas o valor da célula B1 (Primeira célula da Segunda coluna).

```
conjunto := capturar("A1:B2")
MsgBox % conjunto[1, 2]
```
___
## 6 - `salvar()`

É utilizada para salvar a planilha que está sendo trabalhada. Essa função recebe apenas um parâmetro (não obrigatório). Caso informado, o valor do parâmetro deve ser o caminho completo da planilha (Exemplo 1) ou passando o caminho completo da planilha, porém mudando apenas o nome (Exemplo 3), possibilitando desta forma ser criada e salva uma nova planilha, sem sobreescrever a atual.

#### Exemplo 1: Salvando a planilha atual informando o seu caminho completo (Mesma função que o Exemplo 2, considerando que a planilha atual tem o nome de "estoque").

`salvar("C:\Users\usuario\Desktop\estoque.xlsx")`

#### Exemplo 2: Salvando a planilha atual

`salvar()`

#### Exemplo 3: Salvando uma nova planilha sem sobreescrever a planilha atual (Considerando que a planilha atual tem o nome de "estoque").

`salvar("C:\Users\usuario\Desktop\novo-estoque.xlsx")`
___
## 7 - `sair()`

É utilizada para garantir que a conexão do Excel com as planilhas criadas ou abertas, foi encerrada corretamente. Essa função não recebe nenhum parâmetro. É sempre recomendado utilizar essa função em todo script, após terminar o uso da biblioteca.

#### Exemplo 1:

`sair()`
___
# Exemplos Práticos

#### Uso 1: Criando uma planilha visível, inserindo o valor "Exemplo" na célula "A1", salvando, e em seguida finalizando a conexão com o Excel.

```
#Include <ExcelTools>

criar()

escrever("A1", "Exemplo")

salvar("C:\Users\usuario\Desktop\uso-1.xlsx")

sair()
```
___
#### Uso 2: Abrindo uma planilha já existente, inserindo o valor "22" nas células "A1", "A2", "B1", "B2", salvando, e em seguida finalizando a conexão com o Excel.

```
#Include <ExcelTools>

abrir("C:\Users\usuario\Desktop\uso-2.xlsx")

escrever("A1:B2", 22)

salvar()

sair()
```
___
#### Uso 3: Conectando em uma planilha ativa (já aberta na Área de trabalho), salvando o valor capturado da célula "A1" em uma variável chamada "celula", exibindo o valor da célula salva e em seguida finalizando a conexão com o Excel.

```
#Include <ExcelTools>

conectar()

celula := capturar("A1")

MsgBox % celula

sair()
```
___
#### Uso 4: Abrindo uma planilha já existente (invisível), salvando o valor capturado do conjunto de células "A1", "A2", "B1", "B2" em uma variável chamada "conjunto", exibindo o valor da célula "B2" (Segunda célula da segunda coluna), e em seguida finalizando a conexão com o Excel.

```
#Include <ExcelTools>

abrir("C:\Users\usuario\Desktop\uso-4.xlsx", 0)

conjunto := capturar("A1:B2")

MsgBox % conjunto[2, 2]

sair()
```
___
#### Uso 5: Abrindo uma planilha já existente (invisível), salvando o valor capturado da célula "B2" em uma variável chamada "celula" e em seguida finalizando a conexão com o Excel. Depois criando uma planilha invisível, inserindo o valor salvo anteriormente da variável "celula" da outra planilha, na célula "A1" da nova planilha, salvando e em seguida finalizando a conexão com o Excel.

```
#Include <ExcelTools>

abrir("C:\Users\usuario\Desktop\uso-1.xlsx", 0)

celula := capturar("B2")

sair()

criar(0)

escrever("A1", celula)

salvar("C:\Users\usuario\Desktop\novo-uso-1.xlsx")

sair()
```
___
# Ficou com dúvidas?
### Entre em contato comigo:

\
*E-mail: automatetutoriais@hotmail.com*

*Telegram: [https://t.me/noslined047](https://t.me/noslined047)*

*Discord: automacoestutoriais#1959*
