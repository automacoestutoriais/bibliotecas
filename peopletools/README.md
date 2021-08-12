# Biblioteca PeopleTools
Para utilizar a biblioteca, copie o arquivo PeopleTools.ahk para dentro da pasta **Lib** do AutoHotkey (C:\Users\usuario\Documents\AutoHotkey\Lib) e inclua o seguinte comando na primeira linha do seu script antes de utilizar qualquer função:

**`#Include <PeopleTools>`**
___
# Sobre
Essa biblioteca faz a geração de dados aleatórios de uma pessoa, sendo essa geração feita através de uma conexão direta com o site *4Devs - [https://www.4devs.com.br/gerador_de_pessoas](https://www.4devs.com.br/gerador_de_pessoas)*. Vale ressaltar as seguintes informações referente ao uso dessa ferramenta que podem ser encontrados no site:

_IMPORTANTE: Nosso gerador online de Pessoa tem como intenção ajudar estudantes, programadores, analistas e testadores a gerar documentos. Normalmente necessários parar testar seus softwares em desenvolvimento.
A má utilização dos dados aqui gerados é de total responsabilidade do usuário.
Os números são gerados de forma aleatória, respeitando as regras de criação de cada documento._
_Todos os conteúdos e dados deste site são apenas para fins informativos, não devem ser considerados completos, atualizados, e não se destinam a ser utilizado no lugar de uma consulta jurídica, médica, financeira, ou de qualquer outro profissional. Os conteúdos são fornecidos sem qualquer tipo de garantia. Todo e qualquer risco da utilização dos conteúdos é assumido pelo próprio usuário._
___
# Como utilizar
Para gerar um dado, basta fazer como no exemplo abaixo. Para consultar quais dados podem ser gerados e qual o nome do atributo, basta usar o comando `pessoa.ajuda()`.

#### Exemplo 1: Exibindo um nome.

`MsgBox % pessoa.nome`

#### Exemplo 2: Exibindo uma cidade.

`MsgBox % pessoa.cidade`

#### Exemplo 3: Exibindo 10 nomes diferentes.

*Observação do Exemplo 3:* Lembrando que, ao incluir a biblioteca pela primeira vez, utilizando o comando *#Include <PeopleTools>*, automaticamente o código de geração de pessoa(PeopleTools) é executado, desta forma, quando exibindo um dado, como no Exemplo 1, caso o código seja repetido, será exibido o mesmo dado. Para gerar um dados diferentes no mesmo código, como no Exemplo 3, é necessário antes remover o comando *#Include <PeopleTools>* da primeira linha.

```
Loop 10
{
    #IncludeAgain <PeopleTools>
    MsgBox % pessoa.nome
}
```
___
# Ficou com dúvidas?
### Entre em contato comigo:

\
*E-mail: automatetutoriais@hotmail.com*

*Telegram: [https://t.me/noslined047](https://t.me/noslined047)*

*Discord: automacoestutoriais#1959*
