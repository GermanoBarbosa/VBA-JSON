# hJsonBag

**hJsonBag** é uma classe para **parse** e **serialização** de JSON escrita em **Visual Basic 6.0**.

Ela permite:
- Ler JSON em objetos VB6.
- Escrever objetos VB6 em JSON.
- Suporte a aspas simples e duplas em strings.
- Suporte a valores nulos (`None`).
- Métodos úteis como `getPath`, `toArray`, `fromFile`.

Esta é uma adaptação baseada no trabalho original de **Robert D. Riemersma Jr.**, sob licença Apache 2.0.

## Características

- Parser de JSON com suporte a objetos e arrays.
- Serializador de objetos para JSON.
- Manuseio flexível de espaços em branco.
- Compatível com Visual Basic 6.0 e VBA.
- Sem dependências externas.
 
## Exemplo de uso
```vb
Copiar
Editar
Dim json As New hJsonBag
Dim dados As String
dados = "{""nome"":""João"",""idade"":30}"

Call json.parse(dados)

MsgBox json.Item("nome") ' Exibe: João
```
## Serializar objeto:

```vb
Copiar
Editar
Dim json As New hJsonBag
Call json.Add("nome", "Maria")
Call json.Add("idade", 25)

MsgBox json.stringify() ' Exibe: {"nome":"Maria","idade":25}

```

## Métodos principais

Método	Descrição
parse(jsonText As String)	Converte uma string JSON em um objeto manipulável.
stringify()	Gera uma string JSON a partir do objeto atual.
Add(key As String, value As Variant)	Adiciona um novo par chave-valor ao objeto.
Item(key As String)	Acessa o valor associado a uma chave.
getPath(path As String)	Acessa valores aninhados através de uma string de caminho.
toArray()	Converte o objeto JSON para um array VB6.
fromFile(filePath As String)	Lê JSON diretamente de um arquivo.

## Instalação
Baixe o arquivo hJsonBag.cls.

Adicione-o ao seu projeto VB6 (Project > Add Class Module > Existing).

Comece a usar!

## Licença
Este projeto é licenciado sob a Licença Apache 2.0.
Veja o arquivo LICENSE para mais detalhes.
