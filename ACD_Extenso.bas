Attribute VB_Name = "ExtensoACD"
'Macro:         ACD_Extenso
'Versão:        2.4 (Última atualização 17/08/2020)
'Finalidade:    Converte um valor numérico em uma string
'               com o extenso monetário correspondente.
'Linguagem:     VBA
'Autor:         Antonio Carlos Doná
'               acdona@hotmail.com
'Distribuição:  Livre e sem garantias, use por sua conta e risco.
'Observações:   1) Sempre deixar um espaço em branco no início
'                  do número a ser convertido para extenso.
'               2) Favor reportar eventuais bugs.
'               3) Suporta valores até $922.337.203.685.477,5807
'               4) Não foi feito nenhum teste com valores negativos
'
'Início da rotina ACD_Extenso()
Sub ACD_Extenso()
'Verifica se houve erro e pulo para Fim:
On Error GoTo Fim_Err
'declara as variáveis
  Dim strValor As String        'alfanumérico
  Dim strRetorno As String      'alfanumério
  Dim blnNoInicio As Boolean    'Falso/verdadeiro
  Dim strTmp As String          'alfanumérico
  Dim x As Integer              'inteiro
      
  'Verifica que não existe algo selecionado
  'atribui Verdadeiro para blnInicio e sai da macro
  If Selection.Type = wdSelectionIP And Selection.Start = 0 Then blnNoInicio = True
  If blnNoInicio = True Then Exit Sub
  
  'Move para esquerda
  'unit = por caracter
  'count = um por vez
  'Extend = move para o final do número extendido
  Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
  'quando achar um espaço em branco
  While Selection.Text = " "
    'se for início do documento e não achou espaço sai fora
    If WordBasic.AtStartOfDocument() Then Exit Sub
    'volta onde estava sem marcar nada
    Selection.ExtendMode = False
     'Move para esquerda
     'unit = por caracter
     'count = um por vez
     'Extend = move para final do número e pula para esquerda
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
  Wend
  'volta para direita selecionando todo o número
  Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
  Selection.ExtendMode = True
  'procura na seleção
  With Selection.Find
      'em direção ao início
      .Forward = False
      'quando encontrar o fim, para.
      .Wrap = wdFindStop
      'procura o espaço em branco
      .Execute FindText:=" "
  End With
  'exibe texto na janela imediata
  Debug.Print Selection.Text
  'atribui texto selecionado à macro
  'o CCur é para transformar de texto para monetário
  strValor = Extenso(CCur(Selection.Text))
  'desmarca seleção
  Selection.ExtendMode = False
  'volta para direita um caracter
  Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
  'esta linha abaixo, retorna os parenteses e hífens
  'caso não queira nada apenas o extenso,
  'mude-a para  strTmp = strValor
   
  'Deixa apenas a primeira letra maiúscula
  strTmp = " (-" & UCase(Left(strValor, 1)) & Trim$(Right(strValor, Len(strValor) - 1)) & "-)"
  
 ' strTmp = " (- " & Trim$(strValor) & " -)"
  x = Len(strTmp)
  If x > 0 And strTmp <> " " Then
     ' transforma todo o extenso em maiúsculas
     ' Para todas minúsculas use LCase(strTmp)
     ' Para apenas as primeiras maiúsculas use essa abaixo
     ' Selection.TypeText Text:=StrConv(strTmp, vbProperCase)
     ' e inclui o extenso após o número
     ' wdTitleSentence
     ' Selection.TypeText Text:=UCase(strTmp)
     Selection.TypeText Text:=strTmp
    ' selection.TypeText
     
     
   End If
'Rotina para sair da macro
Fim_Err: ' para tratar erros

Selection.ExtendMode = False
Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
   Exit Sub

End Sub
'Fim da rotina extenso ACD_Extenso()

'Função:        Extenso
'Attribute VB_Name = "Extenso"
Function Extenso( _
  Valor As Currency, _
  Optional MoedaNoSingular As String = "real", _
  Optional MoedaNoPlural As String = "reais", _
  Optional CentavosNoSingular As String = "centavo", _
  Optional CentavosNoPlural As String = "centavos") _
As String
  Dim ParteInteira As Currency, ParteDecimal As Long
  Dim s As String
  
  ParteInteira = Fix(Valor)
  ParteDecimal = Fix((Valor - ParteInteira) * 100)
  
  s = ""
  If ParteInteira > 0 Then
    s = ConcatCentenas(ParteInteira)
    If s = "um" Then 'ParteInteira = 1 ou 1.0 ou 1# não funciona
      s = s & " " & MoedaNoSingular
    Else
      s = s & " " & MoedaNoPlural
    End If
    If ParteDecimal > 0 Then
      s = s & " e "
    End If
  End If
  
  If ParteDecimal > 0 Then
    If ParteDecimal = 1 Then
      s = s & "um " & CentavosNoSingular
    Else
      s = s & Centena(ParteDecimal) & " " & CentavosNoPlural
    End If
  End If
  Extenso = s
End Function

Function Resto(A As Currency, B As Long) As Currency
  Dim Aux As String, r As Currency
  Aux = Format(A / B, "###0.0000")
  Aux = Right$(Aux, 4)
  Resto = Aux * B / 10000
  If Resto < 1 And Resto > 0.99 Then
    Resto = 1
  End If
  Aux = Format(Resto, "###0.0000")
  Aux = Right$(Aux, 4)
  Resto = Resto - Aux / 10000
End Function

Function DivInt(A As Currency, B As Long) As Currency
  Dim Aux As String
  DivInt = A / B
  Aux = Format(DivInt, "###0.0000")
  Aux = Right$(Aux, 4)
  DivInt = DivInt - Aux / 10000
End Function

Private Function Unidade(N As Long) As String
  Select Case N
  Case 0
    Unidade = ""
  Case 1
    Unidade = "um"
  Case 2
    Unidade = "dois"
  Case 3
    Unidade = "três"
  Case 4
    Unidade = "quatro"
  Case 5
    Unidade = "cinco"
  Case 6
    Unidade = "seis"
  Case 7
    Unidade = "sete"
  Case 8
    Unidade = "oito"
  Case 9
    Unidade = "nove"
  Case Else
    Err.Raise vbObjectError + 513, , "O número deve estar entre 0 e 9"
  End Select
End Function

Private Function Dezena(N As Long) As String
  Dim d As Long, u As Long
  Dim s As String
  
  d = N \ 10
  u = N Mod 10
  
  Select Case d
  Case 0
    Dezena = Unidade(N)
    Exit Function
  Case 1
    Select Case u
    Case 0
      Dezena = "dez"
    Case 1
      Dezena = "onze"
    Case 2
      Dezena = "doze"
    Case 3
      Dezena = "treze"
    Case 4
      Dezena = "quatorze"
    Case 5
      Dezena = "quinze"
    Case 6
      Dezena = "dezesseis"
    Case 7
      Dezena = "dezessete"
    Case 8
      Dezena = "dezoito"
    Case 9
      Dezena = "dezenove"
    End Select
    Exit Function
  Case 2
    s = "vinte"
  Case 3
    s = "trinta"
  Case 4
    s = "quarenta"
  Case 5
    s = "cinquenta"
  Case 6
    s = "sessenta"
  Case 7
    s = "setenta"
  Case 8
    s = "oitenta"
  Case 9
    s = "noventa"
  Case Else
    Err.Raise vbObjectError + 513, , "O número deve estar entre 0 e 99"
  End Select
    
  If u = 0 Then
    Dezena = s
  Else
    Dezena = s & " e " & Unidade(u)
  End If
End Function

Private Function Centena(N As Long) As String
  Dim c As Long, d As Long
  Dim s As String
  c = N \ 100
  d = N Mod 100
  
  Select Case c
  Case 0
    Centena = Dezena(N)
    Exit Function
  Case 1
    If d = 0 Then
      Centena = "cem"
    Else
      Centena = "cento e " & Dezena(d)
    End If
    Exit Function
  Case 2
    s = "duzentos"
  Case 3
    s = "trezentos"
  Case 4
    s = "quatrocentos"
  Case 5
    s = "quinhentos"
  Case 6
    s = "seiscentos"
  Case 7
    s = "setecentos"
  Case 8
    s = "oitocentos"
  Case 9
    s = "novecentos"
  Case Else
    Err.Raise vbObjectError + 513, , "O número deve estar entre 0 e 999"
  End Select
  
  If d = 0 Then
    Centena = s
  Else
    Centena = s & " e " & Dezena(d)
  End If
End Function

Private Function SingleAlg(N As Currency) As Boolean
  Dim s As String, i As Integer
  s = N
  SingleAlg = False
  For i = 1 To Len(s)
    If Mid$(s, i, 1) <> 0 Then
      If SingleAlg Then
        SingleAlg = False
        Exit Function
      Else
        SingleAlg = True
      End If
    End If
  Next i
End Function

Private Function ConcatCentenas(N As Currency) As String
  Dim Trilhao As Long, Bilhao As Long, _
    Milhao As Long, Milhar As Long, Um As Long, _
    Menores As Integer
  Dim s As String, m As Currency
  
  s = ""
  m = N
  
  Um = Resto(N, 1000)  'Um = N Mod 1000
  N = DivInt(N, 1000)      'N = N \ 1000
  Milhar = Resto(N, 1000)  'Milhar = N Mod 1000
  N = DivInt(N, 1000)      'N = N \ 1000
  Milhao = Resto(N, 1000)  'Milhao = N Mod 1000
  N = DivInt(N, 1000)      'N = N \ 1000
  Bilhao = Resto(N, 1000)  'Bilhao = N Mod 1000
  N = DivInt(N, 1000)      'N = N \ 1000
  Trilhao = Resto(N, 1000) 'Trilhao = N Mod 1000000000

  m = m - Trilhao * 1000000000000@
  Menores = Bilhao + Milhao + Milhar + Um
  If Trilhao > 0 Then
    If Trilhao = 1 Then
      s = "um trilhão"
    Else
      s = Centena(Trilhao) & " trilhões"
    End If
    If Menores > 0 Then
      If SingleAlg(m) Then
        s = s & " e "
      Else
        s = s & ", "
      End If
    Else
      s = s & " de"
    End If
  End If
  
  m = m - Bilhao * 1000000000@
  Menores = Milhao + Milhar + Um
  If Bilhao > 0 Then
    If Bilhao = 1 Then
      s = s & "um bilhão"
    Else
      s = s & Centena(Bilhao) & " bilhões"
    End If
    If Menores > 0 Then
      If SingleAlg(m) Then
        s = s & " e "
      Else
        s = s & ", "
      End If
    Else
      s = s & " de"
    End If
  End If
  
  m = m - Milhao * 1000000
  Menores = Milhar + Um
  If Milhao > 0 Then
    If Milhao = 1 Then
      s = s & "um milhão"
    Else
        s = s & Centena(Milhao) & " milhões"
    End If
    If Menores > 0 Then
      If SingleAlg(m) Then
        s = s & " e "
      Else
        s = s & ", "
      End If
    Else
      s = s & " de"
    End If
  End If
  
  m = -(Milhar * 1000) + m
  Menores = Um
  If Milhar > 0 Then
    s = s & Centena(Milhar) & " mil "
    If Menores > 0 Then
      If SingleAlg(m) Then
        s = s & " e "
      Else
      '
      
      's = s & ", "
      End If
    End If
  End If
  
  s = s & Centena(Um)
  ConcatCentenas = s
End Function
'fim da função

'---------------------------------------------------------------------
'início da macro WordExtenso()
Public Sub WordExtenso()
'/--------------------------------------------------------------------\
' WordExtenso Macro                                                   '
' Macro criada 24/05/2006 por Antonio Carlos Doná                     '
' essa macro é para ser usada em formulários                          '
' Obrigatoriamente os campos devem ser:                               '
' Valor em currency                                                   '
' ValorPorExtenso em string                                           '
' Exemplo: você coloca um campo com o nome de valor                   '
' e outro com o nome de ValorPorExtenso e em                          '
' executar macro na Entreda: Wordextenso                              '
'\--------------------------------------------------------------------/
    'Verifica se houve erro e pulo para Fim:
    On Error GoTo Fim
    'Atribui tipo Moeda para variável xValor
    Dim xValor As Currency
    'Atribui a variável xValor o indicador Valor do documento
    xValor = ActiveDocument.FormFields("Valor").Result
    'Preenche o indicador ValorPorExtenso do documento com extenso e maiúscula
    ActiveDocument.FormFields("ValorPorExtenso").Result = UCase(Módulo1.Extenso(xValor))
  
'Rotina para sair da macro
WordExtenso_Fim:
    Exit Sub

' a rotina caso de erro
' mostra o número e descrição do erro
' e vai para outra rotina WordExtenso_Fim
Fim:
MsgBox " ERRO " & Err.Number & " - " & Err.Description
    Resume WordExtenso_Fim
End Sub
'fim da macro WordExtenso()
'---------------------------------------------------------------------

