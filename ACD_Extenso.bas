Attribute VB_Name = "ExtensoACD"
'Macro:         ACD_Extenso
'Versão:        3.1 (Última atualização 20/08/2020)
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
'Gramática portuguesa:
'Regra Geral: Não se intercala a conjunção 'e' e nem vírgula entre posições de milhar.
'Exceção: Se a milhar posterior for menor que 100 ou for centena inteira (100,200,300...)
'Alguns gramáticos não aceitam essa exceção e outros já aceitam a vírgula.
'Nota: Segundo diversos gramáticos nunca deverá ser usada a vírgula na escrita de numerais por extenso.

Sub ACD_Extenso()

On Error GoTo Fim_Err

  Dim strValor As String
  Dim strRetorno As String
  Dim blnNoInicio As Boolean
  Dim strTmp As String
  Dim x As Integer
  
  Dim InicioExtenso As String
  Dim FimExtenso As String
  
  InicioExtenso = " (-"
  FimExtenso = "-) "
  
  If Selection.Type = wdSelectionIP And Selection.Start = 0 Then blnNoInicio = True
  If blnNoInicio = True Then Exit Sub
  
  Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
  
  While Selection.Text = " "
  
    If WordBasic.AtStartOfDocument() Then Exit Sub
      Selection.ExtendMode = False
      Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
  Wend
  
  Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
  Selection.ExtendMode = True
  
  With Selection.Find
      .Forward = False
      .Wrap = wdFindStop
      .Execute FindText:=" "
  End With
  
  Debug.Print Selection.Text
  strValor = Extenso(CCur(Selection.Text))
  
  Selection.ExtendMode = False
  
  Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
  'caso queira tudo em minúscula, tire o UCASE abaixo 
  strTmp = InicioExtenso & UCASE(strValor) & FimExtenso
  
  x = Len(strTmp)
  If x > 0 And strTmp <> " " Then
     Selection.TypeText Text:=strTmp
   End If

Fim_Err:

Selection.ExtendMode = False
Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
   Exit Sub

End Sub

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
  
    If ParteInteira = 0 Then
    
        If ParteDecimal > 0 Then
            If ParteDecimal = 1 Then
              s = s & "um " & CentavosNoSingular & " de " & MoedaNoSingular
            Else
              s = s & Centena(ParteDecimal) & " " & CentavosNoPlural & " de " & MoedaNoSingular
            End If
        End If
        Extenso = s
        Exit Function
    
    End If
  
    If ParteInteira > 0 Then
      
      s = ConcatCentenas(ParteInteira)
      If s = "um" Then
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
    Dim varUnidade As Variant
    varUnidade = Array("", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove")
    Unidade = varUnidade(N)
End Function

Private Function Dezena(N As Long) As String
    Dim d As Long, u As Long
    Dim s As String
    Dim varDezena1 As Variant, varDezena2 As Variant
    
    varDezena1 = Array("dez", "onze", "doze", "treze", "quatorze", _
                      "quinze", "dezesseis", "dezessete", "dezoito", _
                      "dezenove")
    
    varDezena2 = Array("vinte", "trinta", "quarenta", "cinquenta", _
                       "sessenta", "setenta", "oitenta", "noventa")
                      
    d = N \ 10   '\ divide 2 numeros e retorna o resultado inteiro
    u = N Mod 10 'mod retorna o resto da divisão

If d = 0 Then
    Dezena = Unidade(N)
    Exit Function
Else
    If d = 1 And u = 0 Then
        Dezena = varDezena1(0)
        Exit Function
    
    End If
    If d = 1 And u > 0 Then
        Dezena = varDezena1(u)
        Exit Function
    End If
End If

If d > 1 Then '-> mudei aqui para arrumar as centenas
    If u = 0 Then
        Dezena = varDezena2(d - 2) & Unidade(u)
        Else
        Dezena = varDezena2(d - 2) & " e " & Unidade(u)
    End If
End If

End Function

Private Function Centena(N As Long) As String
  Dim c As Long, d As Long
  Dim s As String
  Dim varCentena As Variant
  
  varCentena = Array("duzentos", "trezentos", "quatrocentos", "quinhentos", _
                     "seiscentos", "setecentos", "oitocentos", "novecentos")
                     
  c = N \ 100
  d = N Mod 100
  
  If c = 0 Then
    Centena = Dezena(N)
    Exit Function
  End If
  
  If c = 1 Then
    If d = 0 Then
        Centena = "cem"
    Else
        Centena = "cento e " & Dezena(d)
    End If
    Exit Function
  End If
  
  If d = 0 Then
    Centena = varCentena(c - 2)
  Else
    Centena = varCentena(c - 2) & " e " & Dezena(d)
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
  Dim TiraUmMil As Boolean
  Dim s As String, m As Currency
  
  TiraUmMil = True 'Mude aqui para tirar o Um Mil (True=Mil False=Um Mil)
  
  s = ""
  m = N
  
  Um = Resto(N, 1000)      'Um = N Mod 1000
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
        's = s & ", " 'retire aqui virguls
        s = s & " "
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
        's = s & ", " 'retire aqui virgula
        s = s & " "
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
        '       s = s & ", " 'retire aqui sem vírgula no milhar
       s = s & " "
      End If
    Else
      s = s & " de"
    End If
  End If
  
  m = -(Milhar * 1000) + m
  Menores = Um
  If Milhar > 0 Then
  
    s = s & Centena(Milhar) & " mil "
    'AQUI TIRA O UM MIL
    If TiraUmMil Then
        If Left$(s, 7) = "um mil " Then s = Mid$(s, 4)
    End If


    If Menores > 0 Then
      If SingleAlg(m) Then
        s = s & "e "
      Else
       ' s = s & ", " '->Retirei aqui para não sair vígula no milhar
      End If
    End If
  End If
  s = s & Centena(Um)
  ConcatCentenas = s
End Function
'fim da função ACD_Extenso
