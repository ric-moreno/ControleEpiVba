Attribute VB_Name = "valor_extenso"
Function EA_Valor_Escrito(valor As Double) As String
'exceleaccess.com
    Dim strMoeda    As String
    Dim cents       As Variant
    Dim decimalSep  As String
    If valor > 999999999999999# Then
        Extenso_Valor = "Valor excede 999.999.999.999.999"
        Exit Function
    End If
    If WorksheetFunction.RoundDown(valor, 0) = 1 Then
        strMoeda = " real"
        ElseIf WorksheetFunction.RoundDown(valor, 0) > 1 Then
            strMoeda = " reais"
    End If
    cents = valor - WorksheetFunction.RoundDown(valor, 0)
    valor = valor - CDbl(cents)
        cents = centavos(CDbl(cents) * 100)
    If cents <> "" And valor >= 1 Then
        cents = " e " & cents
    End If
    strMoeda = Trim(Trilhoes(valor)) & strMoeda & cents
    strMoeda = Replace(strMoeda, ", e", " e")
    strMoeda = Replace(strMoeda, ", r", " r")
    If Left(strMoeda, 2) = "e " Then
        strMoeda = Mid(strMoeda, 3, Len(strMoeda))
    End If
    vzz = "00000000000000000000"
    vtam = Len(Trim(Mid(Trim(valor), 2, 100)))
    If Right(vzz + vzz + vzz + vzz, vtam) = Mid(Trim(valor), 2, 100) And InStr(UCase(strMoeda), UCase("es ")) > 0 Then
        vetor = Split(strMoeda, " ")
        vtrocar = vetor(UBound(vetor))
        strMoeda = Replace(strMoeda, vtrocar, "de " + vtrocar)
    End If
    EA_Valor_Escrito = strMoeda

End Function

Private Function centavos(valor As Double) As String
'exceleaccess.com
    Dim dezena      As Integer
    Dim unidade     As Integer
    
    valor = Round(CDbl(valor / 100), 2)
    If valor = 0.01 Then
        centavos = "um centavo"
        Exit Function
    End If
    valor = valor * 100
    If dezenas(valor) = "" Then
        centavos = ""
    Else
        centavos = dezenas(valor) & " centavos"
    End If

End Function

Private Function unidades(unidade As Double) As String
'exceleaccess.com
    Dim unid(9)
    unid(1) = "um": unid(2) = "dois": unid(3) = "três": unid(4) = "quatro": unid(5) = "cinco"
    unid(6) = "seis": unid(7) = "sete": unid(8) = "oito": unid(9) = "nove"
    unidades = Trim(unid(unidade))
End Function

Private Function dezenas(dezena As Double) As String
'exceleaccess.com
    Dim dezes(9)
    Dim dez(9)
    Dim intDezena       As Double
    Dim intUnidade      As Double
    Dim tmpStr          As String

    dezes(1) = "onze": dezes(2) = "doze": dezes(3) = "treze": dezes(4) = "quatorze": dezes(5) = "quinze"
    dezes(6) = "dezesseis": dezes(7) = "dezessete": dezes(8) = "dezoito": dezes(9) = "dezenove"
    
    dez(1) = "dez": dez(2) = "vinte": dez(3) = "trinta": dez(4) = "quarenta": dez(5) = "cinquenta"
    dez(6) = "sessenta": dez(7) = "setenta": dez(8) = "oitenta": dez(9) = "noventa"
    
    intDezena = Int(dezena / 10)
    intUnidade = dezena Mod 10
    If intDezena = 0 Then
        dezenas = unidades(intUnidade)
        Exit Function
    Else:
        dezenas = dez(intDezena)
    End If

    If (intDezena = 1 And intUnidade > 0) Then
        dezenas = dezes(intUnidade)
    Else
        If (intDezena > 1 And intUnidade > 0) Then
            dezenas = dezenas & " e " & unidades(intUnidade)
        End If
    End If
    dezenas = dezenas
End Function

Private Function centenas(centena As Double) As String
'exceleaccess.com
    Dim tmpCento      As Double
    Dim tmpDez        As Double
    Dim tmpUni        As Double
    Dim tmpUniMod     As Double
    Dim tmpModDez     As Double
    Dim centoString   As String
    Dim cento(9)

    cento(1) = "cento": cento(2) = "duzentos": cento(3) = "trezentos": cento(4) = "quatrocentos": cento(5) = "quinhentos"
    cento(6) = "seiscentos": cento(7) = "setecentos": cento(8) = "oitocentos": cento(9) = "novecentos"
        
    tmpCento = Int(centena / 100)
    tmpDez = centena - (tmpCento * 100)
    tmpUni = Int(tmpDez / 10)
    tmpUniMod = tmpUni Mod 10
    tmpModDez = tmpDez Mod 10
    If centena = 100 Then
        centoString = "cem "
    Else
        centoString = cento(tmpCento)
    End If
    If (tmpUni >= 0 And tmpUniMod >= 0 And tmpDez >= 1 And tmpCento >= 1) Then
        centoString = centoString & " e "
    End If
    centenas = Trim(centoString & dezenas(tmpDez))
End Function

Private Function milhares(milhar As Double) As String
'exceleaccess.com
    Dim tmpMilhar      As Double
    Dim tmpCento       As Double
    Dim milString      As String
    
    tmpMilhar = Int(milhar / 1000)
    tmpCento = milhar - (tmpMilhar * 1000)
    If tmpMilhar = 0 Then milString = ""
    If (tmpMilhar >= 1 And tmpMilhar < 10) Then
        milString = unidades(tmpMilhar) & " mil, "
        ElseIf (tmpMilhar >= 10 And tmpMilhar < 100) Then
            milString = dezenas(tmpMilhar) & " mil, "
            ElseIf (tmpMilhar >= 100 And tmpMilhar < 1000) Then
                milString = centenas(tmpMilhar) & " mil, "
    End If
    If (tmpCento >= 1 And tmpCento <= 100) Then milString = milString & "e "
    milhares = Trim(milString & centenas(tmpCento))
End Function

Private Function milhoes(milhao As Double) As String
'exceleaccess.com
    Dim tmpMilhao      As Double
    Dim tmpMilhares    As Double
    Dim miString       As String
    
    tmpMilhao = Int(milhao / 1000000)
    tmpMilhares = milhao - (tmpMilhao * 1000000)
    If tmpMilhao = 0 Then miString = ""
    If (tmpMilhao = 1) Then
        miString = unidades(tmpMilhao) & " milhão, "
        ElseIf (tmpMilhao > 1 And tmpMilhao < 10) Then
            miString = unidades(tmpMilhao) & " milhões, "
            ElseIf (tmpMilhao >= 10 And tmpMilhao < 100) Then
                miString = dezenas(tmpMilhao) & " milhões, "
                ElseIf (tmpMilhao >= 100 And tmpMilhao < 1000) Then
                    miString = centenas(tmpMilhao) & " milhões, "
    End If
    If milhao = 1000000# Then miString = "um milhão de "
    milhoes = Trim(miString & milhares(tmpMilhares))
End Function

Private Function bilhoes(bilhao As Double) As String
'exceleaccess.com
    Dim tmpBilhao     As Double
    Dim tmpMilhao       As Double
    Dim biString       As String
    
    tmpBilhao = Int(bilhao / 1000000000)
    tmpMilhao = bilhao - (tmpBilhao * 1000000000)
    If (tmpBilhao = 1) Then
        biString = unidades(tmpBilhao) & " bilhão, "
        ElseIf (tmpBilhao > 1 And tmpBilhao < 10) Then
            biString = unidades(tmpBilhao) & " bilhões, "
            ElseIf (tmpBilhao >= 10 And tmpBilhao < 100) Then
                biString = dezenas(tmpBilhao) & " bilhões, "
                ElseIf (tmpBilhao >= 100 And tmpBilhao < 1000) Then
                    biString = centenas(tmpBilhao) & " bilhões, "
    End If
    If bilhao = 1000000000# Then biString = "um bilhão de "
    bilhoes = Trim(biString & milhoes(tmpMilhao))
End Function

Private Function Trilhoes(Trilhao As Double) As String
'exceleaccess.com
    Dim tmpTrilhao     As Double
    Dim tmpBilhao       As Double
    Dim triString       As String
    
    tmpTrilhao = Int(Trilhao / 1000000000000#)
    tmpBilhao = Trilhao - (tmpTrilhao * 1000000000000#)
    If (tmpTrilhao = 1) Then
        triString = unidades(tmpTrilhao) & " trilhão, "
        ElseIf (tmpTrilhao > 1 And tmpTrilhao < 10) Then
            triString = unidades(tmpTrilhao) & " trilhões, "
            ElseIf (tmpTrilhao >= 10 And tmpTrilhao < 100) Then
                triString = dezenas(tmpTrilhao) & " trilhões, "
                ElseIf (tmpTrilhao >= 100 And tmpTrilhao < 1000) Then
                    triString = centenas(tmpTrilhao) & " trilhões, "
    End If
    If Trilhao = 1000000000000# Then triString = "um trilhão de "
    Trilhoes = Trim(triString & bilhoes(tmpBilhao))
End Function

Function arredBaixo(valor)
'exceleaccess.com
    Dim tmpValor
    tmpValor = Round(CDbl(Right(Round(valor, 2) * 100, 2)) / 100, 2)
    arredBaixo = Round(Round(valor, 2) - tmpValor, 0)
End Function



