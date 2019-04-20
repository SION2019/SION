Public Class ConvertirNumeros
    Dim sValor As String, siValor As Single = Nothing
    Dim num_entero, num_decimal As Single
    Public NumeroConvertido As String

    Public Function Num2Text(ByVal value As Double) As String
        value = Int(value)
        Select Case value
            Case 0 : Num2Text = "CERO"
            Case 1 : Num2Text = "UN"
            Case 2 : Num2Text = "DOS"
            Case 3 : Num2Text = "TRES"
            Case 4 : Num2Text = "CUATRO"
            Case 5 : Num2Text = "CINCO"
            Case 6 : Num2Text = "SEIS"
            Case 7 : Num2Text = "SIETE"
            Case 8 : Num2Text = "OCHO"
            Case 9 : Num2Text = "NUEVE"
            Case 10 : Num2Text = "DIEZ"
            Case 11 : Num2Text = "ONCE"
            Case 12 : Num2Text = "DOCE"
            Case 13 : Num2Text = "TRECE"
            Case 14 : Num2Text = "CATORCE"
            Case 15 : Num2Text = "QUINCE"
            Case Is < 20 : Num2Text = "DIECI" & Num2Text(value - 10)
            Case 20 : Num2Text = "VEINTE"
            Case Is < 30 : Num2Text = "VEINTI" & Num2Text(value - 20)
            Case 30 : Num2Text = "TREINTA"
            Case 40 : Num2Text = "CUARENTA"
            Case 50 : Num2Text = "CINCUENTA"
            Case 60 : Num2Text = "SESENTA"
            Case 70 : Num2Text = "SETENTA"
            Case 80 : Num2Text = "OCHENTA"
            Case 90 : Num2Text = "NOVENTA"
            Case Is < 100 : Num2Text = Num2Text(Int(value \ 10) * 10) & " Y " & Num2Text(value Mod 10)
            Case 100 : Num2Text = "CIEN"
            Case Is < 200 : Num2Text = "CIENTO " & Num2Text(value - 100)
            Case 200, 300, 400, 600, 800 : Num2Text = Num2Text(Int(value \ 100)) & "CIENTOS"
            Case 500 : Num2Text = "QUINIENTOS"
            Case 700 : Num2Text = "SETECIENTOS"
            Case 900 : Num2Text = "NOVECIENTOS"
            Case Is < 1000 : Num2Text = Num2Text(Int(value \ 100) * 100) & " " & Num2Text(value Mod 100)
            Case 1000 : Num2Text = "MIL"
            Case Is < 2000 : Num2Text = "MIL " & Num2Text(value Mod 1000)
            Case Is < 1000000 : Num2Text = Num2Text(Int(value \ 1000)) & " MIL"
                If value Mod 1000 Then Num2Text = Num2Text & " " & Num2Text(value Mod 1000)
            Case 1000000 : Num2Text = "UN MILLON"
            Case Is < 2000000 : Num2Text = "UN MILLON" & Num2Text(value Mod 1000000)
            Case Is < 1000000000000.0# : Num2Text = Num2Text(Int(value / 1000000)) & " MILLONES"
                If (value - Int(value / 1000000) * 1000000) Then Num2Text = Num2Text & " " & Num2Text(value - Int(value / 1000000) * 1000000)
            Case 1000000000000.0# : Num2Text = "UN BILLON"
            Case Is < 2000000000000.0# : Num2Text = "UN BILLON" & Num2Text(value - Int(value / 1000000000000.0#) * 1000000000000.0#)
            Case Else : Num2Text = Num2Text(Int(value / 1000000000000.0#)) & " BILLONES"
                If (value - Int(value / 1000000000000.0#) * 1000000000000.0#) Then Num2Text = Num2Text & " " & Num2Text(value - Int(value / 1000000000000.0#) * 1000000000000.0#)
        End Select
    End Function

    Public Sub Convertir_Numero(ByVal numero As Double)
        sValor = Val(Replace(Format(numero, "General Number"), ",", ".").ToString)
        num_entero = Int(Val(sValor))
        num_decimal = CInt((sValor - num_entero) * 100)
        ' A ESPERAS DE DOLARIZACION
        num_decimal = 0
        If num_decimal > 0 Then
            NumeroConvertido = Num2Text(Val(sValor)) & " PESOS CON " + Num2Text(num_decimal) + " CENTAVOS MCTE"
        Else
            NumeroConvertido = Num2Text(Val(sValor)) & " PESOS MCTE"
        End If
        Return
    End Sub


#Region "FUNCIONES SOLO PARA CONTABILIDAD Y NOMINA"

    Public Function Num2MoneyTxt(ByVal value As Double) As String
        value = Int(value)
        Select Case value
            Case 0 : Num2MoneyTxt = "CERO"
            Case 1 : Num2MoneyTxt = "UN"
            Case 2 : Num2MoneyTxt = "DOS"
            Case 3 : Num2MoneyTxt = "TRES"
            Case 4 : Num2MoneyTxt = "CUATRO"
            Case 5 : Num2MoneyTxt = "CINCO"
            Case 6 : Num2MoneyTxt = "SEIS"
            Case 7 : Num2MoneyTxt = "SIETE"
            Case 8 : Num2MoneyTxt = "OCHO"
            Case 9 : Num2MoneyTxt = "NUEVE"
            Case 10 : Num2MoneyTxt = "DIEZ"
            Case 11 : Num2MoneyTxt = "ONCE"
            Case 12 : Num2MoneyTxt = "DOCE"
            Case 13 : Num2MoneyTxt = "TRECE"
            Case 14 : Num2MoneyTxt = "CATORCE"
            Case 15 : Num2MoneyTxt = "QUINCE"
            Case Is < 20 : Num2MoneyTxt = "DIECI" & Num2MoneyTxt(value - 10)
            Case 20 : Num2MoneyTxt = "VEINTE"
            Case Is < 30 : Num2MoneyTxt = "VEINTI" & Num2MoneyTxt(value - 20)
            Case 30 : Num2MoneyTxt = "TREINTA"
            Case 40 : Num2MoneyTxt = "CUARENTA"
            Case 50 : Num2MoneyTxt = "CINCUENTA"
            Case 60 : Num2MoneyTxt = "SESENTA"
            Case 70 : Num2MoneyTxt = "SETENTA"
            Case 80 : Num2MoneyTxt = "OCHENTA"
            Case 90 : Num2MoneyTxt = "NOVENTA"
            Case Is < 100 : Num2MoneyTxt = Num2MoneyTxt(Int(value \ 10) * 10) & " Y " & Num2MoneyTxt(value Mod 10)
            Case 100 : Num2MoneyTxt = "CIEN"
            Case Is < 200 : Num2MoneyTxt = "CIENTO " & Num2MoneyTxt(value - 100)
            Case 200, 300, 400, 600, 800 : Num2MoneyTxt = Num2MoneyTxt(Int(value \ 100)) & "CIENTOS"
            Case 500 : Num2MoneyTxt = "QUINIENTOS"
            Case 700 : Num2MoneyTxt = "SETECIENTOS"
            Case 900 : Num2MoneyTxt = "NOVECIENTOS"
            Case Is < 1000 : Num2MoneyTxt = Num2MoneyTxt(Int(value \ 100) * 100) & " " & Num2MoneyTxt(value Mod 100)
            Case 1000 : Num2MoneyTxt = "MIL"
            Case Is < 2000 : Num2MoneyTxt = "MIL " & Num2MoneyTxt(value Mod 1000)
            Case Is < 1000000 : Num2MoneyTxt = Num2MoneyTxt(Int(value \ 1000)) & " MIL"
                If value Mod 1000 Then Num2MoneyTxt = Num2MoneyTxt & " " & Num2MoneyTxt(value Mod 1000)
            Case 1000000 : Num2MoneyTxt = "UN MILLON DE"
            Case Is < 2000000 : Num2MoneyTxt = "UN MILLON DE" & Num2MoneyTxt(value Mod 1000000)
            Case Is < 1000000000000.0# : Num2MoneyTxt = Num2MoneyTxt(Int(value / 1000000)) & " MILLONES DE"
                If (value - Int(value / 1000000) * 1000000) Then Num2MoneyTxt = Num2MoneyTxt & " " & Num2MoneyTxt(value - Int(value / 1000000) * 1000000)
            Case 1000000000000.0# : Num2MoneyTxt = "UN BILLON DE"
            Case Is < 2000000000000.0# : Num2MoneyTxt = "UN BILLON DE" & Num2MoneyTxt(value - Int(value / 1000000000000.0#) * 1000000000000.0#)
            Case Else : Num2MoneyTxt = Num2MoneyTxt(Int(value / 1000000000000.0#)) & " BILLONES DE"
                If (value - Int(value / 1000000000000.0#) * 1000000000000.0#) Then Num2MoneyTxt = Num2MoneyTxt & " " & Num2MoneyTxt(value - Int(value / 1000000000000.0#) * 1000000000000.0#)
        End Select
    End Function

    Public Function funcionConvertir(ByVal numero As Double) As String
        sValor = Val(Replace(Format(numero, "General Number"), ",", ".").ToString)
        num_entero = Int(Val(sValor))
        num_decimal = CInt((sValor - num_entero) * 100)
        If num_decimal > 0 Then
            NumeroConvertido = Num2MoneyTxt(Val(sValor)) & " PESOS CON " + Num2MoneyTxt(num_decimal) + " CENTAVOS"
        Else
            NumeroConvertido = Num2MoneyTxt(Val(sValor)) & " PESOS"
        End If
        Return StrConv(NumeroConvertido, VbStrConv.ProperCase)
    End Function

#End Region

End Class
