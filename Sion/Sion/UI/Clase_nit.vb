Public Class DigitoVerificacion
    Private Property mDocumento As String = vbNullString
    Private Property mDigito As Integer
    Public Function CalculaNit(Numero As String) As Integer
        Numero = LimpiarSeparadores(Numero)
        mDocumento = Format(Val(Numero), "000000000000000")
        CalcularDigito()
        Return mDigito
    End Function
    Public Function VerificaNit(Numero As String, Digito As Integer) As Boolean
        Numero = LimpiarSeparadores(Numero)
        mDocumento = Format(Val(Numero), "000000000000000")
        CalcularDigito()
        Return mDigito = Digito
    End Function
    Public Function VerificaNit(Numero As String) As Boolean
        Dim Digito As Integer
        Numero = LimpiarSeparadores(Numero)
        mDocumento = Format(Val(Numero), "0000000000000000")
        Digito = Val(Mid(mDocumento, 16, 1))
        mDocumento = Mid(mDocumento, 1, 15)
        CalcularDigito()
        Return mDigito = Digito
    End Function
    Private Sub CalcularDigito()
        Dim Total As Long
        Dim Residuo As Integer
        Total = Val(Mid(mDocumento, 1, 1)) * 71
        Total += Val(Mid(mDocumento, 2, 1)) * 67
        Total += Val(Mid(mDocumento, 3, 1)) * 59
        Total += Val(Mid(mDocumento, 4, 1)) * 53
        Total += Val(Mid(mDocumento, 5, 1)) * 47
        Total += Val(Mid(mDocumento, 6, 1)) * 43
        Total += Val(Mid(mDocumento, 7, 1)) * 41
        Total += Val(Mid(mDocumento, 8, 1)) * 37
        Total += Val(Mid(mDocumento, 9, 1)) * 29
        Total += Val(Mid(mDocumento, 10, 1)) * 23
        Total += Val(Mid(mDocumento, 11, 1)) * 19
        Total += Val(Mid(mDocumento, 12, 1)) * 17
        Total += Val(Mid(mDocumento, 13, 1)) * 13
        Total += Val(Mid(mDocumento, 14, 1)) * 7
        Total += Val(Mid(mDocumento, 15, 1)) * 3
        Residuo = Total Mod 11
        Select Case Residuo
            Case 0 : mDigito = 0
            Case 1 : mDigito = 1
            Case Else : mDigito = 11 - Residuo
        End Select
    End Sub
    Private Function LimpiarSeparadores(Numero As String) As String
        Numero = Replace(Numero, ",", "")
        Numero = Replace(Numero, ".", "")
        Numero = Replace(Numero, "-", "")
        Numero = Replace(Numero, " ", "")
        Return Numero
    End Function
End Class

