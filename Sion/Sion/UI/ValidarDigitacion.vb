Public Class ValidacionDigitacion
    Public Shared Sub validarAlfanumerico(ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
         Or (Asc(e.KeyChar) >= 65 And Asc(e.KeyChar) <= 90) _
        Or (Asc(e.KeyChar) >= 97 And Asc(e.KeyChar) <= 122) _
        Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 32) _
        Or (e.KeyChar = "ñ") Or (e.KeyChar = "Ñ") Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Shared Sub validarAlfabetico(ByVal e As KeyPressEventArgs)

        If (Asc(e.KeyChar) >= 65 And Asc(e.KeyChar) <= 90) _
  Or (Asc(e.KeyChar) >= 97 And Asc(e.KeyChar) <= 122) _
  Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 32) _
  Or (e.KeyChar = "ñ") Or (e.KeyChar = "Ñ") Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Shared Sub validarNumerosTelefono(ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
            Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 32) _
            Or (Asc(e.KeyChar) = 45) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Shared Sub validarValoresNumericos(ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
            Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 44) _
            Or (Asc(e.KeyChar) = 46) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Shared Sub validarSoloNumerosPositivo(ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
            Or (Asc(e.KeyChar) = 8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
    Public Shared Sub devolverKeyPress(ByVal sender As System.Object, ByVal e As KeyPressEventArgs)
        e.Handled = False
    End Sub
    Public Shared Sub validarSoloNumerosNegativo(ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
           Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 45) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Shared Sub validarAlfanumerico(ByVal sender As System.Object, ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
         Or (Asc(e.KeyChar) >= 65 And Asc(e.KeyChar) <= 90) _
        Or (Asc(e.KeyChar) >= 97 And Asc(e.KeyChar) <= 122) _
        Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 32) _
        Or (e.KeyChar = "ñ") Or (e.KeyChar = "Ñ") Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Shared Sub validarAlfabetico(ByVal sender As System.Object, ByVal e As KeyPressEventArgs)

        If (Asc(e.KeyChar) >= 65 And Asc(e.KeyChar) <= 90) _
  Or (Asc(e.KeyChar) >= 97 And Asc(e.KeyChar) <= 122) _
  Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 32) _
  Or (e.KeyChar = "ñ") Or (e.KeyChar = "Ñ") Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Shared Sub validarNumerosTelefono(ByVal sender As System.Object, ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
            Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 32) _
            Or (Asc(e.KeyChar) = 45) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Shared Sub validarValoresNumericos(ByVal sender As System.Object, ByVal e As KeyPressEventArgs)

        If ((Asc(e.KeyChar) = 46) Or (Asc(e.KeyChar) = 44)) Then
            e.KeyChar = Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator
        End If

        If ((Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
            Or (Asc(e.KeyChar) = 8) Or e.KeyChar = Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
    Public Shared Sub validarSoloNumerosNegativo(ByVal sender As System.Object, ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
           Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 45) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Shared Sub validarSoloNumerosPositivo(ByVal sender As System.Object, ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
            Or (Asc(e.KeyChar) = 8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Shared Sub validarSoloNumerosPositivoNo0(ByVal sender As System.Object, ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 49 And Asc(e.KeyChar) <= 57) _
            Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 48 And Not String.IsNullOrEmpty(sender.text)) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
    Public Shared Sub validarDinero(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        'si el caracter es Letra
        If Char.IsDigit(e.KeyChar) Then
            'acepta el cracter
            e.Handled = False
            'si es un caracter de control como Enter
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
            'si es un espacio en blanco
        ElseIf Char.IsPunctuation(e.KeyChar) Then
            e.KeyChar = ","
            If sender.text.ToString.Contains(",") Then
                e.Handled = True
            Else
                e.Handled = False
            End If
            'si es un espacio en blanco
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else
            ' de lo contario al poner e.handled en True
            'cancelamos la pulsación.
            e.Handled = True
        End If
        sender.text.ToString.Trim()
    End Sub

    Public Shared Function validarEmail(ByVal email As String) As Boolean
        Dim emailRegex As New System.Text.RegularExpressions.Regex("^(?<user>[^@]+)@(?<host>.+)$")
        Dim emailMatch As System.Text.RegularExpressions.Match = emailRegex.Match(email)

        Return emailMatch.Success
    End Function
End Class
