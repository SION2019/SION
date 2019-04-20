Public Class ValidarDigito
    Public Sub validar_alfanumerico(ByVal e As KeyPressEventArgs)
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

    Public Sub validar_alfabetico(ByVal e As KeyPressEventArgs)

        If (Asc(e.KeyChar) >= 65 And Asc(e.KeyChar) <= 90) _
  Or (Asc(e.KeyChar) >= 97 And Asc(e.KeyChar) <= 122) _
  Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 32) _
  Or (e.KeyChar = "ñ") Or (e.KeyChar = "Ñ") Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Sub validar_numeros_telefono(ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
            Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 32) _
            Or (Asc(e.KeyChar) = 45) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Sub validar_valores_numericos(ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
            Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 44) _
            Or (Asc(e.KeyChar) = 46) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Sub validar_solo_numeros(ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
            Or (Asc(e.KeyChar) = 8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Sub validar_solo_numeros_(ByVal e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) _
           Or (Asc(e.KeyChar) = 8) Or (Asc(e.KeyChar) = 45) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub


    Public Sub validar_dinero(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        'si el caracter es Letra
        If Char.IsDigit(e.KeyChar) Then
            'acepta el cracter
            e.Handled = False
            'si es un caracter de control como Enter
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
            'si es un espacio en blanco
        ElseIf Char.IsPunctuation(e.KeyChar) Then
            e.KeyChar = CChar(",")
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
End Class
