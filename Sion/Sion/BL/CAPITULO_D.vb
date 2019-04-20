
Public Class CAPITULO_D

    Dim cups As New CapituloDAL
    Public Function llenar_dgv(TextBox1 As String) As DataTable
        Dim dt As DataTable
        dt = cups.llenar_dgv(TextBox1)
        Return dt
    End Function

    Public Function guardar(Form_CAPITULO_SOAT As Form_CAPITULO_SOAT) As Boolean
        If cups.guardar(Form_CAPITULO_SOAT) = True Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
