
Public Class SECCION_D
    Dim cups As New SoatSeccionDAL
    Public Function llenar_dgv(TextBox1 As String) As DataTable
        Dim dt As DataTable
        dt = cups.llenar_dgv(TextBox1)
        Return dt
    End Function

    Public Function guardar(form_seccion_soat As Form_SECCION_SOAT) As Boolean
        If cups.guardar(form_seccion_soat) = True Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
