
Public Class GRUPO_D
    Dim cups As New SoatGrupoDAL
    Public Function llenar_dgv(TextBox1 As String) As DataTable
        Dim dt As DataTable
        dt = cups.llenar_dgv(TextBox1)
        Return dt
    End Function

    Public Function guardar(form As Form_GRUPO_SOAT) As Boolean
        If cups.guardar(form) = True Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
