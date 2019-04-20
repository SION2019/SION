
Public Class GRUPOQX_D
    Dim cups As New SoatGrupoQuirurgicoDAL
    Public Function llenar_dgv(TextBox1 As String) As DataTable
        Dim dt As DataTable
        dt = cups.llenar_dgv(TextBox1)
        Return dt
    End Function

    Public Function guardar() As Boolean
        If cups.guardar() = True Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
