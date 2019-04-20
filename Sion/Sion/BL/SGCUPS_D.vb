
Public Class SGCUPS_D
    Dim cups As New CupsSubgrupoDAL
    Public Function llenar_dgv(TextBox1 As String) As DataTable
        Dim dt As DataTable
        dt = cups.llenar_dgv(TextBox1)
        Return dt
    End Function

    Public Function guardar(ByRef sgcups As SGCUPS) As Boolean
        If cups.guardar(sgcups) = True Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
