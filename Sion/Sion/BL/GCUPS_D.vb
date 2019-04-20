
Public Class GCUPS_D

    Dim cups As New CupsGrupoDAL
    Public Function llenar_dgv(TextBox1 As String) As DataTable
        Return cups.llenar_dgv(TextBox1)
    End Function

    Public Function guardar(ByRef dgvprocedimiento As DataGridView, usuario As Integer) As Boolean
        Return cups.guardar(dgvprocedimiento, usuario)
    End Function
End Class
