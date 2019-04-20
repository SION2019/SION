
Public Class CUPS_D
    Dim cups As New CupsDAL
    Public Function llenardgv(TextBox1 As String, codigo As String) As DataTable
        Dim dt As DataTable
        dt = cups.llenardgv(TextBox1, codigo)
        Return dt
    End Function

    Public Function guardar(ByRef dgv As DataGridView, ByVal codigoRef As String) As Boolean
        If cups.guardar(dgv, codigoRef) = True Then
            Return True
        Else
            Return False
        End If
    End Function


End Class
