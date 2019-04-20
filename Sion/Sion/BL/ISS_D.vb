
Public Class ISS_D
    Dim cups As New IssDAL
    Public Function guardar(ByVal dgvprocedimiento As DataGridView, ByVal codigoRef As String) As Boolean
        If cups.guardar(dgvprocedimiento, codigoRef) = True Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
