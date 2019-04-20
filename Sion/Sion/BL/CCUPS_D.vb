
Public Class CCUPS_D
    Dim cups As New CupsCapituloDAL
    Public Function guardar(ByRef dgvprocedimiento As DataGridView) As Boolean
        If cups.guardar(dgvprocedimiento) = True Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
