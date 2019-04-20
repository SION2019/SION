

Public Class SoatBLL
    Dim cups As New SoatDAL
    Public Function guardar(ByRef dgvprocedimiento As DataGridView, codigoManual As Integer) As Boolean
        If cups.guardar(dgvprocedimiento, codigoManual) = True Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
