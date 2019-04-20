
Public Class CATCUPS_D
    Dim cups As New CupsCategoriaDAL
    Public Function guardar(dgvprocediminto As DataGridView) As Boolean
        If cups.guardar(dgvprocediminto) = True Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
