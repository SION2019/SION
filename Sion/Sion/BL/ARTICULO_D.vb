
Public Class ARTICULO_D
    Dim cups As New ArticuloDAL
    Public Function llenar_dgv(TextBox1 As String) As DataTable
        Dim dt As DataTable
        dt = cups.llenar_dgv(TextBox1)
        Return dt
    End Function

    Public Function guardar(ByRef articulosoat As ArticuloSoat) As Boolean
        If cups.guardar(articulosoat) = True Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
