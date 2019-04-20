Public Class CARTICULO_D

    Dim cups As New ArticuloSOATCapituloDAL
    Public Function llenar_dgv(TextBox1 As String) As DataTable
        Dim dt As DataTable
        dt = cups.llenar_dgv(TextBox1)
        Return dt
    End Function

    Public Function guardar(form_CARTICULO_SOAT As Form_CARTICULO_SOAT) As Boolean
        If cups.guardar(form_CARTICULO_SOAT) = True Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
