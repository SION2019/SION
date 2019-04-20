
Public Class SALA_D
    Dim cups As New SoatSalaProcedimientosDAL
    Public Function llenar_dgv(TextBox1 As String) As DataTable
        Dim dt As DataTable
        dt = cups.llenar_dgv(TextBox1)
        Return dt
    End Function

    Public Function guardar(FORM_SALA_SOAT As Form_SALA_SOAT) As Boolean
        If cups.guardar(FORM_SALA_SOAT) = True Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
