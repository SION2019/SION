Public Class ResolucionBLL
    Dim objResolucionDAL As New ResolucionDAL
    Public Function guardarResolucion(ByVal objResolucion As Resolucion)
        If String.IsNullOrEmpty(objResolucion.codigo) Then
            Return objResolucionDAL.guardarResolucion(objResolucion)
        Else
            objResolucionDAL.actualizarResolucion(objResolucion)
        End If
        Return objResolucion.codigo
    End Function
End Class
