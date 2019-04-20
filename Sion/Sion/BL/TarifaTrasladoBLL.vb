Public Class Traslado_Ambulancia_D
    Dim objTrasladoDAL As New Traslado_Ambulancia_C
    Public Function guardarTraslado(ByVal objTrasladoAmbulancia As TrasladoAmbulancia)
        If String.IsNullOrEmpty(objTrasladoAmbulancia.codigo) Then
            Return objTrasladoDAL.guardarTraslado(objTrasladoAmbulancia)
        Else
            objTrasladoDAL.actualizarTraslado(objTrasladoAmbulancia)
        End If
        Return objTrasladoAmbulancia.codigo
    End Function

End Class
