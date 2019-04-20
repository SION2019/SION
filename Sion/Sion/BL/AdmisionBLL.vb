Public Class AdmisionBLL
    Dim objAdmisionDAL As New AdmisionDAL
    Public Function guardarAdmision(ByVal objAdmision As Admision)
        If String.IsNullOrEmpty(objAdmision.nRegistro) Then
            If String.IsNullOrEmpty(objAdmision.solicitud) Then
                objAdmision.solicitud = Constantes.VALOR_PREDETERMINADO
            End If
            Return objAdmisionDAL.guardarAdmision(objAdmision)
        Else
            objAdmisionDAL.actualizarAdmision(objAdmision)
        End If
        Return objAdmision.nRegistro
    End Function
    Public Shared Function guardarTraslado(ByVal objAdmision As Admision, elementoAEliminar As List(Of String)) As String
        Dim objAdmisionDAL As New AdmisionDAL
        Return objAdmisionDAL.guardarTraslado(objAdmision, elementoAEliminar)
        Return objAdmision.nRegistro
    End Function
    Public Sub guardarDocumentos(objAdmisionDoc As Admision, consulta As String)
        Dim objAdmision As New AdmisionDAL
        objAdmision.guardarDocumento(objAdmisionDoc, consulta)
    End Sub

End Class
