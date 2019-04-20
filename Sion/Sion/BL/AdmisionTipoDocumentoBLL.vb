Public Class AdmisionTipoDocumentoBLL
    Dim objDocumentoDAL As New DocumentoDAL
    Public Function guardarDocumento(ByVal objDocumento As TipoDocumentos)
        If String.IsNullOrEmpty(objDocumento.codigo) Then
            Return objDocumentoDAL.guardarDocumento(objDocumento)
        Else
            objDocumentoDAL.actualizarDocumento(objDocumento)
        End If
        Return objDocumento.codigo
    End Function

End Class
