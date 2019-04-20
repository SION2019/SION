Public Class EmpresaBLL
    Dim objEmpresaDAL As New EmpresaDAL
    Public Function guardarEmpresa(ByVal objEmpresa As Empresa)
        If String.IsNullOrEmpty(objEmpresa.idEmpresa) Then
            Return objEmpresaDAL.guardarEmpresa(objEmpresa)
        Else
            objEmpresaDAL.actualizarEmpresa(objEmpresa)
        End If
        Return objEmpresa.idEmpresa
    End Function

End Class
