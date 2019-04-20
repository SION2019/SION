Public Class PrincipalBLL
    Dim objPrincipalDAL As New PrincipalDAL
    Public Sub guardarVersion(objVersion As Principal)
        objPrincipalDAL.guardarVersion(objVersion)
    End Sub

End Class
