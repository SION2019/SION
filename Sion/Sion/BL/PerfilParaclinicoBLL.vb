Public Class PerfilParaclinicoBLL
    Public Shared Sub guardarPerfilParaclinico(ObjPerfilP As PerfilParaclinico)
        PerfilParaclinicoDAL.guardarPerfilParaclinico(ObjPerfilP)
    End Sub
    Public Function guardarConfPerfilParaclinicoBLL(objConfPerfilP As ConfiguracionPerfilParaclinico)
        PerfilParaclinicoDAL.GuardarconfiguracionprocedimientosDAL(objConfPerfilP)
        Return objConfPerfilP
    End Function
End Class
