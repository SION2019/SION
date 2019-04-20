Public Class Buscar_Permisos_generales
    Dim permG As New BusquedaPermisoDAL
    Public Function buscarPermisoGeneral(ByVal nombre_form As String, Optional modulo As String = "") As String
        Return permG.buscar_permiso_general(nombre_form, modulo)
    End Function
End Class
