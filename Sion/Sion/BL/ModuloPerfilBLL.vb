Public Class ModuloPerfilBLL

    Dim objModulo_perfil_c As New PerfilModuloDAL

    Public Function actualizarPermisos(ByVal objMenu As ModuloPerfil,
                                         ByVal pUsuario As Integer) As ModuloPerfil

        objModulo_perfil_c.actualizarPermisos(objMenu, pUsuario)

        Return objMenu
    End Function

    Public Sub cargarMenu(pcodigoEP As Integer, ByRef dsCuentas As DataSet)

        objModulo_perfil_c.cargarMenuPadre(pcodigoEP, dsCuentas)
        objModulo_perfil_c.cargarMenuHijas(pcodigoEP, dsCuentas)

    End Sub

End Class
