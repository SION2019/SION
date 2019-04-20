Imports Celer

Public Class perfilDAL
    Public Shared Sub crearPerfil(pPerfil As TerceroPerfil)
        Try
            Using consulta = New SqlCommand()
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandText = "EXEC SP_PERFIL_USUARIO_CREAR @Nombre='" & pPerfil.nombre & "',@observaciones='" & pPerfil.observacion & "',@id_empresa='" & SesionActual.idEmpresa & "',
                                            @Usuario_Creacion='" & SesionActual.idUsuario & "'"
                pPerfil.codigoPerfil = CType(consulta.ExecuteScalar, String)
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub actualizarPerfil(pPerfil As TerceroPerfil)
        Try
            Using consulta = New SqlCommand()
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandText = "EXEC SP_PERFIL_USUARIO_ACTUALIZAR @Codigo_perfil='" & pPerfil.codigoPerfil & "',@Nombre='" & pPerfil.nombre & "',@observaciones='" & pPerfil.observacion & "',@id_empresa='" & SesionActual.idEmpresa & "',
                                           @Usuario_actualizacion='" & SesionActual.idUsuario & "'"
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Shared Sub anularPerfil(perfil As TerceroPerfil)
        Try
            Using consulta = New SqlCommand()
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandText = "EXEC SP_PERFIL_USUARIO_ANULAR @Usuario_actualizacion='" & SesionActual.idUsuario & "',@id_empresa='" & SesionActual.idEmpresa & "',@CODIGO_PERFIL='" & perfil.codigoPerfil & "'"
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
