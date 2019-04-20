Imports System.Data.SqlClient
Public Class PerfilModuloDAL


    Public Sub actualizarPermisos(moduloPerfil As ModuloPerfil, pUsuario As Integer)
        Try
            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = "SP_PUC_DETALLE_ACTUALIZAR"

                'Parametros
                dbCommand.Parameters.Add(New SqlParameter("@CodigoMenu", SqlDbType.NVarChar))
                dbCommand.Parameters.Add(New SqlParameter("@Descripcion", SqlDbType.NVarChar))
                dbCommand.Parameters.Add(New SqlParameter("@Formulario", SqlDbType.NVarChar))
                dbCommand.Parameters.Add(New SqlParameter("@MenuPadre", SqlDbType.Int))


                'Valores
                dbCommand.Parameters("@CodigoMenu").Value = moduloPerfil.getCodigoMenu
                dbCommand.Parameters("@Descripcion").Value = moduloPerfil.getDescripcion
                dbCommand.Parameters("@Formulario").Value = moduloPerfil.getFormulario
                dbCommand.Parameters("@MenuPadre").Value = moduloPerfil.getMenuPadre


                dbCommand.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub cargarMenuPadre(ByVal pCodigo As String,
                                  ByRef dsCuentas As DataSet)

        Try
            Using dbCommand As New SqlCommand("SP_MENU_PADRE_CARGAR " & pCodigo, FormPrincipal.cnxion)
                Using daCuentaPadre As New SqlDataAdapter(dbCommand)
                    daCuentaPadre.Fill(dsCuentas, "Padre")
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub cargarMenuHijas(ByVal pCodigo As String,
                                  ByRef dsCuentas As DataSet)

        Try
            Using dbCommand As New SqlCommand("SP_MENU_HIJO_CARGAR " & pCodigo, FormPrincipal.cnxion)
                Using daCuentaHija As New SqlDataAdapter(dbCommand)
                    daCuentaHija.Fill(dsCuentas, "Hijas")
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub



End Class
