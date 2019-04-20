Imports System.Data.SqlClient
Public Class PuntoServicioDAL

    Public Function guardar(ByRef pPuntoServicio As PuntoServicio) As Boolean
        Using comando = New SqlCommand()
            Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                comando.Connection = FormPrincipal.cnxion
                comando.Transaction = trnsccion
                comando.CommandType = CommandType.StoredProcedure
                If pPuntoServicio.codigo <> "" Then
                    comando.CommandText = "SP_PUNTO_DE_SERVICIO_ACTUALIZAR"

                    comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int)).Value = pPuntoServicio.codigo
                    comando.Parameters.Add(New SqlParameter("@observaciones", SqlDbType.NVarChar)).Value = pPuntoServicio.observaciones
                    comando.Parameters.Add(New SqlParameter("@Direccion", SqlDbType.NVarChar)).Value = pPuntoServicio.direccion
                    comando.Parameters.Add(New SqlParameter("@Nombre", SqlDbType.NVarChar)).Value = pPuntoServicio.nombre
                    comando.Parameters.Add(New SqlParameter("@Telefono", SqlDbType.NVarChar)).Value = pPuntoServicio.telefono
                    comando.Parameters.Add(New SqlParameter("@Correspondencia", SqlDbType.NVarChar)).Value = pPuntoServicio.correspondencia
                    comando.Parameters.Add(New SqlParameter("@Codigo_PAIS", SqlDbType.NVarChar)).Value = pPuntoServicio.Combopais
                    comando.Parameters.Add(New SqlParameter("@Codigo_dpto", SqlDbType.NVarChar)).Value = pPuntoServicio.Combodepartamento
                    comando.Parameters.Add(New SqlParameter("@codigo_municipio", SqlDbType.NVarChar)).Value = pPuntoServicio.Combociudad
                    comando.Parameters.Add(New SqlParameter("@Id_responsable", SqlDbType.Int)).Value = pPuntoServicio.id_responsable
                    comando.Parameters.Add(New SqlParameter("@Estado", SqlDbType.Bit)).Value = pPuntoServicio.activo
                    comando.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.NVarChar)).Value = SesionActual.idUsuario
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "SP_EMPRESA_PUNTO_AGREGAR"
                    comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int)).Value = pPuntoServicio.codigo
                    comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int)).Value = SesionActual.idEmpresa
                    comando.ExecuteNonQuery()
                Else
                    comando.Parameters.Clear()
                    comando.CommandText = "SP_PUNTO_DE_SERVICIO_CREAR"
                    comando.Parameters.Add(New SqlParameter("@observaciones", SqlDbType.NVarChar)).Value = pPuntoServicio.observaciones
                    comando.Parameters.Add(New SqlParameter("@Direccion", SqlDbType.NVarChar)).Value = pPuntoServicio.direccion
                    comando.Parameters.Add(New SqlParameter("@Nombre", SqlDbType.NVarChar)).Value = pPuntoServicio.nombre
                    comando.Parameters.Add(New SqlParameter("@Telefono", SqlDbType.NVarChar)).Value = pPuntoServicio.telefono
                    comando.Parameters.Add(New SqlParameter("@Correspondencia", SqlDbType.NVarChar)).Value = pPuntoServicio.correspondencia
                    comando.Parameters.Add(New SqlParameter("@Codigo_PAIS", SqlDbType.NVarChar)).Value = pPuntoServicio.Combopais
                    comando.Parameters.Add(New SqlParameter("@Codigo_dpto", SqlDbType.NVarChar)).Value = pPuntoServicio.Combodepartamento
                    comando.Parameters.Add(New SqlParameter("@codigo_municipio", SqlDbType.NVarChar)).Value = pPuntoServicio.Combociudad
                    comando.Parameters.Add(New SqlParameter("@Id_responsable", SqlDbType.Int)).Value = pPuntoServicio.id_responsable
                    comando.Parameters.Add(New SqlParameter("@Estado", SqlDbType.Bit)).Value = pPuntoServicio.activo
                    comando.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.NVarChar)).Value = SesionActual.idUsuario

                    pPuntoServicio.codigo = CType(comando.ExecuteScalar, String)
                    comando.Parameters.Clear()
                    comando.CommandText = "SP_EMPRESA_PUNTO_AGREGAR "
                    comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int)).Value = pPuntoServicio.codigo
                    comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int)).Value = SesionActual.idEmpresa
                    comando.ExecuteNonQuery()
                End If

                Try
                    trnsccion.Commit()
                    Return True

                Catch ex As Exception
                    Throw ex
                    Return False
                End Try
            End Using
        End Using
    End Function

    Public Function anular(codigo As String) As Boolean
        Dim repuesta As String
        Using consulta = New SqlCommand()
            Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                consulta.Connection = FormPrincipal.cnxion
                consulta.Transaction = trnsccion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_PUNTO_DE_SERVICIO_ANULAR"
                consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.NVarChar))
                consulta.Parameters("@Usuario_actualizacion").Value = SesionActual.idUsuario
                consulta.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                consulta.Parameters("@Codigo_Punto").Value = codigo
                repuesta = consulta.ExecuteScalar
                If repuesta = 1 Then
                    MsgBox("Este punto de servicio esta asignada a uno o varios empleados y empresas, no se puede anular.", MsgBoxStyle.Information)
                    Return False
                End If
                Try
                    trnsccion.Commit()
                    Return True
                Catch ex As Exception
                    Throw ex
                    trnsccion.Rollback()
                    Return False

                End Try
            End Using
        End Using
    End Function


End Class
