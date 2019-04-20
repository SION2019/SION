Imports System.Data.SqlClient
Public Class TarifaServicioDAL
    Public Function guardarTarifa(ByVal objTarifaServicio As TarifaDeServicios) As String
        Try

            Using comando As New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.Parameters.Clear()
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "SP_TARIFA_SERVICIO_CREAR"
                comando.Parameters.Add(New SqlParameter("@Descripcion", SqlDbType.NVarChar))
                comando.Parameters("@Descripcion").Value = objTarifaServicio.descripcion
                comando.Parameters.Add(New SqlParameter("@Codigo_Tipo_Codigo", SqlDbType.Int))
                comando.Parameters("@Codigo_Tipo_Codigo").Value = objTarifaServicio.codigoTipoCodigo
                comando.Parameters.Add(New SqlParameter("@Codigo_ep", SqlDbType.Int))
                comando.Parameters("@Codigo_ep").Value = objTarifaServicio.codigoEp
                comando.Parameters.Add(New SqlParameter("@anho", SqlDbType.NVarChar))
                comando.Parameters("@anho").Value = objTarifaServicio.año
                comando.Parameters.Add(New SqlParameter("@Salario", SqlDbType.Decimal))
                comando.Parameters("@Salario").Value = objTarifaServicio.salario
                comando.Parameters.Add(New SqlParameter("@Porcentaje", SqlDbType.NVarChar))
                comando.Parameters("@Porcentaje").Value = objTarifaServicio.porcentaje
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                comando.Parameters("@USUARIO").Value = objTarifaServicio.usuarioCreacion
                objTarifaServicio.codigoHa = CType(comando.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception
            objTarifaServicio.codigoHa = ""
            Throw ex
        End Try
        Return objTarifaServicio.codigoHa
    End Function
    Public Sub actualizarTarifa(ByVal objTarifaServicio As TarifaDeServicios)
        Try
            Using consulta As New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.Parameters.Clear()
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TARIFA_SERVICIO_ACTUALIZAR"
                consulta.Parameters.Add(New SqlParameter("@Descripcion", SqlDbType.NVarChar))
                consulta.Parameters("@Descripcion").Value = objTarifaServicio.descripcion
                consulta.Parameters.Add(New SqlParameter("@Codigo_Tipo_Codigo", SqlDbType.Int))
                consulta.Parameters("@Codigo_Tipo_Codigo").Value = objTarifaServicio.codigoTipoCodigo
                consulta.Parameters.Add(New SqlParameter("@Codigo_ep", SqlDbType.Int))
                consulta.Parameters("@Codigo_ep").Value = objTarifaServicio.codigoEp
                consulta.Parameters.Add(New SqlParameter("@anho", SqlDbType.NVarChar))
                consulta.Parameters("@anho").Value = objTarifaServicio.año
                consulta.Parameters.Add(New SqlParameter("@Salario", SqlDbType.Decimal))
                consulta.Parameters("@Salario").Value = objTarifaServicio.salario
                consulta.Parameters.Add(New SqlParameter("@Porcentaje", SqlDbType.NVarChar))
                consulta.Parameters("@Porcentaje").Value = objTarifaServicio.porcentaje
                consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                consulta.Parameters("@USUARIO").Value = objTarifaServicio.usuarioCreacion
                consulta.Parameters.Add(New SqlParameter("@Codigo_ha", SqlDbType.NVarChar))
                consulta.Parameters("@Codigo_ha").Value = objTarifaServicio.codigoHa
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw
        End Try
    End Sub

End Class
