Imports System.Data.SqlClient
Public Class AntibioticoJustificacionDAL

    Public Shared Sub guardarAntibiotico(objAntibiotico As JustificacionMedicamento)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = objAntibiotico.guardarAntibiotico
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.NVarChar)).Value = objAntibiotico.codigoOrden
                comando.Parameters.Add(New SqlParameter("@codigo", SqlDbType.NVarChar)).Value = objAntibiotico.codigo
                comando.Parameters.Add(New SqlParameter("@Medespecial", SqlDbType.NVarChar)).Value = objAntibiotico.codigoInterno
                comando.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int)).Value = objAntibiotico.nRegistro
                comando.Parameters.Add(New SqlParameter("@AISLAMIENTO", SqlDbType.NVarChar)).Value = objAntibiotico.aislamiento
                comando.Parameters.Add(New SqlParameter("@ANTIBIOTICO", SqlDbType.NVarChar)).Value = objAntibiotico.antibiotico
                comando.Parameters.Add(New SqlParameter("@FECHAMUESTRA", SqlDbType.DateTime)).Value = objAntibiotico.fechaMuestra
                comando.Parameters.Add(New SqlParameter("@JUSTIFICACION", SqlDbType.NVarChar)).Value = objAntibiotico.justificacion
                comando.Parameters.Add(New SqlParameter("@TOMOMUESTRA", SqlDbType.Bit)).Value = objAntibiotico.timoMuestra
                comando.Parameters.Add(New SqlParameter("@TIEMPOUSO", SqlDbType.NVarChar)).Value = objAntibiotico.tiempoUso
                comando.Parameters.Add(New SqlParameter("@TIPOMUESTRA", SqlDbType.NVarChar)).Value = objAntibiotico.tipoMuestra
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objAntibiotico.usuario
                comando.Parameters.Add(New SqlParameter("@DETALLE_ANTIBIOTICO", SqlDbType.Structured)).Value = objAntibiotico.dtDiagnosticos
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objAntibiotico.codigoEP
                objAntibiotico.codigo = CType(comando.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub actualizarAntibiotico(objAntibiotico As JustificacionMedicamento)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = objAntibiotico.actualizarAntibiotico
                comando.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int)).Value = objAntibiotico.codigo
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objAntibiotico.codigoOrden
                comando.Parameters.Add(New SqlParameter("@AISLAMIENTO", SqlDbType.NVarChar)).Value = objAntibiotico.aislamiento
                comando.Parameters.Add(New SqlParameter("@FECHAMUESTRA", SqlDbType.DateTime)).Value = objAntibiotico.fechaMuestra
                comando.Parameters.Add(New SqlParameter("@JUSTIFICACION", SqlDbType.NVarChar)).Value = objAntibiotico.justificacion
                comando.Parameters.Add(New SqlParameter("@TOMOMUESTRA", SqlDbType.Bit)).Value = objAntibiotico.timoMuestra
                comando.Parameters.Add(New SqlParameter("@TIPOMUESTRA", SqlDbType.NVarChar)).Value = objAntibiotico.tipoMuestra
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objAntibiotico.usuario
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objAntibiotico.codigoEP
                comando.Parameters.Add(New SqlParameter("@DETALLE_ANTIBIOTICO", SqlDbType.Structured)).Value = objAntibiotico.dtDiagnosticos
                comando.ExecuteNonQuery()

            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


End Class
