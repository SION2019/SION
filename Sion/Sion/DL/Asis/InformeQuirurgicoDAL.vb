Imports System.Data.SqlClient
Public Class InformeQuirurgicoDAL
    Public Shared Function guardarInformeQx(objInformeQx As InformeQx)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = objInformeQx.sqlGuardarRegistro
                    comando.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objInformeQx.CodigoProcedimiento
                    comando.Parameters.Add(New SqlParameter("@IDQX", SqlDbType.NVarChar)).Value = objInformeQx.codigo
                    comando.Parameters.Add(New SqlParameter("@Codigo_orden", SqlDbType.Int)).Value = objInformeQx.CodigoOrden
                    comando.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objInformeQx.usuario
                    comando.Parameters.Add(New SqlParameter("@USUARIO_REAL", SqlDbType.Int)).Value = objInformeQx.usuarioReal
                    comando.Parameters.Add(New SqlParameter("@Fecha_ini", SqlDbType.DateTime)).Value = objInformeQx.fechaInicio
                    comando.Parameters.Add(New SqlParameter("@Fecha_fin", SqlDbType.DateTime)).Value = objInformeQx.fechaFin
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objInformeQx.idEmpresa
                    comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Int)).Value = objInformeQx.editado
                    comando.Parameters.Add(New SqlParameter("@Codigo_ayudante", SqlDbType.Int)).Value = objInformeQx.codigoAyudante
                    comando.Parameters.Add(New SqlParameter("@Codigo_anestesiologo", SqlDbType.NVarChar)).Value = objInformeQx.setGetCodigoMedico
                    comando.Parameters.Add(New SqlParameter("@Codigo_tipo_anestesia", SqlDbType.NVarChar)).Value = objInformeQx.setGetCodigoMedicamento
                    comando.Parameters.Add(New SqlParameter("@codigo_Via", SqlDbType.Int)).Value = objInformeQx.codigoVia
                    comando.Parameters.Add(New SqlParameter("@Hallazgo", SqlDbType.NVarChar)).Value = objInformeQx.hallazgo
                    comando.Parameters.Add(New SqlParameter("@Descripcion", SqlDbType.NVarChar)).Value = objInformeQx.descripcionProcedimiento
                    comando.Parameters.Add(New SqlParameter("@Conducta", SqlDbType.NVarChar)).Value = objInformeQx.conductaSeguir
                    comando.Parameters.Add(New SqlParameter("@Justificacion", SqlDbType.NVarChar)).Value = objInformeQx.justificacionProcedimiento
                    comando.Parameters.Add(New SqlParameter("@TABLA_MEDICAMENTO", SqlDbType.Structured)).Value = objInformeQx.dtMedicamento
                    comando.Parameters.Add(New SqlParameter("@TABLA_PROCED", SqlDbType.Structured)).Value = objInformeQx.dtProcedimiento
                    comando.Parameters.Add(New SqlParameter("@TABLA_DIAG", SqlDbType.Structured)).Value = objInformeQx.dtDiagpos
                    comando.Parameters.Add(New SqlParameter("@TABLA_DIAG_PRE", SqlDbType.Structured)).Value = objInformeQx.dtDiagPre
                    objInformeQx.codigo = comando.ExecuteScalar()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objInformeQx
    End Function

End Class
