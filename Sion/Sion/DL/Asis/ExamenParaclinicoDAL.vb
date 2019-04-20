Imports System.Data.SqlClient
Public Class ExamenParaclinicoDAL
    Public Shared Function guardarParaclinicoLab(objParaclinicoLab As ParaclinicoLaboratorio)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = objParaclinicoLab.sqlExamenesGuardar
                    comando.Parameters.Add(New SqlParameter("@Codigo_orden", SqlDbType.Int)).Value = objParaclinicoLab.codigoOrden
                    comando.Parameters.Add(New SqlParameter("@Codigo_Procedimiento", SqlDbType.NVarChar)).Value = objParaclinicoLab.codigoProcedimiento
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objParaclinicoLab.usuario
                    comando.Parameters.Add(New SqlParameter("@Usuario_Real", SqlDbType.Int)).Value = objParaclinicoLab.usuarioReal
                    comando.Parameters.Add(New SqlParameter("@FechaReal", SqlDbType.DateTime)).Value = objParaclinicoLab.fechaReal
                    comando.Parameters.Add(New SqlParameter("@FechaMuestra", SqlDbType.DateTime)).Value = objParaclinicoLab.fechaMuestra
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objParaclinicoLab.codigoEP
                    comando.Parameters.Add(New SqlParameter("@Tabla", SqlDbType.Structured)).Value = objParaclinicoLab.dtExamen
                    comando.Parameters.Add(New SqlParameter("@editar", SqlDbType.Int)).Value = objParaclinicoLab.editado
                    comando.Parameters.Add(New SqlParameter("@tipo_Examen", SqlDbType.NVarChar)).Value = objParaclinicoLab.areaExamen
                    comando.ExecuteNonQuery()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objParaclinicoLab
    End Function

    Public Shared Function AnularParaclinicoLab(objParaclinicoLab As ParaclinicoLaboratorio)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = objParaclinicoLab.sqlExamenesAnular
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objParaclinicoLab.usuario
                    comando.Parameters.Add(New SqlParameter("@CodigoEP", SqlDbType.Int)).Value = objParaclinicoLab.codigoEP
                    comando.Parameters.Add(New SqlParameter("@Tabla", SqlDbType.Structured)).Value = objParaclinicoLab.dtExamen
                    comando.Parameters.Add(New SqlParameter("@editado", SqlDbType.Int)).Value = objParaclinicoLab.editado
                    comando.ExecuteNonQuery()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objParaclinicoLab
    End Function

End Class
