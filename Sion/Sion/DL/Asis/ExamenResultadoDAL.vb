Public Class ExamenResultadoDAL
    Public Shared Function guardarResultado(objResultado As ExamenResultado)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction(IsolationLevel.ReadUncommitted)
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = objResultado.cargarConsultaGuardar
                    comando.Parameters.Add(New SqlParameter("@Codigo_Orden", SqlDbType.Int)).Value = objResultado.CodigoOrden
                    comando.Parameters.Add(New SqlParameter("@codigo", SqlDbType.NVarChar)).Value = objResultado.CodigoProcedimiento
                    comando.Parameters.Add(New SqlParameter("@codigo_Tipo", SqlDbType.Int)).Value = objResultado.codigoTipo
                    comando.Parameters.Add(New SqlParameter("@Nota", SqlDbType.NVarChar)).Value = objResultado.observacion
                    comando.Parameters.Add(New SqlParameter("@Usuario_Real", SqlDbType.Int)).Value = objResultado.usuarioReal
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objResultado.usuario
                    comando.Parameters.Add(New SqlParameter("@editado", SqlDbType.Int)).Value = objResultado.editado
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objResultado.idEmpresa
                    comando.Parameters.Add(New SqlParameter("@ArchivoAudio", SqlDbType.VarBinary)).Value = objResultado.archivoAudio
                    comando.Parameters.Add(New SqlParameter("@Tabla", SqlDbType.Structured)).Value = objResultado.dtExamen
                    comando.Parameters.Add(New SqlParameter("@TablaEliminados", SqlDbType.Structured)).Value = objResultado.dtExamenEliminados
                    comando.ExecuteNonQuery()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objResultado
    End Function

    Public Shared Sub guardarResultadoLab(objResultado As ExamenResultado)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = "SP_RESULTADOS_LABORATORIO_CREAR"
                    comando.Parameters.Add(New SqlParameter("@Codigo_atencion", SqlDbType.Int)).Value = objResultado.CodigoOrden
                    comando.Parameters.Add(New SqlParameter("@codigo", SqlDbType.NVarChar)).Value = objResultado.CodigoProcedimiento
                    comando.Parameters.Add(New SqlParameter("@codigo_Tipo", SqlDbType.Int)).Value = objResultado.codigoTipo
                    comando.Parameters.Add(New SqlParameter("@Nota", SqlDbType.NVarChar)).Value = objResultado.observacion
                    comando.Parameters.Add(New SqlParameter("@Usuario_Real", SqlDbType.Int)).Value = objResultado.usuarioReal
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objResultado.usuario
                    comando.Parameters.Add(New SqlParameter("@editado", SqlDbType.Int)).Value = objResultado.editado
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objResultado.idEmpresa
                    comando.Parameters.Add(New SqlParameter("@ArchivoAudio", SqlDbType.VarBinary)).Value = objResultado.archivoAudio
                    comando.Parameters.Add(New SqlParameter("@tipo", SqlDbType.NVarChar)).Value = objResultado.tipoExam
                    comando.Parameters.Add(New SqlParameter("@Tabla", SqlDbType.Structured)).Value = objResultado.dtExamen
                    comando.Parameters.Add(New SqlParameter("@TablaEliminados", SqlDbType.Structured)).Value = objResultado.dtExamenEliminados
                    comando.ExecuteNonQuery()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub guardarResultadoExamenLab(objResultado As ExamenResultado)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = "SP_RESULTADOS_EXAMEN_LABORATORIO_CREAR"
                    comando.Parameters.Add(New SqlParameter("@Codigo_atencion", SqlDbType.Int)).Value = objResultado.CodigoOrden
                    comando.Parameters.Add(New SqlParameter("@codigo", SqlDbType.NVarChar)).Value = objResultado.CodigoProcedimiento
                    comando.Parameters.Add(New SqlParameter("@codigo_Tipo", SqlDbType.Int)).Value = objResultado.codigoTipo
                    comando.Parameters.Add(New SqlParameter("@Nota", SqlDbType.NVarChar)).Value = objResultado.observacion
                    comando.Parameters.Add(New SqlParameter("@Usuario_Real", SqlDbType.Int)).Value = objResultado.usuarioReal
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objResultado.usuario
                    comando.Parameters.Add(New SqlParameter("@editado", SqlDbType.Int)).Value = objResultado.editado
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objResultado.idEmpresa
                    comando.Parameters.Add(New SqlParameter("@ArchivoAudio", SqlDbType.VarBinary)).Value = objResultado.archivoAudio
                    comando.Parameters.Add(New SqlParameter("@tipo", SqlDbType.NVarChar)).Value = objResultado.tipoExam
                    comando.Parameters.Add(New SqlParameter("@Tabla", SqlDbType.Structured)).Value = objResultado.dtExamen
                    comando.Parameters.Add(New SqlParameter("@TablaEliminados", SqlDbType.Structured)).Value = objResultado.dtExamenEliminados
                    comando.ExecuteNonQuery()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
