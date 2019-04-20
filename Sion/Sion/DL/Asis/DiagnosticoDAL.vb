Imports System.Data.SqlClient
Public Class DiagnosticoDAL
    Public Shared Function guardarDiagnostico(objDiagnostico As ConfiguracionDiagnostico)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = Consultas.DIAGNOSTICO_CREAR
                    comando.Parameters.Add(New SqlParameter("@ID", SqlDbType.NVarChar)).Value = objDiagnostico.codigo
                    comando.Parameters.Add(New SqlParameter("@Descripcion", SqlDbType.NVarChar)).Value = objDiagnostico.descripcion
                    comando.Parameters.Add(New SqlParameter("@codigo_categoria", SqlDbType.NVarChar)).Value = objDiagnostico.codigoCategoria
                    comando.Parameters.Add(New SqlParameter("@genero", SqlDbType.Int)).Value = objDiagnostico.codigoGenero
                    comando.Parameters.Add(New SqlParameter("@edadinicial", SqlDbType.Int)).Value = objDiagnostico.edadInicial
                    comando.Parameters.Add(New SqlParameter("@edadfinal", SqlDbType.Int)).Value = objDiagnostico.edadFinal
                    comando.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objDiagnostico.usuario
                    objDiagnostico.codigo = CType(comando.ExecuteScalar, String)
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objDiagnostico
    End Function

End Class
