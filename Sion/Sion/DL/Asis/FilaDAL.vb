Public Class FilaDAL
    Public Shared Function guardarFila(objFila As Fila)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = objFila.sqlGuardar
                    comando.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objFila.codigoFila
                    comando.Parameters.Add(New SqlParameter("@Codigo_Estante", SqlDbType.NVarChar)).Value = objFila.codigoEstante
                    comando.Parameters.Add(New SqlParameter("@Codigo_Bloque", SqlDbType.NVarChar)).Value = objFila.codigoBloque
                    comando.Parameters.Add(New SqlParameter("@Fila", SqlDbType.NVarChar)).Value = objFila.fila
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objFila.usuario
                    comando.ExecuteNonQuery()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objFila
    End Function
End Class
