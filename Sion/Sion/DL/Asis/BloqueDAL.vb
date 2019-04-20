Public Class BloqueDAL
    Public Shared Function guardarBloque(objBloque As Bloque)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = objBloque.sqlGuardar
                    comando.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objBloque.codigoBloque
                    comando.Parameters.Add(New SqlParameter("@Codigo_Estante", SqlDbType.NVarChar)).Value = objBloque.codigoEstante
                    comando.Parameters.Add(New SqlParameter("@Bloque", SqlDbType.NVarChar)).Value = objBloque.Bloque
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objBloque.usuario
                    comando.ExecuteNonQuery()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objBloque
    End Function
End Class
