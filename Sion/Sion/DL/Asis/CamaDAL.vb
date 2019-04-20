Imports System.Data.SqlClient
Public Class CamaDAL
    Public Shared Function guardarCama(objCama As ConfiguracionCama)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = Consultas.CAMA_CREAR
                    comando.Parameters.Add(New SqlParameter("@ID", SqlDbType.NVarChar)).Value = objCama.codigo
                    comando.Parameters.Add(New SqlParameter("@Descripcion", SqlDbType.NVarChar)).Value = objCama.descripcion
                    comando.Parameters.Add(New SqlParameter("@Codigo_area", SqlDbType.Int)).Value = objCama.codigoArea
                    comando.Parameters.Add(New SqlParameter("@Codigo_espec", SqlDbType.Int)).Value = objCama.codigoEspecialidad
                    comando.Parameters.Add(New SqlParameter("@piso", SqlDbType.NVarChar)).Value = objCama.piso
                    comando.Parameters.Add(New SqlParameter("@habitacion", SqlDbType.NVarChar)).Value = objCama.habitacion
                    comando.Parameters.Add(New SqlParameter("@Observacion", SqlDbType.NVarChar)).Value = objCama.observacion
                    comando.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objCama.usuario
                    objCama.codigo = CType(comando.ExecuteScalar, String)
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objCama
    End Function
End Class
