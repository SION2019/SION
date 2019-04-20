Imports System.Data.SqlClient
Public Class ConfigProcedExamDAL
    Public Function Guardarconfiguracionprocedimientos(objProcedimiento As ConfigReferenciaExam)
        Try
            Using comando As New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Connection = FormPrincipal.cnxion
                    comando.CommandText = Consultas.REFERENCIA_EXAMEN_CONF_CREAR
                    comando.Parameters.Add(New SqlParameter("@CODIGO_GRUP", SqlDbType.Int)).Value = objProcedimiento.CodigoTipo
                    comando.Parameters.Add(New SqlParameter("@CODIGO_GENERO", SqlDbType.Int)).Value = objProcedimiento.CodigoGenero
                    comando.Parameters.Add(New SqlParameter("@TABLA", SqlDbType.Structured)).Value = objProcedimiento.dtReferenciaExam
                    comando.ExecuteNonQuery()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw
        End Try
        Return objProcedimiento
    End Function
End Class
