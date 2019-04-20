Imports System.Data.SqlClient
Public Class InsumoPeriodicidadDAL
    Public Shared Sub guardarInsumos(objPerInsumos As PeriodicInsumos)
        For indice = 0 To objPerInsumos.dtInsumos.Rows.Count - 1
            Try
                Using dbCommand As New SqlCommand
                    dbCommand.Connection = FormPrincipal.cnxion
                    dbCommand.Parameters.Clear()
                    dbCommand.CommandType = CommandType.StoredProcedure
                    dbCommand.CommandText = "SP_PERIODICIDAD_INSUMOS_CREAR"
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_Interno", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_Interno").Value = objPerInsumos.dtInsumos.Rows(indice).Item(0).ToString
                    dbCommand.Parameters.Add(New SqlParameter("@AreaServicio", SqlDbType.Int))
                    dbCommand.Parameters("@AreaServicio").Value = objPerInsumos.areaServicio
                    dbCommand.Parameters.Add(New SqlParameter("@Cantidad", SqlDbType.Int))
                    dbCommand.Parameters("@Cantidad").Value = objPerInsumos.dtInsumos.Rows(indice).Item(2).ToString
                    dbCommand.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int))
                    dbCommand.Parameters("@Usuario").Value = objPerInsumos.usuario
                    dbCommand.ExecuteNonQuery()

                End Using
            Catch ex As Exception
                Throw ex
            End Try
        Next
    End Sub
End Class
