Imports System.Data.SqlClient
Public Class CpisDAL

    Public Shared Function guardarCpis(ByVal N_Registro As String, ByVal Fecha As Date, ByVal USUARIO As String, dtCPIS As DataTable) As String

        Try

            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()

                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_CPIS_CREAR"
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = N_Registro
                    consulta.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.Date))
                    consulta.Parameters("@FECHA").Value = Fecha
                    consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                    consulta.Parameters("@USUARIO").Value = USUARIO
                    consulta.ExecuteNonQuery()

                    consulta.Parameters.Clear()
                    consulta.CommandText = "SP_CPIS_DETALLE_CREAR"
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = N_Registro
                    consulta.Parameters.Add(New SqlParameter("@Temperatura", SqlDbType.Int))
                    consulta.Parameters("@Temperatura").Value = dtCPIS.Rows(0).Item(4).ToString
                    consulta.Parameters.Add(New SqlParameter("@Leucocitos", SqlDbType.Int))
                    consulta.Parameters("@Leucocitos").Value = dtCPIS.Rows(1).Item(4).ToString
                    consulta.Parameters.Add(New SqlParameter("@traqueales", SqlDbType.Int))
                    consulta.Parameters("@traqueales").Value = dtCPIS.Rows(2).Item(4).ToString
                    consulta.Parameters.Add(New SqlParameter("@Oxigenacion", SqlDbType.Int))
                    consulta.Parameters("@Oxigenacion").Value = dtCPIS.Rows(3).Item(4).ToString
                    consulta.Parameters.Add(New SqlParameter("@Radiologia", SqlDbType.Int))
                    consulta.Parameters("@Radiologia").Value = dtCPIS.Rows(4).Item(4).ToString
                    consulta.ExecuteNonQuery()
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        N_Registro = ""
                        trnsccion.Rollback()
                        general.manejoErrores(ex) 
                    End Try
                End Using
            End Using
        Catch ex As Exception
            N_Registro = ""
            general.manejoErrores(ex) 
        End Try
        Return N_Registro
    End Function

End Class
