Imports System.Data.SqlClient
Public Class EscalaBradenDAL
    Public Function guardarBraden(ByVal N_Registro As String, ByVal Fecha As Date, ByVal USUARIO As String, dtEscala As DataTable) As String

        Try

            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()

                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_ESCALA_BRADEN_CREAR"
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = N_Registro
                    consulta.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.Date))
                    consulta.Parameters("@FECHA").Value = Fecha
                    consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                    consulta.Parameters("@USUARIO").Value = USUARIO
                    consulta.ExecuteNonQuery()

                    consulta.Parameters.Clear()
                    consulta.CommandText = "SP_ESCALA_BRADEN_DETALLE_CREAR"
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = N_Registro
                    consulta.Parameters.Add(New SqlParameter("@SENSORIAL", SqlDbType.Int))
                    consulta.Parameters("@SENSORIAL").Value = dtEscala.Rows(0).Item(5).ToString
                    consulta.Parameters.Add(New SqlParameter("@HUMEDAD", SqlDbType.Int))
                    consulta.Parameters("@HUMEDAD").Value = dtEscala.Rows(1).Item(5).ToString
                    consulta.Parameters.Add(New SqlParameter("@ACTIVIDAD", SqlDbType.Int))
                    consulta.Parameters("@ACTIVIDAD").Value = dtEscala.Rows(2).Item(5).ToString
                    consulta.Parameters.Add(New SqlParameter("@MOVILIDAD", SqlDbType.Int))
                    consulta.Parameters("@MOVILIDAD").Value = dtEscala.Rows(3).Item(5).ToString
                    consulta.Parameters.Add(New SqlParameter("@NUTRICION", SqlDbType.Int))
                    consulta.Parameters("@NUTRICION").Value = dtEscala.Rows(4).Item(5).ToString
                    consulta.Parameters.Add(New SqlParameter("@DESLIZAMIENTO", SqlDbType.Int))
                    consulta.Parameters("@DESLIZAMIENTO").Value = dtEscala.Rows(5).Item(5).ToString
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
