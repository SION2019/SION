Imports System.Data.SqlClient
Public Class EscalaMorseDAL

    Public Shared Sub guardarMorse(objEscala As EscalaMorse)

        Try

            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()

                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_ESCALA_MORSE_CREAR"
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = objEscala.registro
                    consulta.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.Date))
                    consulta.Parameters("@FECHA").Value = objEscala.dateFecha
                    consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                    consulta.Parameters("@USUARIO").Value = objEscala.usuario
                    consulta.ExecuteNonQuery()

                    consulta.Parameters.Clear()
                    consulta.CommandText = "SP_ESCALA_MORSE_DETALLE_CREAR"
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = objEscala.registro
                    consulta.Parameters.Add(New SqlParameter("@hc", SqlDbType.Int))
                    consulta.Parameters("@hc").Value = objEscala.l1
                    consulta.Parameters.Add(New SqlParameter("@ds", SqlDbType.Int))
                    consulta.Parameters("@ds").Value = objEscala.l2
                    consulta.Parameters.Add(New SqlParameter("@apd", SqlDbType.Int))
                    consulta.Parameters("@apd").Value = objEscala.l3
                    consulta.Parameters.Add(New SqlParameter("@v", SqlDbType.Int))
                    consulta.Parameters("@v").Value = objEscala.l4
                    consulta.Parameters.Add(New SqlParameter("@m", SqlDbType.Int))
                    consulta.Parameters("@m").Value = objEscala.l5
                    consulta.Parameters.Add(New SqlParameter("@em", SqlDbType.Int))
                    consulta.Parameters("@em").Value = objEscala.l6
                    consulta.ExecuteNonQuery()
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception

                        trnsccion.Rollback()
                        general.manejoErrores(ex) 
                    End Try
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
End Class
