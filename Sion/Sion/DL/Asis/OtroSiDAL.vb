Imports System.Data.SqlClient
Public Class OtroSiDAL
    Public Function Guardarotrosi(ByRef pOtroSi As OtroSi) As String

        Dim codigor As String = ""
        Try

            Using comando As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    If pOtroSi.codigo <> "" Then
                        comando.Parameters.Clear()
                        comando.CommandType = CommandType.StoredProcedure
                        comando.CommandText = "SP_OTRO_SI_ACTUALIZAR"
                        comando.Parameters.Add(New SqlParameter("@codigo_contrato", SqlDbType.Int)).Value = pOtroSi.Codigo_contrato
                        comando.Parameters.Add(New SqlParameter("@Valor_si", SqlDbType.BigInt)).Value = pOtroSi.valor
                        comando.Parameters.Add(New SqlParameter("@Fecha_Ini_Si", SqlDbType.Date)).Value = pOtroSi.Fecha_Ini
                        comando.Parameters.Add(New SqlParameter("@Fecha_final_Si", SqlDbType.Date)).Value = pOtroSi.Fecha_Final
                        comando.Parameters.Add(New SqlParameter("@vigente", SqlDbType.Bit)).Value = pOtroSi.vigente
                        comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar)).Value = SesionActual.idUsuario
                        comando.Parameters.Add(New SqlParameter("@Codigo_si", SqlDbType.Int)).Value = pOtroSi.codigo
                        comando.ExecuteNonQuery()

                        codigor = pOtroSi.codigo

                    Else
                        comando.Parameters.Clear()
                        comando.CommandType = CommandType.StoredProcedure
                        comando.CommandText = "SP_OTRO_SI_CREAR"
                        comando.Parameters.Add(New SqlParameter("@codigo_contrato", SqlDbType.Int)).Value = pOtroSi.Codigo_contrato
                        comando.Parameters.Add(New SqlParameter("@Valor_si", SqlDbType.BigInt)).Value = pOtroSi.valor
                        comando.Parameters.Add(New SqlParameter("@Fecha_Ini_Si", SqlDbType.Date)).Value = pOtroSi.Fecha_Ini
                        comando.Parameters.Add(New SqlParameter("@Fecha_final_Si", SqlDbType.Date)).Value = pOtroSi.Fecha_Final
                        comando.Parameters.Add(New SqlParameter("@vigente", SqlDbType.Bit)).Value = pOtroSi.vigente
                        comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar)).Value = SesionActual.idUsuario
                        codigor = CType(comando.ExecuteScalar, Integer)

                    End If

                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        codigor = ""
                        trnsccion.Rollback()
                        general.manejoErrores(ex) 
                    End Try
                End Using
            End Using
        Catch ex As Exception
            codigor = ""
            general.manejoErrores(ex) 
        End Try
        Return codigor
    End Function
    Public Function otrosiAnular(ByVal codigo As String, ByVal usuario As String) As Boolean
        Try
            Using consulta As New SqlCommand("SP_OTRO_SI_ANULAR")
                consulta.Parameters.Clear()
                consulta.CommandType = CommandType.StoredProcedure
                consulta.Connection = FormPrincipal.cnxion
                consulta.Parameters.Add(New SqlParameter("@Codigo_si", SqlDbType.NVarChar))
                consulta.Parameters("@Codigo_si").Value = codigo
                consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.NVarChar))
                consulta.Parameters("@Usuario").Value = usuario
                consulta.ExecuteNonQuery()
            End Using
            Return True
        Catch ex As Exception
            general.manejoErrores(ex) 
            Return False
        End Try
    End Function
End Class
