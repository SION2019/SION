Imports System.Data.SqlClient
Public Class TrasladoPacienteDAL

    Public Shared Function traslado(objTraslado As TrasladoPacienteSede) As TrasladoPacienteSede
        Try
            Using Consulta As New SqlCommand(objTraslado.sqlTraslado)
                Consulta.Connection = FormPrincipal.cnxion
                Consulta.CommandType = CommandType.StoredProcedure
                Consulta.Parameters.Add(New SqlParameter("@N_RegistroOrigen", SqlDbType.Int)).Value = objTraslado.Registro
                Consulta.Parameters.Add(New SqlParameter("@codigoEpDestino", SqlDbType.Int)).Value = objTraslado.idPunto
                Consulta.Parameters.Add(New SqlParameter("@codigoContratoDestino", SqlDbType.Int)).Value = objTraslado.idContrato
                Consulta.Parameters.Add(New SqlParameter("@codigoCamaDestino", SqlDbType.Int)).Value = objTraslado.idCama
                Consulta.Parameters.Add(New SqlParameter("@codigoAreaServicioDestino", SqlDbType.Int)).Value = objTraslado.idArea
                Consulta.Parameters.Add(New SqlParameter("@N_RegistroDestino", SqlDbType.Int)).Direction = ParameterDirection.Output
                Consulta.Parameters.Add(New SqlParameter("@codigoRespuesta", SqlDbType.Int)).Direction = ParameterDirection.Output
                Consulta.ExecuteNonQuery()
                objTraslado.registroDestino = If(IsDBNull(Consulta.Parameters("@N_RegistroDestino").Value), 1,
                                                 Consulta.Parameters("@N_RegistroDestino").Value)
                objTraslado.respuesta = If(IsDBNull(Consulta.Parameters("@codigoRespuesta").Value), 1,
                                                  Consulta.Parameters("@codigoRespuesta").Value)
            End Using
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
        Return objTraslado
    End Function
End Class
