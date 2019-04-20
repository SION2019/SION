Imports System.Data.SqlClient
Public Class CateterCentralDAL
    Public Function crearCateter(ByVal objCateter As CateterCentral) As String
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "SP_CATETER_CENTRAL_CREAR"
                comando.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int)).Value = objCateter.registro
                comando.Parameters.Add(New SqlParameter("@Fecha_insercion", SqlDbType.Date)).Value = objCateter.fecha_insercion
                comando.Parameters.Add(New SqlParameter("@Fecha_retiro", SqlDbType.Date)).Value = objCateter.fecha_retiro
                comando.Parameters.Add(New SqlParameter("@Causa", SqlDbType.NVarChar)).Value = objCateter.causa
                comando.Parameters.Add(New SqlParameter("@Observacion", SqlDbType.NVarChar)).Value = objCateter.observacion
                comando.Parameters.Add(New SqlParameter("@INOTROPICOS", SqlDbType.Bit)).Value = objCateter.inotropicos
                comando.Parameters.Add(New SqlParameter("@NTP", SqlDbType.Bit)).Value = objCateter.ntp
                comando.Parameters.Add(New SqlParameter("@MIDAZOLAM", SqlDbType.Bit)).Value = objCateter.midazolan
                comando.Parameters.Add(New SqlParameter("@ANFOTERICINA", SqlDbType.Bit)).Value = objCateter.anfotericina
                comando.Parameters.Add(New SqlParameter("@SOLUCION", SqlDbType.Bit)).Value = objCateter.solucion
                comando.Parameters.Add(New SqlParameter("@OTRA", SqlDbType.Bit)).Value = objCateter.otra
                comando.Parameters.Add(New SqlParameter("@CUAL", SqlDbType.NVarChar)).Value = objCateter.cual
                comando.Parameters.Add(New SqlParameter("@SALACX", SqlDbType.Bit)).Value = objCateter.salacx
                comando.Parameters.Add(New SqlParameter("@SALAPROC", SqlDbType.Bit)).Value = objCateter.salaproc
                comando.Parameters.Add(New SqlParameter("@OTRASALA", SqlDbType.Bit)).Value = objCateter.otrasala
                comando.Parameters.Add(New SqlParameter("@CUALSALA", SqlDbType.NVarChar)).Value = objCateter.cualsala
                comando.Parameters.Add(New SqlParameter("@YUGULAR", SqlDbType.Bit)).Value = objCateter.yugular
                comando.Parameters.Add(New SqlParameter("@SUBCLAVIO", SqlDbType.Bit)).Value = objCateter.subclavio
                comando.Parameters.Add(New SqlParameter("@FEMORAL", SqlDbType.Bit)).Value = objCateter.femoral
                comando.Parameters.Add(New SqlParameter("@MONO", SqlDbType.Bit)).Value = objCateter.mono
                comando.Parameters.Add(New SqlParameter("@BI", SqlDbType.Bit)).Value = objCateter.bi
                comando.Parameters.Add(New SqlParameter("@TRI", SqlDbType.Bit)).Value = objCateter.tri
                comando.Parameters.Add(New SqlParameter("@DRUM", SqlDbType.Bit)).Value = objCateter.drum
                comando.Parameters.Add(New SqlParameter("@EPIC", SqlDbType.Bit)).Value = objCateter.epic
                comando.Parameters.Add(New SqlParameter("@SWAN", SqlDbType.Bit)).Value = objCateter.swain
                comando.Parameters.Add(New SqlParameter("@LARTERIAL", SqlDbType.Bit)).Value = objCateter.larterial
                comando.Parameters.Add(New SqlParameter("@OTRO", SqlDbType.Bit)).Value = objCateter.otro
                comando.Parameters.Add(New SqlParameter("@CUALCAT", SqlDbType.NVarChar)).Value = objCateter.cualcat
                comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objCateter.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIO_REALIZA", SqlDbType.NVarChar)).Value = objCateter.usuarioRealiza
                comando.Parameters.Add(New SqlParameter("@detalle_cateter", SqlDbType.Structured)).Value = objCateter.dtCateter
                objCateter.Codigo = CType(comando.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objCateter.Codigo
    End Function

    Public Function actualizarCateter(ByVal objCateter As CateterCentral) As String
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "SP_CATETER_CENTRAL_ACTUALIZAR"
                comando.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objCateter.Codigo
                comando.Parameters.Add(New SqlParameter("@Fecha_insercion", SqlDbType.Date)).Value = objCateter.fecha_insercion
                comando.Parameters.Add(New SqlParameter("@Fecha_retiro", SqlDbType.Date)).Value = objCateter.fecha_retiro
                comando.Parameters.Add(New SqlParameter("@Causa", SqlDbType.NVarChar)).Value = objCateter.causa
                comando.Parameters.Add(New SqlParameter("@Observacion", SqlDbType.NVarChar)).Value = objCateter.observacion
                comando.Parameters.Add(New SqlParameter("@INOTROPICOS", SqlDbType.Bit)).Value = objCateter.inotropicos
                comando.Parameters.Add(New SqlParameter("@NTP", SqlDbType.Bit)).Value = objCateter.ntp
                comando.Parameters.Add(New SqlParameter("@MIDAZOLAM", SqlDbType.Bit)).Value = objCateter.midazolan
                comando.Parameters.Add(New SqlParameter("@ANFOTERICINA", SqlDbType.Bit)).Value = objCateter.anfotericina
                comando.Parameters.Add(New SqlParameter("@SOLUCION", SqlDbType.Bit)).Value = objCateter.solucion
                comando.Parameters.Add(New SqlParameter("@OTRA", SqlDbType.Bit)).Value = objCateter.otra
                comando.Parameters.Add(New SqlParameter("@CUAL", SqlDbType.NVarChar)).Value = objCateter.cual
                comando.Parameters.Add(New SqlParameter("@SALACX", SqlDbType.Bit)).Value = objCateter.salacx
                comando.Parameters.Add(New SqlParameter("@SALAPROC", SqlDbType.Bit)).Value = objCateter.salaproc
                comando.Parameters.Add(New SqlParameter("@OTRASALA", SqlDbType.Bit)).Value = objCateter.otrasala
                comando.Parameters.Add(New SqlParameter("@CUALSALA", SqlDbType.NVarChar)).Value = objCateter.cualsala
                comando.Parameters.Add(New SqlParameter("@YUGULAR", SqlDbType.Bit)).Value = objCateter.yugular
                comando.Parameters.Add(New SqlParameter("@SUBCLAVIO", SqlDbType.Bit)).Value = objCateter.subclavio
                comando.Parameters.Add(New SqlParameter("@FEMORAL", SqlDbType.Bit)).Value = objCateter.femoral
                comando.Parameters.Add(New SqlParameter("@MONO", SqlDbType.Bit)).Value = objCateter.mono
                comando.Parameters.Add(New SqlParameter("@BI", SqlDbType.Bit)).Value = objCateter.bi
                comando.Parameters.Add(New SqlParameter("@TRI", SqlDbType.Bit)).Value = objCateter.tri
                comando.Parameters.Add(New SqlParameter("@DRUM", SqlDbType.Bit)).Value = objCateter.drum
                comando.Parameters.Add(New SqlParameter("@EPIC", SqlDbType.Bit)).Value = objCateter.epic
                comando.Parameters.Add(New SqlParameter("@SWAN", SqlDbType.Bit)).Value = objCateter.swain
                comando.Parameters.Add(New SqlParameter("@LARTERIAL", SqlDbType.Bit)).Value = objCateter.larterial
                comando.Parameters.Add(New SqlParameter("@OTRO", SqlDbType.Bit)).Value = objCateter.otro
                comando.Parameters.Add(New SqlParameter("@CUALCAT", SqlDbType.NVarChar)).Value = objCateter.cualcat
                comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objCateter.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIO_REALIZA", SqlDbType.NVarChar)).Value = objCateter.usuarioRealiza
                comando.Parameters.Add(New SqlParameter("@detalle_cateter", SqlDbType.Structured)).Value = objCateter.dtCateter
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objCateter.Codigo
    End Function
End Class
