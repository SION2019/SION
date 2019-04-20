Imports System.Data.SqlClient
Public Class EpicrisisDAL
    Public Sub guardarEpicrisisHc(ByRef objEpicrisis As Epicrisis)
        Try
            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = "SP_EPICRISIS_CREAR"
                dbCommand.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                dbCommand.Parameters("@REGISTRO").Value = objEpicrisis.registro
                dbCommand.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.DateTime))
                dbCommand.Parameters("@FECHA").Value = objEpicrisis.fecha
                dbCommand.Parameters.Add(New SqlParameter("@NOTA", SqlDbType.NVarChar))
                dbCommand.Parameters("@NOTA").Value = objEpicrisis.nota
                dbCommand.Parameters.Add(New SqlParameter("@ESTADO_SALIDA", SqlDbType.Int))
                dbCommand.Parameters("@ESTADO_SALIDA").Value = objEpicrisis.estadoSalida
                dbCommand.Parameters.Add(New SqlParameter("@DESTINO", SqlDbType.Int))
                dbCommand.Parameters("@DESTINO").Value = objEpicrisis.destino
                dbCommand.Parameters.Add(New SqlParameter("@INSTITUCION", SqlDbType.Int))
                dbCommand.Parameters("@INSTITUCION").Value = objEpicrisis.institucion
                dbCommand.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                dbCommand.Parameters("@OBSERVACION").Value = objEpicrisis.observacion
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO_Real", SqlDbType.Int))
                dbCommand.Parameters("@USUARIO_Real").Value = objEpicrisis.usuarioReal
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO_CREACION", SqlDbType.Int))
                dbCommand.Parameters("@USUARIO_CREACION").Value = objEpicrisis.usuario
                dbCommand.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                dbCommand.Parameters("@CODIGO_EP").Value = objEpicrisis.codigoEp
                objEpicrisis.codigo = CType(dbCommand.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub actualizarEpicrisisHc(objEpicrisis As Epicrisis)
        Try
            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = "SP_EPICRISIS_ACTUALIZAR"
                dbCommand.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                dbCommand.Parameters("@REGISTRO").Value = objEpicrisis.registro
                dbCommand.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.DateTime))
                dbCommand.Parameters("@FECHA").Value = objEpicrisis.fecha
                dbCommand.Parameters.Add(New SqlParameter("@NOTA", SqlDbType.NVarChar))
                dbCommand.Parameters("@NOTA").Value = objEpicrisis.nota
                dbCommand.Parameters.Add(New SqlParameter("@ESTADO_SALIDA", SqlDbType.Int))
                dbCommand.Parameters("@ESTADO_SALIDA").Value = objEpicrisis.estadoSalida
                dbCommand.Parameters.Add(New SqlParameter("@DESTINO", SqlDbType.Int))
                dbCommand.Parameters("@DESTINO").Value = objEpicrisis.destino
                dbCommand.Parameters.Add(New SqlParameter("@INSTITUCION", SqlDbType.Int))
                dbCommand.Parameters("@INSTITUCION").Value = objEpicrisis.institucion
                dbCommand.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                dbCommand.Parameters("@OBSERVACION").Value = objEpicrisis.observacion
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO_ACTUALIZACION", SqlDbType.Int))
                dbCommand.Parameters("@USUARIO_ACTUALIZACION").Value = objEpicrisis.usuario
                dbCommand.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                dbCommand.Parameters("@CODIGO_EP").Value = objEpicrisis.codigoEp
                dbCommand.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub guardarEpicrisisAm(ByRef objEpicrisis As EpicrisisR)
        Try
            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = "SP_EPICRISIS_CREAR_R"
                dbCommand.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                dbCommand.Parameters("@REGISTRO").Value = objEpicrisis.registro
                dbCommand.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.DateTime))
                dbCommand.Parameters("@FECHA").Value = objEpicrisis.fecha
                dbCommand.Parameters.Add(New SqlParameter("@NOTA", SqlDbType.NVarChar))
                dbCommand.Parameters("@NOTA").Value = objEpicrisis.nota
                dbCommand.Parameters.Add(New SqlParameter("@ESTADO_SALIDA", SqlDbType.Int))
                dbCommand.Parameters("@ESTADO_SALIDA").Value = objEpicrisis.estadoSalida
                dbCommand.Parameters.Add(New SqlParameter("@DESTINO", SqlDbType.Int))
                dbCommand.Parameters("@DESTINO").Value = objEpicrisis.destino
                dbCommand.Parameters.Add(New SqlParameter("@INSTITUCION", SqlDbType.Int))
                dbCommand.Parameters("@INSTITUCION").Value = objEpicrisis.institucion
                dbCommand.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                dbCommand.Parameters("@OBSERVACION").Value = objEpicrisis.observacion
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO_Real", SqlDbType.Int))
                dbCommand.Parameters("@USUARIO_Real").Value = objEpicrisis.usuarioReal
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO_CREACION", SqlDbType.Int))
                dbCommand.Parameters("@USUARIO_CREACION").Value = objEpicrisis.usuario
                dbCommand.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                dbCommand.Parameters("@CODIGO_EP").Value = objEpicrisis.codigoEp
                objEpicrisis.codigo = CType(dbCommand.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub actualizarEpicrisisAm(objEpicrisis As EpicrisisR)
        Try
            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = "SP_EPICRISIS_ACTUALIZAR_R"
                dbCommand.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                dbCommand.Parameters("@REGISTRO").Value = objEpicrisis.registro
                dbCommand.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.DateTime))
                dbCommand.Parameters("@FECHA").Value = objEpicrisis.fecha
                dbCommand.Parameters.Add(New SqlParameter("@NOTA", SqlDbType.NVarChar))
                dbCommand.Parameters("@NOTA").Value = objEpicrisis.nota
                dbCommand.Parameters.Add(New SqlParameter("@ESTADO_SALIDA", SqlDbType.Int))
                dbCommand.Parameters("@ESTADO_SALIDA").Value = objEpicrisis.estadoSalida
                dbCommand.Parameters.Add(New SqlParameter("@DESTINO", SqlDbType.Int))
                dbCommand.Parameters("@DESTINO").Value = objEpicrisis.destino
                dbCommand.Parameters.Add(New SqlParameter("@INSTITUCION", SqlDbType.Int))
                dbCommand.Parameters("@INSTITUCION").Value = objEpicrisis.institucion
                dbCommand.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                dbCommand.Parameters("@OBSERVACION").Value = objEpicrisis.observacion
                dbCommand.Parameters.Add(New SqlParameter("@EDICION", SqlDbType.Bit))
                dbCommand.Parameters("@EDICION").Value = 1
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO_ACTUALIZACION", SqlDbType.Int))
                dbCommand.Parameters("@USUARIO_ACTUALIZACION").Value = objEpicrisis.usuario
                dbCommand.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                dbCommand.Parameters("@CODIGO_EP").Value = objEpicrisis.codigoEp
                dbCommand.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub guardarEpicrisisAf(ByRef objEpicrisis As EpicrisisRR)
        Try
            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = "SP_EPICRISIS_CREAR_RR"
                dbCommand.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                dbCommand.Parameters("@REGISTRO").Value = objEpicrisis.registro
                dbCommand.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.DateTime))
                dbCommand.Parameters("@FECHA").Value = objEpicrisis.fecha
                dbCommand.Parameters.Add(New SqlParameter("@NOTA", SqlDbType.NVarChar))
                dbCommand.Parameters("@NOTA").Value = objEpicrisis.nota
                dbCommand.Parameters.Add(New SqlParameter("@ESTADO_SALIDA", SqlDbType.Int))
                dbCommand.Parameters("@ESTADO_SALIDA").Value = objEpicrisis.estadoSalida
                dbCommand.Parameters.Add(New SqlParameter("@DESTINO", SqlDbType.Int))
                dbCommand.Parameters("@DESTINO").Value = objEpicrisis.destino
                dbCommand.Parameters.Add(New SqlParameter("@INSTITUCION", SqlDbType.Int))
                dbCommand.Parameters("@INSTITUCION").Value = objEpicrisis.institucion
                dbCommand.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                dbCommand.Parameters("@OBSERVACION").Value = objEpicrisis.observacion
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO_Real", SqlDbType.Int))
                dbCommand.Parameters("@USUARIO_Real").Value = objEpicrisis.usuarioReal
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO_CREACION", SqlDbType.Int))
                dbCommand.Parameters("@USUARIO_CREACION").Value = objEpicrisis.usuario
                objEpicrisis.codigo = CType(dbCommand.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub actualizarEpicrisisAf(objEpicrisis As EpicrisisRR)
        Try
            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = "SP_EPICRISIS_ACTUALIZAR_RR"
                dbCommand.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                dbCommand.Parameters("@REGISTRO").Value = objEpicrisis.registro
                dbCommand.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.DateTime))
                dbCommand.Parameters("@FECHA").Value = objEpicrisis.fecha
                dbCommand.Parameters.Add(New SqlParameter("@NOTA", SqlDbType.NVarChar))
                dbCommand.Parameters("@NOTA").Value = objEpicrisis.nota
                dbCommand.Parameters.Add(New SqlParameter("@ESTADO_SALIDA", SqlDbType.Int))
                dbCommand.Parameters("@ESTADO_SALIDA").Value = objEpicrisis.estadoSalida
                dbCommand.Parameters.Add(New SqlParameter("@DESTINO", SqlDbType.Int))
                dbCommand.Parameters("@DESTINO").Value = objEpicrisis.destino
                dbCommand.Parameters.Add(New SqlParameter("@INSTITUCION", SqlDbType.Int))
                dbCommand.Parameters("@INSTITUCION").Value = objEpicrisis.institucion
                dbCommand.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                dbCommand.Parameters("@OBSERVACION").Value = objEpicrisis.observacion
                dbCommand.Parameters.Add(New SqlParameter("@EDICION", SqlDbType.Bit))
                dbCommand.Parameters("@EDICION").Value = 1
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO_ACTUALIZACION", SqlDbType.Int))
                dbCommand.Parameters("@USUARIO_ACTUALIZACION").Value = objEpicrisis.usuario
                dbCommand.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
