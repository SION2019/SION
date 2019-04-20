Imports System.Data.SqlClient
Public Class EcocardiogramaDAL
    Public Shared Function guardarEcocardiograma(objEcocardiograma As Eco)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = objEcocardiograma.objEcocardiograma.sqlGuardarRegistro
                    comando.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objEcocardiograma.CodigoProcedimiento
                    comando.Parameters.Add(New SqlParameter("@IDECO", SqlDbType.Int)).Value = objEcocardiograma.codigo
                    comando.Parameters.Add(New SqlParameter("@Codigo_orden", SqlDbType.Int)).Value = objEcocardiograma.CodigoOrden
                    comando.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objEcocardiograma.usuario
                    comando.Parameters.Add(New SqlParameter("@USUARIO_R", SqlDbType.Int)).Value = objEcocardiograma.usuarioReal
                    comando.Parameters.Add(New SqlParameter("@FECHA_EX", SqlDbType.DateTime)).Value = objEcocardiograma.fechaRegistro
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objEcocardiograma.idEmpresa
                    comando.Parameters.Add(New SqlParameter("@Tabla", SqlDbType.Structured)).Value = objEcocardiograma.objEcocardiograma.dtParametro
                    comando.Parameters.Add(New SqlParameter("@editar", SqlDbType.Int)).Value = objEcocardiograma.objEcocardiograma.editado
                    comando.Parameters.Add(New SqlParameter("@VENTANA", SqlDbType.NVarChar)).Value = objEcocardiograma.objEcocardiograma.ventana
                    comando.Parameters.Add(New SqlParameter("@Descripcion", SqlDbType.NVarChar)).Value = objEcocardiograma.descripcionHallazgo
                    comando.Parameters.Add(New SqlParameter("@CONCLUSION", SqlDbType.NVarChar)).Value = objEcocardiograma.objEcocardiograma.conclusion
                    objEcocardiograma.codigo = CType(comando.ExecuteScalar(), String)
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objEcocardiograma
    End Function
    Public Shared Function guardarEcocardiogramaStres(objEcocardiogramaStres As Eco)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = objEcocardiogramaStres.objEcocardiogramaEstres.sqlGuardarRegistro
                    comando.Parameters.Add(New SqlParameter("@IDECO", SqlDbType.NVarChar)).Value = objEcocardiogramaStres.codigo
                    comando.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objEcocardiogramaStres.CodigoProcedimiento
                    comando.Parameters.Add(New SqlParameter("@Codigo_orden", SqlDbType.Int)).Value = objEcocardiogramaStres.CodigoOrden
                    comando.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objEcocardiogramaStres.usuario
                    comando.Parameters.Add(New SqlParameter("@USUARIO_R", SqlDbType.Int)).Value = objEcocardiogramaStres.usuarioReal
                    comando.Parameters.Add(New SqlParameter("@FECHA_EX", SqlDbType.DateTime)).Value = objEcocardiogramaStres.fechaRegistro
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objEcocardiogramaStres.idEmpresa
                    comando.Parameters.Add(New SqlParameter("@editar", SqlDbType.Int)).Value = objEcocardiogramaStres.objEcocardiograma.editado
                    comando.Parameters.Add(New SqlParameter("@Descripcion", SqlDbType.NVarChar)).Value = objEcocardiogramaStres.descripcionHallazgo
                    objEcocardiogramaStres.codigo = comando.ExecuteScalar
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objEcocardiogramaStres
    End Function
End Class
