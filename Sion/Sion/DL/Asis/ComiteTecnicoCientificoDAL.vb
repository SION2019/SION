Imports System.Data.SqlClient
Public Class ComiteTecnicoCientificoDAL

    Public Shared Function guardarCTC(objComite As ComiteCTC)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = objComite.sqlGuardarRegistro
                    comando.Parameters.Add(New SqlParameter("@Codigo_Interno", SqlDbType.Int)).Value = objComite.dtMedicamento.Rows(0).Item("Codigo")
                    comando.Parameters.Add(New SqlParameter("@IDCTC", SqlDbType.NVarChar)).Value = objComite.codigo
                    comando.Parameters.Add(New SqlParameter("@codigo_solic", SqlDbType.NVarChar)).Value = objComite.Codigo_Solic
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objComite.usuario
                    comando.Parameters.Add(New SqlParameter("@USUARIO_REAL", SqlDbType.Int)).Value = objComite.usuarioReal
                    comando.Parameters.Add(New SqlParameter("@Fecha_CTC", SqlDbType.DateTime)).Value = objComite.fechaRegistro
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objComite.codigoEp
                    comando.Parameters.Add(New SqlParameter("@Codigo_Pem", SqlDbType.NVarChar)).Value = objComite.Codigo_PEM
                    comando.Parameters.Add(New SqlParameter("@editado", SqlDbType.Bit)).Value = objComite.editado
                    comando.Parameters.Add(New SqlParameter("@Encabezado", SqlDbType.NVarChar)).Value = objComite.Encabezado
                    comando.Parameters.Add(New SqlParameter("@Dosis", SqlDbType.Int)).Value = objComite.dtMedicamento.Rows(0).Item("dosis")
                    comando.Parameters.Add(New SqlParameter("@Cant", SqlDbType.Int)).Value = objComite.dtMedicamento.Rows(0).Item("cant")
                    comando.Parameters.Add(New SqlParameter("@Dias", SqlDbType.Int)).Value = objComite.dtMedicamento.Rows(0).Item("dias")
                    comando.Parameters.Add(New SqlParameter("@Decision", SqlDbType.NVarChar)).Value = objComite.Decision
                    comando.Parameters.Add(New SqlParameter("@DIAG_D", SqlDbType.Structured)).Value = objComite.dtdiag
                    comando.Parameters.Add(New SqlParameter("@ASIS_D", SqlDbType.Structured)).Value = objComite.dtAsistentes
                    objComite.codigo = comando.ExecuteScalar()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objComite
    End Function

End Class
