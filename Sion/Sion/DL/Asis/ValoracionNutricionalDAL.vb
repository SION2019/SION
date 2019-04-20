Imports System.Data.SqlClient
Public Class ValoracionNutricionalDAL

    Public Function crearValoracion(objValoracion As ValoracionNutricional) As String
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "SP_V_NUTRICIONAL_CREAR"
                comando.Parameters.Add(New SqlParameter("@Registro", SqlDbType.NVarChar)).Value = objValoracion.Registro
                comando.Parameters.Add(New SqlParameter("@Fechavaloracion", SqlDbType.DateTime)).Value = objValoracion.Fechavaloracion
                comando.Parameters.Add(New SqlParameter("@PesoA", SqlDbType.NVarChar)).Value = objValoracion.PesoA
                comando.Parameters.Add(New SqlParameter("@Talla", SqlDbType.NVarChar)).Value = objValoracion.Talla
                comando.Parameters.Add(New SqlParameter("@pt", SqlDbType.NVarChar)).Value = objValoracion.pt
                comando.Parameters.Add(New SqlParameter("@pe", SqlDbType.NVarChar)).Value = objValoracion.pe
                comando.Parameters.Add(New SqlParameter("@te", SqlDbType.NVarChar)).Value = objValoracion.te
                comando.Parameters.Add(New SqlParameter("@DiagN", SqlDbType.NVarChar)).Value = objValoracion.DiagN
                comando.Parameters.Add(New SqlParameter("@PresD", SqlDbType.NVarChar)).Value = objValoracion.PresD
                comando.Parameters.Add(New SqlParameter("@AnamA", SqlDbType.NVarChar)).Value = objValoracion.AnamA
                comando.Parameters.Add(New SqlParameter("@AlimP", SqlDbType.NVarChar)).Value = objValoracion.AlimP
                comando.Parameters.Add(New SqlParameter("@CaloR", SqlDbType.NVarChar)).Value = objValoracion.CaloR
                comando.Parameters.Add(New SqlParameter("@Prot", SqlDbType.NVarChar)).Value = objValoracion.Prot
                comando.Parameters.Add(New SqlParameter("@AlimR", SqlDbType.NVarChar)).Value = objValoracion.AlimR
                comando.Parameters.Add(New SqlParameter("@Gras", SqlDbType.NVarChar)).Value = objValoracion.Gras
                comando.Parameters.Add(New SqlParameter("@CHO", SqlDbType.NVarChar)).Value = objValoracion.CHO
                comando.Parameters.Add(New SqlParameter("@ExamL", SqlDbType.NVarChar)).Value = objValoracion.ExamL
                comando.Parameters.Add(New SqlParameter("@IntoA", SqlDbType.NVarChar)).Value = objValoracion.IntoA
                comando.Parameters.Add(New SqlParameter("@Nutricionista", SqlDbType.NVarChar)).Value = objValoracion.Nutricionista
                comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.NVarChar)).Value = objValoracion.Usuario
                objValoracion.CodigoValoracion = CType(comando.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objValoracion.CodigoValoracion
    End Function

    Public Function actualizarValoracion(objValoracion As ValoracionNutricional) As String
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "SP_V_NUTRICIONAL_ACTUALIZAR"
                comando.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objValoracion.CodigoValoracion
                comando.Parameters.Add(New SqlParameter("@Fechavaloracion", SqlDbType.DateTime)).Value = objValoracion.Fechavaloracion
                comando.Parameters.Add(New SqlParameter("@PesoA", SqlDbType.NVarChar)).Value = objValoracion.PesoA
                comando.Parameters.Add(New SqlParameter("@Talla", SqlDbType.NVarChar)).Value = objValoracion.Talla
                comando.Parameters.Add(New SqlParameter("@pt", SqlDbType.NVarChar)).Value = objValoracion.pt
                comando.Parameters.Add(New SqlParameter("@pe", SqlDbType.NVarChar)).Value = objValoracion.pe
                comando.Parameters.Add(New SqlParameter("@te", SqlDbType.NVarChar)).Value = objValoracion.te
                comando.Parameters.Add(New SqlParameter("@DiagN", SqlDbType.NVarChar)).Value = objValoracion.DiagN
                comando.Parameters.Add(New SqlParameter("@PresD", SqlDbType.NVarChar)).Value = objValoracion.PresD
                comando.Parameters.Add(New SqlParameter("@AnamA", SqlDbType.NVarChar)).Value = objValoracion.AnamA
                comando.Parameters.Add(New SqlParameter("@AlimP", SqlDbType.NVarChar)).Value = objValoracion.AlimP
                comando.Parameters.Add(New SqlParameter("@CaloR", SqlDbType.NVarChar)).Value = objValoracion.CaloR
                comando.Parameters.Add(New SqlParameter("@Prot", SqlDbType.NVarChar)).Value = objValoracion.Prot
                comando.Parameters.Add(New SqlParameter("@AlimR", SqlDbType.NVarChar)).Value = objValoracion.AlimR
                comando.Parameters.Add(New SqlParameter("@Gras", SqlDbType.NVarChar)).Value = objValoracion.Gras
                comando.Parameters.Add(New SqlParameter("@CHO", SqlDbType.NVarChar)).Value = objValoracion.CHO
                comando.Parameters.Add(New SqlParameter("@ExamL", SqlDbType.NVarChar)).Value = objValoracion.ExamL
                comando.Parameters.Add(New SqlParameter("@IntoA", SqlDbType.NVarChar)).Value = objValoracion.IntoA
                comando.Parameters.Add(New SqlParameter("@Nutricionista", SqlDbType.NVarChar)).Value = objValoracion.Nutricionista
                comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.NVarChar)).Value = objValoracion.Usuario
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objValoracion.CodigoValoracion
    End Function
End Class
