Public Class PacienteDAL
    Public Shared Function guardarPaciente(objPaciente As Paciente)
        Try
            Using comando = New SqlCommand()
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.Parameters.Clear()
                comando.CommandText = Consultas.PACIENTES_CREAR
                comando.Parameters.Add(New SqlParameter("@ID", SqlDbType.NVarChar)).Value = objPaciente.codigo
                comando.Parameters.Add(New SqlParameter("@COD_INDENTI", SqlDbType.Int)).Value = objPaciente.codIdentificacion
                comando.Parameters.Add(New SqlParameter("@DOC_PACIENTE", SqlDbType.NVarChar)).Value = objPaciente.documentoPaciente
                comando.Parameters.Add(New SqlParameter("@COD_PAIS_EXP", SqlDbType.NVarChar)).Value = objPaciente.codPaisExp
                comando.Parameters.Add(New SqlParameter("@COD_DEPAR_EXP", SqlDbType.NVarChar)).Value = objPaciente.codDepartamentoExp
                comando.Parameters.Add(New SqlParameter("@COD_MUN_EXP", SqlDbType.NVarChar)).Value = objPaciente.codMunicipioExp
                comando.Parameters.Add(New SqlParameter("@FECHA_NAC", SqlDbType.Date)).Value = objPaciente.fechaNacimiento
                comando.Parameters.Add(New SqlParameter("@COD_PAIS_NAC", SqlDbType.NVarChar)).Value = objPaciente.codPaisNac
                comando.Parameters.Add(New SqlParameter("@COD_DEPAR_NAC", SqlDbType.NVarChar)).Value = objPaciente.codDepartamentoNac
                comando.Parameters.Add(New SqlParameter("@COD_MUN_NAC", SqlDbType.NVarChar)).Value = objPaciente.codMunicipioNac
                comando.Parameters.Add(New SqlParameter("@PAPELLIDO", SqlDbType.NVarChar)).Value = objPaciente.primerApellido
                comando.Parameters.Add(New SqlParameter("@SAPELLIDO", SqlDbType.NVarChar)).Value = objPaciente.segundoApellido
                comando.Parameters.Add(New SqlParameter("@PNOMBRE", SqlDbType.NVarChar)).Value = objPaciente.primerNombre
                comando.Parameters.Add(New SqlParameter("@SNOMBRE", SqlDbType.NVarChar)).Value = objPaciente.segundoNombre
                comando.Parameters.Add(New SqlParameter("@COD_ESTADO", SqlDbType.Int)).Value = objPaciente.codEstadoCivil
                comando.Parameters.Add(New SqlParameter("@COD_GENERO", SqlDbType.Int)).Value = objPaciente.codGenero
                comando.Parameters.Add(New SqlParameter("@COD_EPS", SqlDbType.Int)).Value = objPaciente.codEPS
                comando.Parameters.Add(New SqlParameter("@COD_TIPO_AFIL", SqlDbType.Int)).Value = objPaciente.codTipoAfiliacion
                comando.Parameters.Add(New SqlParameter("@COD_OCUPAC", SqlDbType.Int)).Value = objPaciente.codOcupacion
                comando.Parameters.Add(New SqlParameter("@COD_CLASE_SOC", SqlDbType.Int)).Value = objPaciente.codClaseSocial
                comando.Parameters.Add(New SqlParameter("@COD_ZONA", SqlDbType.Int)).Value = objPaciente.codZona
                comando.Parameters.Add(New SqlParameter("@DIRECCION", SqlDbType.NVarChar)).Value = objPaciente.direccionPaciente
                comando.Parameters.Add(New SqlParameter("@COD_PAIS_RES", SqlDbType.NVarChar)).Value = objPaciente.codPaisRes
                comando.Parameters.Add(New SqlParameter("@COD_DEPAR_RES", SqlDbType.NVarChar)).Value = objPaciente.codDepartamentoRes
                comando.Parameters.Add(New SqlParameter("@COD_MUN_RES", SqlDbType.NVarChar)).Value = objPaciente.codMunicipioRes
                comando.Parameters.Add(New SqlParameter("@TELEFONO", SqlDbType.NVarChar)).Value = objPaciente.telefonoPaciente
                comando.Parameters.Add(New SqlParameter("@CELULAR", SqlDbType.NVarChar)).Value = objPaciente.celularPaciente
                comando.Parameters.Add(New SqlParameter("@USUARIO_CREACION", SqlDbType.Int)).Value = objPaciente.usuario
                comando.Parameters.Add(New SqlParameter("@CORREO", SqlDbType.NVarChar)).Value = objPaciente.correo
                objPaciente.codigo = CType(comando.ExecuteScalar, String)
            End Using
        Catch ex As Exception
            Throw
        End Try
        Return objPaciente
    End Function

End Class
