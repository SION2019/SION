Imports System.Data.SqlClient
Public Class TransfusionSanguineaDAL

    Public Sub crearTransfusion(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_MED_CREAR"
                consulta.Parameters.Add(New SqlParameter("@Codigo_orden", SqlDbType.NVarChar)).Value = objTransfusion.CodigoOrden
                consulta.Parameters.Add(New SqlParameter("@codigo_Procedimiento", SqlDbType.NVarChar)).Value = objTransfusion.codigoProcedimiento
                consulta.Parameters.Add(New SqlParameter("@Fecha", SqlDbType.DateTime)).Value = objTransfusion.Fecha
                consulta.Parameters.Add(New SqlParameter("@Grupo_sanguineo", SqlDbType.NVarChar)).Value = objTransfusion.GrupoSanguineo
                consulta.Parameters.Add(New SqlParameter("@Rh", SqlDbType.NVarChar)).Value = objTransfusion.Rh
                consulta.Parameters.Add(New SqlParameter("@Transfundir", SqlDbType.Int)).Value = objTransfusion.Transfundir
                consulta.Parameters.Add(New SqlParameter("@Reserva", SqlDbType.NVarChar)).Value = objTransfusion.Reserva
                consulta.Parameters.Add(New SqlParameter("@sangre", SqlDbType.Bit)).Value = objTransfusion.sangre
                consulta.Parameters.Add(New SqlParameter("@Globulo_roja2", SqlDbType.Bit)).Value = objTransfusion.globulRoroja2
                consulta.Parameters.Add(New SqlParameter("@Globulo_roja", SqlDbType.Bit)).Value = objTransfusion.globuloRoja
                consulta.Parameters.Add(New SqlParameter("@Plasma", SqlDbType.Bit)).Value = objTransfusion.Plasma
                consulta.Parameters.Add(New SqlParameter("@Plaquetas", SqlDbType.Bit)).Value = objTransfusion.Plaquetas
                consulta.Parameters.Add(New SqlParameter("@Otros", SqlDbType.Bit)).Value = objTransfusion.Otros
                consulta.Parameters.Add(New SqlParameter("@Sala", SqlDbType.NVarChar)).Value = objTransfusion.Sala
                consulta.Parameters.Add(New SqlParameter("@Fecha_sala", SqlDbType.DateTime)).Value = objTransfusion.fechaSala
                consulta.Parameters.Add(New SqlParameter("@Transfusiones_previa", SqlDbType.Bit)).Value = objTransfusion.transfusionesPrevia
                consulta.Parameters.Add(New SqlParameter("@Cuantas", SqlDbType.NVarChar)).Value = objTransfusion.Cuantas
                consulta.Parameters.Add(New SqlParameter("@R_A_Sangre", SqlDbType.Bit)).Value = objTransfusion.RASangre
                consulta.Parameters.Add(New SqlParameter("@R_A_plasma", SqlDbType.Bit)).Value = objTransfusion.RAplasma
                consulta.Parameters.Add(New SqlParameter("@R_A_Otros", SqlDbType.Bit)).Value = objTransfusion.RAOtros
                consulta.Parameters.Add(New SqlParameter("@R_A_nacido_muerto", SqlDbType.Bit)).Value = objTransfusion.RAnacidoMuerto
                consulta.Parameters.Add(New SqlParameter("@R_A_Aborto", SqlDbType.Bit)).Value = objTransfusion.RAAborto
                consulta.Parameters.Add(New SqlParameter("@R_A_Enf_Hemolitica_RN_feto", SqlDbType.Bit)).Value = objTransfusion.RAEnfHemoliticaRNfeto
                consulta.Parameters.Add(New SqlParameter("@Diagnostico", SqlDbType.NVarChar)).Value = objTransfusion.Diagnostico
                consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objTransfusion.Usuario
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objTransfusion.codigoEp
                consulta.Parameters.Add(New SqlParameter("@usuarioMedico", SqlDbType.Int)).Value = objTransfusion.usuarioreal
                objTransfusion.CodigoTransfusion = CType(consulta.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub crearTransfusionR(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_MED_CREAR_R"
                consulta.Parameters.Add(New SqlParameter("@Codigo_orden", SqlDbType.NVarChar)).Value = objTransfusion.CodigoOrden
                consulta.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@codigo_Procedimiento", SqlDbType.NVarChar)).Value = objTransfusion.codigoProcedimiento
                consulta.Parameters.Add(New SqlParameter("@Fecha", SqlDbType.DateTime)).Value = objTransfusion.Fecha
                consulta.Parameters.Add(New SqlParameter("@Grupo_sanguineo", SqlDbType.NVarChar)).Value = objTransfusion.GrupoSanguineo
                consulta.Parameters.Add(New SqlParameter("@Rh", SqlDbType.NVarChar)).Value = objTransfusion.Rh
                consulta.Parameters.Add(New SqlParameter("@Transfundir", SqlDbType.Int)).Value = objTransfusion.Transfundir
                consulta.Parameters.Add(New SqlParameter("@Reserva", SqlDbType.NVarChar)).Value = objTransfusion.Reserva
                consulta.Parameters.Add(New SqlParameter("@sangre", SqlDbType.Bit)).Value = objTransfusion.sangre
                consulta.Parameters.Add(New SqlParameter("@Globulo_roja2", SqlDbType.Bit)).Value = objTransfusion.globulRoroja2
                consulta.Parameters.Add(New SqlParameter("@Globulo_roja", SqlDbType.Bit)).Value = objTransfusion.globuloRoja
                consulta.Parameters.Add(New SqlParameter("@Plasma", SqlDbType.Bit)).Value = objTransfusion.Plasma
                consulta.Parameters.Add(New SqlParameter("@Plaquetas", SqlDbType.Bit)).Value = objTransfusion.Plaquetas
                consulta.Parameters.Add(New SqlParameter("@Otros", SqlDbType.Bit)).Value = objTransfusion.Otros
                consulta.Parameters.Add(New SqlParameter("@Sala", SqlDbType.NVarChar)).Value = objTransfusion.Sala
                consulta.Parameters.Add(New SqlParameter("@Fecha_sala", SqlDbType.DateTime)).Value = objTransfusion.fechaSala
                consulta.Parameters.Add(New SqlParameter("@Transfusiones_previa", SqlDbType.Bit)).Value = objTransfusion.transfusionesPrevia
                consulta.Parameters.Add(New SqlParameter("@Cuantas", SqlDbType.NVarChar)).Value = objTransfusion.Cuantas
                consulta.Parameters.Add(New SqlParameter("@R_A_Sangre", SqlDbType.Bit)).Value = objTransfusion.RASangre
                consulta.Parameters.Add(New SqlParameter("@R_A_plasma", SqlDbType.Bit)).Value = objTransfusion.RAplasma
                consulta.Parameters.Add(New SqlParameter("@R_A_Otros", SqlDbType.Bit)).Value = objTransfusion.RAOtros
                consulta.Parameters.Add(New SqlParameter("@R_A_nacido_muerto", SqlDbType.Bit)).Value = objTransfusion.RAnacidoMuerto
                consulta.Parameters.Add(New SqlParameter("@R_A_Aborto", SqlDbType.Bit)).Value = objTransfusion.RAAborto
                consulta.Parameters.Add(New SqlParameter("@R_A_Enf_Hemolitica_RN_feto", SqlDbType.Bit)).Value = objTransfusion.RAEnfHemoliticaRNfeto
                consulta.Parameters.Add(New SqlParameter("@Diagnostico", SqlDbType.NVarChar)).Value = objTransfusion.Diagnostico
                consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objTransfusion.Usuario
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objTransfusion.codigoEp
                consulta.Parameters.Add(New SqlParameter("@usuarioMedico", SqlDbType.Int)).Value = objTransfusion.usuarioreal
                objTransfusion.CodigoTransfusion = CType(consulta.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub crearTransfusionRR(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_MED_CREAR_RR"
                consulta.Parameters.Add(New SqlParameter("@Codigo_orden", SqlDbType.NVarChar)).Value = objTransfusion.CodigoOrden
                consulta.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@codigo_Procedimiento", SqlDbType.NVarChar)).Value = objTransfusion.codigoProcedimiento
                consulta.Parameters.Add(New SqlParameter("@Fecha", SqlDbType.DateTime)).Value = objTransfusion.Fecha
                consulta.Parameters.Add(New SqlParameter("@Grupo_sanguineo", SqlDbType.NVarChar)).Value = objTransfusion.GrupoSanguineo
                consulta.Parameters.Add(New SqlParameter("@Rh", SqlDbType.NVarChar)).Value = objTransfusion.Rh
                consulta.Parameters.Add(New SqlParameter("@Transfundir", SqlDbType.Int)).Value = objTransfusion.Transfundir
                consulta.Parameters.Add(New SqlParameter("@Reserva", SqlDbType.NVarChar)).Value = objTransfusion.Reserva
                consulta.Parameters.Add(New SqlParameter("@sangre", SqlDbType.Bit)).Value = objTransfusion.sangre
                consulta.Parameters.Add(New SqlParameter("@Globulo_roja2", SqlDbType.Bit)).Value = objTransfusion.globulRoroja2
                consulta.Parameters.Add(New SqlParameter("@Globulo_roja", SqlDbType.Bit)).Value = objTransfusion.globuloRoja
                consulta.Parameters.Add(New SqlParameter("@Plasma", SqlDbType.Bit)).Value = objTransfusion.Plasma
                consulta.Parameters.Add(New SqlParameter("@Plaquetas", SqlDbType.Bit)).Value = objTransfusion.Plaquetas
                consulta.Parameters.Add(New SqlParameter("@Otros", SqlDbType.Bit)).Value = objTransfusion.Otros
                consulta.Parameters.Add(New SqlParameter("@Sala", SqlDbType.NVarChar)).Value = objTransfusion.Sala
                consulta.Parameters.Add(New SqlParameter("@Fecha_sala", SqlDbType.DateTime)).Value = objTransfusion.fechaSala
                consulta.Parameters.Add(New SqlParameter("@Transfusiones_previa", SqlDbType.Bit)).Value = objTransfusion.transfusionesPrevia
                consulta.Parameters.Add(New SqlParameter("@Cuantas", SqlDbType.NVarChar)).Value = objTransfusion.Cuantas
                consulta.Parameters.Add(New SqlParameter("@R_A_Sangre", SqlDbType.Bit)).Value = objTransfusion.RASangre
                consulta.Parameters.Add(New SqlParameter("@R_A_plasma", SqlDbType.Bit)).Value = objTransfusion.RAplasma
                consulta.Parameters.Add(New SqlParameter("@R_A_Otros", SqlDbType.Bit)).Value = objTransfusion.RAOtros
                consulta.Parameters.Add(New SqlParameter("@R_A_nacido_muerto", SqlDbType.Bit)).Value = objTransfusion.RAnacidoMuerto
                consulta.Parameters.Add(New SqlParameter("@R_A_Aborto", SqlDbType.Bit)).Value = objTransfusion.RAAborto
                consulta.Parameters.Add(New SqlParameter("@R_A_Enf_Hemolitica_RN_feto", SqlDbType.Bit)).Value = objTransfusion.RAEnfHemoliticaRNfeto
                consulta.Parameters.Add(New SqlParameter("@Diagnostico", SqlDbType.NVarChar)).Value = objTransfusion.Diagnostico
                consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objTransfusion.Usuario
                consulta.Parameters.Add(New SqlParameter("@usuarioMedico", SqlDbType.Int)).Value = objTransfusion.usuarioreal
                objTransfusion.CodigoTransfusion = CType(consulta.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub actualizarTransfusion(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_MED_ACTUALIZAR"
                consulta.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@Fecha", SqlDbType.DateTime)).Value = objTransfusion.Fecha
                consulta.Parameters.Add(New SqlParameter("@Grupo_sanguineo", SqlDbType.NVarChar)).Value = objTransfusion.GrupoSanguineo
                consulta.Parameters.Add(New SqlParameter("@Rh", SqlDbType.NVarChar)).Value = objTransfusion.Rh
                consulta.Parameters.Add(New SqlParameter("@Transfundir", SqlDbType.Int)).Value = objTransfusion.Transfundir
                consulta.Parameters.Add(New SqlParameter("@Reserva", SqlDbType.NVarChar)).Value = objTransfusion.Reserva
                consulta.Parameters.Add(New SqlParameter("@sangre", SqlDbType.Bit)).Value = objTransfusion.sangre
                consulta.Parameters.Add(New SqlParameter("@Globulo_roja2", SqlDbType.Bit)).Value = objTransfusion.globulRoroja2
                consulta.Parameters.Add(New SqlParameter("@Globulo_roja", SqlDbType.Bit)).Value = objTransfusion.globuloRoja
                consulta.Parameters.Add(New SqlParameter("@Plasma", SqlDbType.Bit)).Value = objTransfusion.Plasma
                consulta.Parameters.Add(New SqlParameter("@Plaquetas", SqlDbType.Bit)).Value = objTransfusion.Plaquetas
                consulta.Parameters.Add(New SqlParameter("@Otros", SqlDbType.Bit)).Value = objTransfusion.Otros
                consulta.Parameters.Add(New SqlParameter("@Sala", SqlDbType.NVarChar)).Value = objTransfusion.Sala
                consulta.Parameters.Add(New SqlParameter("@Fecha_sala", SqlDbType.DateTime)).Value = objTransfusion.fechaSala
                consulta.Parameters.Add(New SqlParameter("@Transfusiones_previa", SqlDbType.Bit)).Value = objTransfusion.transfusionesPrevia
                consulta.Parameters.Add(New SqlParameter("@Cuantas", SqlDbType.NVarChar)).Value = objTransfusion.Cuantas
                consulta.Parameters.Add(New SqlParameter("@R_A_Sangre", SqlDbType.Bit)).Value = objTransfusion.RASangre
                consulta.Parameters.Add(New SqlParameter("@R_A_plasma", SqlDbType.Bit)).Value = objTransfusion.RAplasma
                consulta.Parameters.Add(New SqlParameter("@R_A_Otros", SqlDbType.Bit)).Value = objTransfusion.RAOtros
                consulta.Parameters.Add(New SqlParameter("@R_A_nacido_muerto", SqlDbType.Bit)).Value = objTransfusion.RAnacidoMuerto
                consulta.Parameters.Add(New SqlParameter("@R_A_Aborto", SqlDbType.Bit)).Value = objTransfusion.RAAborto
                consulta.Parameters.Add(New SqlParameter("@R_A_Enf_Hemolitica_RN_feto", SqlDbType.Bit)).Value = objTransfusion.RAEnfHemoliticaRNfeto
                consulta.Parameters.Add(New SqlParameter("@Diagnostico", SqlDbType.NVarChar)).Value = objTransfusion.Diagnostico
                consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objTransfusion.Usuario
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objTransfusion.codigoEp
                consulta.Parameters.Add(New SqlParameter("@Usuario_real", SqlDbType.Int)).Value = objTransfusion.usuarioreal
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub actualizarTransfusionR(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_MED_ACTUALIZAR_R"
                consulta.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@Fecha", SqlDbType.DateTime)).Value = objTransfusion.Fecha
                consulta.Parameters.Add(New SqlParameter("@Grupo_sanguineo", SqlDbType.NVarChar)).Value = objTransfusion.GrupoSanguineo
                consulta.Parameters.Add(New SqlParameter("@Rh", SqlDbType.NVarChar)).Value = objTransfusion.Rh
                consulta.Parameters.Add(New SqlParameter("@Transfundir", SqlDbType.Int)).Value = objTransfusion.Transfundir
                consulta.Parameters.Add(New SqlParameter("@Reserva", SqlDbType.NVarChar)).Value = objTransfusion.Reserva
                consulta.Parameters.Add(New SqlParameter("@sangre", SqlDbType.Bit)).Value = objTransfusion.sangre
                consulta.Parameters.Add(New SqlParameter("@Globulo_roja2", SqlDbType.Bit)).Value = objTransfusion.globulRoroja2
                consulta.Parameters.Add(New SqlParameter("@Globulo_roja", SqlDbType.Bit)).Value = objTransfusion.globuloRoja
                consulta.Parameters.Add(New SqlParameter("@Plasma", SqlDbType.Bit)).Value = objTransfusion.Plasma
                consulta.Parameters.Add(New SqlParameter("@Plaquetas", SqlDbType.Bit)).Value = objTransfusion.Plaquetas
                consulta.Parameters.Add(New SqlParameter("@Otros", SqlDbType.Bit)).Value = objTransfusion.Otros
                consulta.Parameters.Add(New SqlParameter("@Sala", SqlDbType.NVarChar)).Value = objTransfusion.Sala
                consulta.Parameters.Add(New SqlParameter("@Fecha_sala", SqlDbType.DateTime)).Value = objTransfusion.fechaSala
                consulta.Parameters.Add(New SqlParameter("@Transfusiones_previa", SqlDbType.Bit)).Value = objTransfusion.transfusionesPrevia
                consulta.Parameters.Add(New SqlParameter("@Cuantas", SqlDbType.NVarChar)).Value = objTransfusion.Cuantas
                consulta.Parameters.Add(New SqlParameter("@R_A_Sangre", SqlDbType.Bit)).Value = objTransfusion.RASangre
                consulta.Parameters.Add(New SqlParameter("@R_A_plasma", SqlDbType.Bit)).Value = objTransfusion.RAplasma
                consulta.Parameters.Add(New SqlParameter("@R_A_Otros", SqlDbType.Bit)).Value = objTransfusion.RAOtros
                consulta.Parameters.Add(New SqlParameter("@R_A_nacido_muerto", SqlDbType.Bit)).Value = objTransfusion.RAnacidoMuerto
                consulta.Parameters.Add(New SqlParameter("@R_A_Aborto", SqlDbType.Bit)).Value = objTransfusion.RAAborto
                consulta.Parameters.Add(New SqlParameter("@R_A_Enf_Hemolitica_RN_feto", SqlDbType.Bit)).Value = objTransfusion.RAEnfHemoliticaRNfeto
                consulta.Parameters.Add(New SqlParameter("@Diagnostico", SqlDbType.NVarChar)).Value = objTransfusion.Diagnostico
                consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objTransfusion.Usuario
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objTransfusion.codigoEp
                consulta.Parameters.Add(New SqlParameter("@Usuario_real", SqlDbType.Int)).Value = objTransfusion.usuarioreal
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub actualizarTransfusionRR(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_MED_ACTUALIZAR_RR"
                consulta.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@Fecha", SqlDbType.DateTime)).Value = objTransfusion.Fecha
                consulta.Parameters.Add(New SqlParameter("@Grupo_sanguineo", SqlDbType.NVarChar)).Value = objTransfusion.GrupoSanguineo
                consulta.Parameters.Add(New SqlParameter("@Rh", SqlDbType.NVarChar)).Value = objTransfusion.Rh
                consulta.Parameters.Add(New SqlParameter("@Transfundir", SqlDbType.Int)).Value = objTransfusion.Transfundir
                consulta.Parameters.Add(New SqlParameter("@Reserva", SqlDbType.NVarChar)).Value = objTransfusion.Reserva
                consulta.Parameters.Add(New SqlParameter("@sangre", SqlDbType.Bit)).Value = objTransfusion.sangre
                consulta.Parameters.Add(New SqlParameter("@Globulo_roja2", SqlDbType.Bit)).Value = objTransfusion.globulRoroja2
                consulta.Parameters.Add(New SqlParameter("@Globulo_roja", SqlDbType.Bit)).Value = objTransfusion.globuloRoja
                consulta.Parameters.Add(New SqlParameter("@Plasma", SqlDbType.Bit)).Value = objTransfusion.Plasma
                consulta.Parameters.Add(New SqlParameter("@Plaquetas", SqlDbType.Bit)).Value = objTransfusion.Plaquetas
                consulta.Parameters.Add(New SqlParameter("@Otros", SqlDbType.Bit)).Value = objTransfusion.Otros
                consulta.Parameters.Add(New SqlParameter("@Sala", SqlDbType.NVarChar)).Value = objTransfusion.Sala
                consulta.Parameters.Add(New SqlParameter("@Fecha_sala", SqlDbType.DateTime)).Value = objTransfusion.fechaSala
                consulta.Parameters.Add(New SqlParameter("@Transfusiones_previa", SqlDbType.Bit)).Value = objTransfusion.transfusionesPrevia
                consulta.Parameters.Add(New SqlParameter("@Cuantas", SqlDbType.NVarChar)).Value = objTransfusion.Cuantas
                consulta.Parameters.Add(New SqlParameter("@R_A_Sangre", SqlDbType.Bit)).Value = objTransfusion.RASangre
                consulta.Parameters.Add(New SqlParameter("@R_A_plasma", SqlDbType.Bit)).Value = objTransfusion.RAplasma
                consulta.Parameters.Add(New SqlParameter("@R_A_Otros", SqlDbType.Bit)).Value = objTransfusion.RAOtros
                consulta.Parameters.Add(New SqlParameter("@R_A_nacido_muerto", SqlDbType.Bit)).Value = objTransfusion.RAnacidoMuerto
                consulta.Parameters.Add(New SqlParameter("@R_A_Aborto", SqlDbType.Bit)).Value = objTransfusion.RAAborto
                consulta.Parameters.Add(New SqlParameter("@R_A_Enf_Hemolitica_RN_feto", SqlDbType.Bit)).Value = objTransfusion.RAEnfHemoliticaRNfeto
                consulta.Parameters.Add(New SqlParameter("@Diagnostico", SqlDbType.NVarChar)).Value = objTransfusion.Diagnostico
                consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objTransfusion.Usuario
                consulta.Parameters.Add(New SqlParameter("@Usuario_real", SqlDbType.Int)).Value = objTransfusion.usuarioreal
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    '----------------------------------------------------------------------------------------------------------------------------------------
    'Laboratorio--------------------------------------------------------------------------------------------------------------------

    Public Sub actualizarLaboratorio(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_LAB_ACTUALIZAR"
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@DETALLE", SqlDbType.Structured)).Value = objTransfusion.dtLaboratorio
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objTransfusion.codigoEp
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub actualizarLaboratorioR(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_LAB_ACTUALIZAR_R"
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@DETALLE", SqlDbType.Structured)).Value = objTransfusion.dtLaboratorio
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objTransfusion.codigoEp
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub actualizarLaboratorioRR(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_LAB_ACTUALIZAR_RR"
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@DETALLE", SqlDbType.Structured)).Value = objTransfusion.dtLaboratorio
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    '----------------------------------------------------------------------------------------------------------------------------
    'Laboratorio2------------------------------------------------------------------------------------------------------------

    Public Sub actualizarLaboratorio2(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_LAB2_ACTUALIZAR"
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@RASTREO", SqlDbType.NVarChar)).Value = objTransfusion.rastreo
                consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar)).Value = objTransfusion.observaciones
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objTransfusion.codigoEp
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub actualizarLaboratorio2R(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_LAB2_ACTUALIZAR_R"
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@RASTREO", SqlDbType.NVarChar)).Value = objTransfusion.rastreo
                consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar)).Value = objTransfusion.observaciones
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objTransfusion.codigoEp
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub actualizarLaboratorio2RR(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_LAB2_ACTUALIZAR_RR"
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@rastreo", SqlDbType.NVarChar)).Value = objTransfusion.rastreo
                consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar)).Value = objTransfusion.observaciones
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    '--------------------------------------------------------------------------------------------------------------------------------------

    'Enfermeria-----------------------------------------------------------------------------------------------------------------------------


    Public Sub actualizarEnfermera(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_ENF_ACTUALIZAR"
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@Detalle_enf", SqlDbType.Structured)).Value = objTransfusion.dtEnfermera
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.NVarChar)).Value = objTransfusion.codigoEp
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub actualizarEnfermeraR(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_ENF_ACTUALIZAR_R"
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@Detalle_enf", SqlDbType.Structured)).Value = objTransfusion.dtEnfermera
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.NVarChar)).Value = objTransfusion.codigoEp
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub actualizarEnfermeraRR(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_ENF_ACTUALIZAR_RR"
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@Detalle_enf", SqlDbType.Structured)).Value = objTransfusion.dtEnfermera
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    '-------------------------------------------------------------------------------------------------------------------------------------------------
    'enfermera2---------------------------------------------------------------------------------------------------------------------------------------


    Public Sub actualizarEnfermera2(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_ENF2_ACTUALIZAR"
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@DETALLE", SqlDbType.Structured)).Value = objTransfusion.dtenfermera2
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.NVarChar)).Value = objTransfusion.codigoEp
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub actualizarEnfermera2R(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_ENF2_ACTUALIZAR_R"
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@DETALLE", SqlDbType.Structured)).Value = objTransfusion.dtenfermera2
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.NVarChar)).Value = objTransfusion.codigoEp
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub actualizarEnfermera2RR(objTransfusion As TransfusionSanguinea)
        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_TRANSFUSION_SANGUINEA_ENF2_ACTUALIZAR_RR"
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar)).Value = objTransfusion.CodigoTransfusion
                consulta.Parameters.Add(New SqlParameter("@DETALLE", SqlDbType.Structured)).Value = objTransfusion.dtenfermera2
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
End Class



