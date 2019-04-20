Public Class InterconsultaMedicaRR
    Inherits InterconsultaMedica
    Public Sub New()
        nombreReporte = ConstantesHC.NOMBRE_PDF_INTERCONSULTARR
        moduloReporte = Constantes.REPORTE_AF
        interconsultaListar = Consultas.INTERCONSULTAS_LISTAR_RR
        interconsultaCargar = Consultas.INTERCONSULTA_CARGARRR
        consultaVerificar = Consultas.INTERCONSULTA_VERIFICARRR
    End Sub

    Public Overrides Sub guardarinterconsultaMedica()
        Try
            HistoriaClinicaDAL.guardarInterconsultaRR(Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub anularInterconsulta()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoOrden)
        params.Add(codigoProcedimiento)
        General.ejecutarSQL(Consultas.ANULAR_INTERCONSULTARR, params)

    End Sub
End Class
