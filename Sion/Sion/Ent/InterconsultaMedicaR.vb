Public Class InterconsultaMedicaR
    Inherits InterconsultaMedica
    Public Sub New()
        nombreReporte = ConstantesHC.NOMBRE_PDF_INTERCONSULTAR
        moduloReporte = Constantes.REPORTE_AM
        interconsultaListar = Consultas.INTERCONSULTAS_LISTAR_R
        interconsultaCargar = Consultas.INTERCONSULTA_CARGARR
        consultaVerificar = Consultas.INTERCONSULTA_VERIFICARR
    End Sub

    Public Overrides Sub guardarinterconsultaMedica()
        Try
            HistoriaClinicaDAL.guardarInterconsultaR(Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub anularInterconsulta()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoOrden)
        params.Add(codigoProcedimiento)
        params.Add(codigoEP)
        General.ejecutarSQL(Consultas.ANULAR_INTERCONSULTAR, params)
    End Sub
End Class
