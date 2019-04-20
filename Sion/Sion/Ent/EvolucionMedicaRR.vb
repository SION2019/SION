Public Class EvolucionMedicaRR
    Inherits EvolucionMedica
    Public Sub New()
        nombreReporte = ConstantesHC.NOMBRE_PDF_EVOLUCION_MEDICARR
        moduloReporte = Constantes.REPORTE_AF
        evolucionListar = Consultas.LISTAR_EVOLUCIONES_RR
        evolucionCargar = Consultas.EVOLUCION_MEDICA_CARGARRR
        diagnosticoCargar = Consultas.EVOLUCION_MEDICA_DIAGNOSTICA_CARGARRR
        consultaVerificar = Consultas.EVOLUCION_VERIFICARRR
        consultaUltimaEvolucion = Consultas.ADMISION_ULTIMA_EVO_RR
        paraclinicoCargar = Consultas.ORDEN_MEDICA_PARACLINICO_EVO_CARGARRR
    End Sub

    Public Overrides Sub guardarEvolucionMedica()
        Try
            redisenarTabla()
            HistoriaClinicaBLL.guardar_evolucionRR(Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Overrides Sub anularEvolucion()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoEvolucion)
        General.ejecutarSQL(Consultas.ANULAR_EVOLUCIONRR, params)
    End Sub
End Class
