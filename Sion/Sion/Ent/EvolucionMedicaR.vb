Public Class EvolucionMedicaR
    Inherits EvolucionMedica
    Public Sub New()
        nombreReporte = ConstantesHC.NOMBRE_PDF_EVOLUCION_MEDICAR
        moduloReporte = Constantes.REPORTE_AM
        evolucionListar = Consultas.LISTAR_EVOLUCIONES_R
        evolucionCargar = Consultas.EVOLUCION_MEDICA_CARGARR
        diagnosticoCargar = Consultas.EVOLUCION_MEDICA_DIAGNOSTICA_CARGARR
        consultaVerificar = Consultas.EVOLUCION_VERIFICARR
        consultaUltimaEvolucion = Consultas.ADMISION_ULTIMA_EVO_R
        paraclinicoCargar = Consultas.ORDEN_MEDICA_PARACLINICO_EVO_CARGARR
    End Sub

    Public Overrides Sub guardarEvolucionMedica()
        Try
            redisenarTabla()
            HistoriaClinicaBLL.guardar_evolucionR(Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub anularEvolucion()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoEvolucion)
        params.Add(codigoEP)
        General.ejecutarSQL(Consultas.ANULAR_EVOLUCIONR, params)
        consultaVerificar = Consultas.EVOLUCION_VERIFICARR
    End Sub
End Class
