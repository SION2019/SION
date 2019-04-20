Public Class InfoIngresoNeonatoRR
    Inherits InfoIngresoNeonato
    Public Sub New()
        sqlDetalleCarga = Consultas.INFO_INGRESON_CARGARRR
        sqlDetalleImpresionCarga = Consultas.INFOIMPRESION_DIAGNOSTICA_CARGARRR
        sqlDetalleRemisionCarga = Consultas.INFOREMISION_DIAGNOSTICA_CARGARRR
        nombreReporte = ConstantesHC.NOMBRE_PDF_INGRESO_MEDICORR
        moduloReporte = Constantes.REPORTE_AF
    End Sub

    Public Overrides Sub guardarDetalle()
        HistoriaClinicaBLL.guardarInfoIngresoNRR(Me)
    End Sub
End Class
