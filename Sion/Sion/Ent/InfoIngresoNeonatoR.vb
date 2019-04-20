Public Class InfoIngresoNeonatoR
    Inherits InfoIngresoNeonato
    Public Sub New()
        sqlDetalleCarga = Consultas.INFO_INGRESON_CARGARR
        sqlDetalleImpresionCarga = Consultas.INFOIMPRESION_DIAGNOSTICA_CARGARR
        sqlDetalleRemisionCarga = Consultas.INFOREMISION_DIAGNOSTICA_CARGARR
        nombreReporte = ConstantesHC.NOMBRE_PDF_INGRESO_MEDICOR
        moduloReporte = Constantes.REPORTE_AM
    End Sub

    Public Overrides Sub guardarDetalle()
        HistoriaClinicaBLL.guardarInfoIngresoNR(Me)
    End Sub
End Class
