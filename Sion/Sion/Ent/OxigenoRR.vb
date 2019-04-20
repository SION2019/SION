Public Class OxigenoRR
    Inherits Oxigeno
    Sub New()
        modulo = Constantes.REPORTE_AF
        oxigenoCargarGuardados = ConsultasHC.OXIGENO_CARGAR_PRIMERA_VEZ_RR
        nombreReporte = ConstantesHC.NOMBRE_PDF_OXIGENO_FISIO_RR
        moduloReporte = Constantes.REPORTE_AF
        consultaFechaEgreso = ConsultasHC.OXIGENO_FECHA_EGRESO_RR
    End Sub

End Class
