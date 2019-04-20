Public Class NebulizacionRR
    Inherits Nebulizacion
    Sub New()
        modulo = Constantes.REPORTE_AF
        oxigenoCargarGuardados = ConsultasHC.OXIGENO_CARGAR_PRIMERA_VEZ_RR
        consultaCargar = ConsultasHC.NEBULIZACION_CARGAR_RR
        consultaFechaEgreso = ConsultasHC.OXIGENO_FECHA_EGRESO_RR
    End Sub


End Class
