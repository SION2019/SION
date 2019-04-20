Public Class NebulizacionR
    Inherits Nebulizacion
    Sub New()
        modulo = Constantes.REPORTE_AM
        oxigenoCargarGuardados = ConsultasHC.OXIGENO_CARGAR_PRIMERA_VEZ_R
        consultaCargar = ConsultasHC.NEBULIZACION_CARGAR_R
        consultaFechaEgreso = ConsultasHC.OXIGENO_FECHA_EGRESO_R
    End Sub

End Class
