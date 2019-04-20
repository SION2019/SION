Public Class InsumoEnfermeriaR
    Inherits InsumoEnfermeria

    Sub New()
        insumosEnfercarga = ConsultasHC.ENFERMERIA_INSUMO_CARGAR_R
        insumosEnferAnular = ConsultasHC.ENFERMERIA_INSUMO_ANULAR_R
        insumosEnferPeriodicidad = ConsultasHC.ENFERMERIA_INSUMO_PERIODICIDAD_R
        insumosEnferConfigAud = ConsultasHC.ENFERMERIA_INSUMO_CONFIGURACION_R
    End Sub
End Class
