Public Class InsumoFisioterapiaR
    Inherits InsumoFisioterapia

    Sub New()
        insumosFisiocarga = ConsultasHC.FISIOTERAPIA_INSUMO_CARGAR_R
        insumosFisioAnular = ConsultasHC.FISIOTERAPIA_INSUMO_ANULAR_R
        insumosFisioPeriodicidad = ConsultasHC.FISIOTERAPIA_INSUMO_PERIODICIDAD_R
        insumosFisioConfigAud = ConsultasHC.FISIOTERAPIA_INSUMO_CONFIGURACION_R
    End Sub


End Class
