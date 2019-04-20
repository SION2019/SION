Public Class InsumoFisioterapiaRR
    Inherits InsumoFisioterapia

    Sub New()
        insumosFisiocarga = ConsultasHC.FISIOTERAPIA_INSUMO_CARGAR_RR
        insumosFisioAnular = ConsultasHC.FISIOTERAPIA_INSUMO_ANULAR_RR
        insumosFisioPeriodicidad = ConsultasHC.FISIOTERAPIA_INSUMO_PERIODICIDAD_RR
        insumosFisioConfigAud = ConsultasHC.FISIOTERAPIA_INSUMO_CONFIGURACION_RR
    End Sub
    Public Overrides Sub anularOrden()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoOrden)
        General.ejecutarSQL(insumosFisioAnular, params)
    End Sub
End Class
