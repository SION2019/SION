Public Class InsumoEnfermeriaRR
    Inherits InsumoEnfermeria

    Sub New()
        insumosEnfercarga = ConsultasHC.ENFERMERIA_INSUMO_CARGAR_RR
        insumosEnferAnular = ConsultasHC.ENFERMERIA_INSUMO_ANULAR_RR
        insumosEnferPeriodicidad = ConsultasHC.ENFERMERIA_INSUMO_PERIODICIDAD_RR
        insumosEnferConfigAud = ConsultasHC.ENFERMERIA_INSUMO_CONFIGURACION_RR
    End Sub
    Public Overrides Sub anularOrden()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoOrden)
        General.ejecutarSQL(insumosEnferAnular, params)
    End Sub
End Class
