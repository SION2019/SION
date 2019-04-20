Public Class ManualSOAT
    Inherits ManualProcedimientos
    Public Sub New()
        sqlCargarProcedimiento = ConsultasAsis.MANUAL_SOAT_CARGAR
    End Sub

    Public Overrides Sub guardarDetalle()
        Throw New NotImplementedException()
    End Sub
End Class
