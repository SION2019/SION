Public Class ManualCUPS
    Inherits ManualProcedimientos
    Public Sub New()
        sqlCargarProcedimiento = ConsultasAsis.MANUAL_CUPS_CARGAR
    End Sub

    Public Overrides Sub guardarDetalle()
        Throw New NotImplementedException()
    End Sub
End Class
