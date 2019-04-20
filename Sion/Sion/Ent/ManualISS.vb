Public Class ManualISS
    Inherits ManualProcedimientos
    Public Sub New()
        sqlCargarProcedimiento = ConsultasAsis.MANUAL_ISS_CARGAR
    End Sub

    Public Overrides Sub guardarDetalle()
        Throw New NotImplementedException()
    End Sub
End Class
