Public Class ListaCheckRR
    Inherits ListaCheck
    Sub New()
        consultaPendiente = Consultas.LISTA_CHECK_RR
        titulo = "LISTA DE TAREAS PENDIENTE: AUDITORIA FACTURACIÓN"
    End Sub

    Public Overrides Sub cargarMenu()
        ListaCheckDAL.cargarMenuPadres(Me)

    End Sub
End Class
