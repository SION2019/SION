Public Class ListaCheckR
    Inherits ListaCheck
    Sub New()
        consultaPendiente = Consultas.LISTA_CHECK_R
        titulo = "LISTA DE TAREAS PENDIENTE: AUDITORIA MÉDICA"
    End Sub

    Public Overrides Sub cargarMenu()
        accion = Constantes.CARGARPADRE
        ListaCheckDAL.cargarMenuPadres(Me)

    End Sub
End Class
