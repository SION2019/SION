Public Class PuntoServicio

    Property codigo As String
    Property observaciones As String
    Property direccion As String
    Property nombre As String
    Property telefono As String
    Property correspondencia As String
    Property Combopais As String
    Property Combodepartamento As String
    Property Combociudad As String
    Property id_responsable As String
    Property activo As Boolean

    Public ReadOnly Property puntoNombre As String
        Get
            Return codigo & " | " & nombre
        End Get
    End Property

End Class
