Public Class ElementoMenu

    Public Property codigo As String
    Public Property nombre As String
    Public Property modulo As String
    Public Property nombrePadre As String

    Sub New()

    End Sub

    Sub New(pCodigo As String, pNombre As String, pModulo As String, pNombrePadre As String)
        codigo = pCodigo
        nombre = pNombre
        modulo = pModulo
        nombrePadre = pNombrePadre
    End Sub

End Class
