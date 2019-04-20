Public Class NotaEnfermeria
    Public Property codigoNota As String
    Public Property fechaNota As DateTime
    Public Property empresa As String
    Public Property editado As Boolean
    Public Property usuario As Integer
    Public Property usuarioNombre As String
    Public Property usuarioReal As Integer
    Public Property nRegistro As Integer
    Public Property nota As String
    Public Property codigoEP As Integer
    Public Property notaEnferAnular As String
    Sub New()
        notaEnferAnular = ConsultasHC.ENFERMERIA_NOTA_ANULAR
    End Sub
    Public Overridable Sub anularNota()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoNota)
        params.Add(codigoEP)
        General.ejecutarSQL(notaEnferAnular, params)
    End Sub
End Class
