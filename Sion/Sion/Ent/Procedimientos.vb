Public MustInherit Class Procedimientos
    Public Property area As String
    Public Property formula As String
    Public Property codigo As String
    Public Property modulo As String
    Public Property nRegistro As Integer
    Public Property CodigoOrden As String
    Public Property CodigoProcedimiento As String
    Public Property idEmpresa As Integer
    Public Property usuario As Integer
    Public Property usuarioReal As Integer
    Public Property sqlGuardarRegistro As String
    Public Property sqlBuscarRegistro As String
    Public Property sqlCargarRegistro As String
    Public Property sqlAnularRegistro As String
    Public Property sqlBuscarPaciente As String
    Public Property sqlCargarPaciente As String
    Public Property sqlCargarInsumoConfig As String
    Public Property sqlCargarParaclinicoConfig As String
    Public Property titulo As String
    Public Property fechaRegistro As DateTime
    Public Property editado As Integer
    Public Property codigoArea As Integer
    Public Property codigoEp As Integer
    Public Property banderaGuardar As Boolean
    Public Sub New()

    End Sub
    Public MustOverride Sub guardarRegistro()

End Class
