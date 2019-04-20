Public Class Justificacion_antibioticos

    Public Property codigo As String
    Public Property nRegistro As Integer
    Public Property codigoOrden As Integer
    Public Property nombreCompleto As String
    Public Property empresa As String
    Public Property Ingreso As String
    Public Property areaServicio As String
    Public Property medEspecial As String
    Public Property si As String
    Public Property no As String
    Public Property fechaMuestra As DateTime
    Public Property fecha As DateTime
    Public Property tipoMuestra As String
    Public Property aislamiento As String
    Public Property codigoInterno As Integer
    Public Property antibiotico As String
    Public Property tiempoUso As String
    Public Property justificacion As String
    Public Property usuario As Integer
    Public Property nombreUsuario As String
    Public Property dtDiagnosticos As DataTable
    Public Property dtJustificacion As DataTable
    Public Property codigoEP As Integer

    Sub New()
        dtDiagnosticos = New DataTable
        dtDiagnosticos.Columns.Add("Codigo", Type.GetType("System.String"))
        dtDiagnosticos.Columns.Add("Codigo_CIE", Type.GetType("System.String"))
        dtDiagnosticos.Columns.Add("Codigo_evo", Type.GetType("System.String"))


        dtJustificacion = New DataTable
        dtJustificacion.Columns.Add("Antibiotico", Type.GetType("System.String"))
        dtJustificacion.Columns.Add("Tiempo_uso", Type.GetType("System.String"))
    End Sub

End Class
