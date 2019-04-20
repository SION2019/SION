Public Class CateterCentral
    Public Property Codigo As String
    Public Property registro As Integer
    Public Property fecha_insercion As Date
    Public Property fecha_retiro As Date
    Public Property causa As String
    Public Property observacion As String
    Public Property inotropicos As Boolean
    Public Property ntp As Boolean
    Public Property midazolan As Boolean
    Public Property anfotericina As Boolean
    Public Property solucion As Boolean
    Public Property otra As Boolean
    Public Property cual As String
    Public Property salacx As Boolean
    Public Property salaproc As Boolean
    Public Property otrasala As Boolean
    Public Property cualsala As String
    Public Property yugular As Boolean
    Public Property subclavio As Boolean
    Public Property femoral As Boolean
    Public Property mono As Boolean
    Public Property bi As Boolean
    Public Property tri As Boolean
    Public Property drum As Boolean
    Public Property epic As Boolean
    Public Property swain As Boolean
    Public Property larterial As Boolean
    Public Property otro As Boolean
    Public Property cualcat As String
    Public Property dtCateter As DataTable
    Public Property usuario As Integer
    Public Property usuarioRealiza As Integer


    Sub New()
        dtCateter = New DataTable
        dtCateter.Columns.Add("Codigo_CTC", Type.GetType("System.String"))
        dtCateter.Columns.Add("Codigo_criterio", Type.GetType("System.String"))
        dtCateter.Columns.Add("Respuesta", Type.GetType("System.String"))
    End Sub



End Class
