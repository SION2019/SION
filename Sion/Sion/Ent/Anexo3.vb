Public Class Anexo3
    Property orden As String
    Property registro As String
    Property nsolicitud As String
    Property tipoanexo As String
    Property manejo As String
    Property justificacion As String
    Property tiposervicio As Integer
    Property prioridad As Integer
    Property ingreso As Integer
    Property servicio As String
    Property servicioNuevo As String
    Property remision As String
    Property usuario As String
    Property dtanexo As DataTable
    Property elementoAEliminar As List(Of String)
    Property codigoNW As String


    Public Sub guardarAnexo3()
        If String.IsNullOrEmpty(remision) Then
            remision = -1
        End If
        Anexo3DAL.GuardarAnexo3(Me)
    End Sub
End Class
