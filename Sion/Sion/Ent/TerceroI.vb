Public Class TerceroI
    Public Property cboidentificacion As String
    Public Property identificacion As String
    Public Property dv As String
    Public Property razon As String
    Public Property pnombre As String
    Public Property nombre As String
    Public Property papellido As String
    Public Property sapellido As String
    Public Property telefono As String
    Public Property telefono2 As String
    Public Property whatsapp As String
    Public Property email As String
    Public Property Combopais As Integer
    Public Property Combodepartamento As String
    Public Property Combociudad As String
    Public Property direccion As String
    Public Property usuario As Integer
    Public Property codigo As String
    Public Property dt As New DataTable

    Public Sub guardarTercero()
        TerceroDAL.GuardarTercero(Me)
    End Sub


End Class
