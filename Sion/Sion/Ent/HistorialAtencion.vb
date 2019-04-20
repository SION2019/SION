Public Class HistorialAtencion
    Private fechaInic As Date
    Private fechaFin As Date
    Private busqueda As String
    Private comboContrato As String
    Private comboContacto As String
    Sub New()

    End Sub

    Public Property setGetFechaIni As Date
        Get
            Return fechaInic
        End Get
        Set(value As Date)
            value = Format(CDate(value), "yyyy-MM-dd")
            fechaInic = value
        End Set
    End Property

    Public Property setGetFechaFin As Date
        Get
            Return fechaFin
        End Get
        Set(value As Date)
            value = Format(CDate(value), "yyyy-MM-dd")
            fechaFin = value
        End Set
    End Property
    Public Property setGetBusqueda As String
        Get
            Return busqueda
        End Get
        Set(value As String)
            value = If(value = String.Empty, Nothing, value)
            busqueda = value
        End Set
    End Property
    Public Property setGetContrato As String
        Get
            Return comboContrato
        End Get
        Set(value As String)
            value = If(value = String.Empty, Nothing, value)
            comboContrato = value
        End Set
    End Property
    Public Property setGetContacto As String
        Get
            Return comboContacto
        End Get
        Set(value As String)
            value = If(value = String.Empty, Nothing, value)
            comboContacto = value
        End Set
    End Property
    Public Sub seleccionarNodos(nodo As TreeView)
        For i = 1 To nodo.Nodes.Count - 1
            nodo.Nodes(i).Checked = True
        Next
    End Sub

    Public Sub ocultarColumnas(nodo As TreeView, grilla As DataGridView)
        Try
            For i = 0 To nodo.Nodes.Count - 1
                grilla.Columns(nodo.Nodes(i).Text.ToString).Visible = nodo.Nodes(i).Checked
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
