Public Class ListaEmpleado
    Public Property dtEmpleado As DataTable
    Public Property sqlCargarLista As String
    Public Property codigoEp As Integer
    Public Property codigoPunto As Integer
    Public Sub New()
        dtEmpleado = New DataTable
        dtEmpleado.Columns.Add("Id_empleado", Type.GetType("System.Int32"))
        dtEmpleado.Columns.Add("Cedula", Type.GetType("System.String"))
        dtEmpleado.Columns.Add("Empleado", Type.GetType("System.String"))
        dtEmpleado.Columns.Add("Cargo", Type.GetType("System.String"))
        dtEmpleado.Columns.Add("Empresa", Type.GetType("System.String"))
        dtEmpleado.Columns.Add("Punto_Servicio", Type.GetType("System.String"))
        dtEmpleado.Columns.Add("Eps", Type.GetType("System.String"))
        dtEmpleado.Columns.Add("Arp", Type.GetType("System.String"))
        dtEmpleado.Columns.Add("Caja", Type.GetType("System.String"))
        dtEmpleado.Columns.Add("Pension", Type.GetType("System.String"))
        dtEmpleado.Columns.Add("Contrato", Type.GetType("System.String"))
        dtEmpleado.Columns.Add("Salario_Contrato", Type.GetType("System.String"))
        sqlCargarLista = Consultas.EMPLEADO_LISTAR
    End Sub

    Public Sub cargarDatos(fechaInicio As Date, fechaFin As Date, inactivo As Int16, cadena As String)
        Dim params As New List(Of String)
        params.Add(fechaInicio)
        params.Add(fechaFin)
        params.Add(codigoEp)
        params.Add(codigoPunto)
        params.Add(inactivo)
        params.Add(cadena)
        General.llenarTabla(sqlCargarLista, params, dtEmpleado)
    End Sub

End Class
