Public Class PedidoFarmacia
    Public Property codigoPedidoFarmaciaPunto As String
    Public Property codigoPedidoFarmacia As String
    Public Property codigoEP As Integer
    Public Property fechaPedido As DateTime
    Public Property codigoOrden As String
    Public Property codigoOrdenEnfermeria As String
    Public Property codigoOrdenFisioterapia As String
    Public Property tablaMedicamentos As DataTable
    Public Property tablaInfusiones As DataTable
    Public Property dtTablaMezclas As DataSet
    Public Property tablaNutricion As DataTable
    Public Property tablaInsumosEnfermeria As DataTable
    Public Property tablaInsumosFisioterapia As DataTable
    Public Property tablaImpregnaciones As DataTable
    Public Property nutVolumenesTotales As Double
    Public Property nutOsmolalidadTotal As Double
    Public Property nutRazon As Double


    Sub New()
        'se inicializan las tablas
        tablaMedicamentos = New DataTable
        tablaInfusiones = New DataTable
        tablaNutricion = New DataTable
        tablaInsumosEnfermeria = New DataTable
        tablaInsumosFisioterapia = New DataTable
        tablaImpregnaciones = New DataTable
        dtTablaMezclas = New DataSet

        'se inicializa la tabla de medicamentos 
        tablaMedicamentos.Columns.Add("Codigo_Interno", Type.GetType("System.Int32"))
        tablaMedicamentos.Columns.Add("Descripcion", Type.GetType("System.String"))
        tablaMedicamentos.Columns.Add("Concentracion", Type.GetType("System.Double"))
        tablaMedicamentos.Columns.Add("Cant_Dosis", Type.GetType("System.Double"))
        tablaMedicamentos.Columns.Add("Frecuencia", Type.GetType("System.Int32"))
        tablaMedicamentos.Columns.Add("Hora_Inicial", Type.GetType("System.Int32"))

        'se inicializa la tabla de infusiones y la de mezclas
        tablaInfusiones.Columns.Add("Consecutivo_Infusion", Type.GetType("System.Int32"))
        tablaInfusiones.Columns.Add("Codigo_Interno", Type.GetType("System.Int32"))
        tablaInfusiones.Columns.Add("Descripcion", Type.GetType("System.String"))
        tablaInfusiones.Columns.Add("Dosis", Type.GetType("System.Double"))
        tablaInfusiones.Columns.Add("Velocidad", Type.GetType("System.Double"))
        tablaInfusiones.Columns.Add("Producto_Disolvente", Type.GetType("System.Int32"))
        tablaInfusiones.Columns.Add("Descripcion_Disolvente", Type.GetType("System.String"))
        tablaInfusiones.Columns.Add("Cantidad_Disolvente", Type.GetType("System.String"))

        'se inicializa la tabla de impregaciones
        tablaImpregnaciones.Columns.Add("Codigo_Interno", Type.GetType("System.Int32"))
        tablaImpregnaciones.Columns.Add("Descripcion", Type.GetType("System.String"))
        tablaImpregnaciones.Columns.Add("Dosis", Type.GetType("System.Double"))
        tablaImpregnaciones.Columns.Add("Velocidad", Type.GetType("System.Double"))
        tablaImpregnaciones.Columns.Add("Producto_Disolvente", Type.GetType("System.Int32"))
        tablaImpregnaciones.Columns.Add("Descripcion_Disolvente", Type.GetType("System.String"))
        tablaImpregnaciones.Columns.Add("Cantidad_Disolvente", Type.GetType("System.String"))
        tablaImpregnaciones.Columns.Add("Hora_Inicial", Type.GetType("System.Int32"))

        'se inicializa la tabla de nutricion 
        tablaNutricion.Columns.Add("Codigo_Interno", Type.GetType("System.Int32"))
        tablaNutricion.Columns.Add("Descripcion", Type.GetType("System.String"))
        tablaNutricion.Columns.Add("Requerimiento", Type.GetType("System.Double"))
        tablaNutricion.Columns.Add("Volumenes", Type.GetType("System.Double"))

        'se inicializa la tabla de insumos enfermeria
        tablaInsumosEnfermeria.Columns.Add("Codigo_Interno", Type.GetType("System.Int32"))
        tablaInsumosEnfermeria.Columns.Add("Descripcion", Type.GetType("System.String"))
        tablaInsumosEnfermeria.Columns.Add("Cantidad", Type.GetType("System.Double"))

        'se inicializa la tabla de insumos Fisioterapia
        tablaInsumosFisioterapia.Columns.Add("Codigo_Interno", Type.GetType("System.Int32"))
        tablaInsumosFisioterapia.Columns.Add("Descripcion", Type.GetType("System.String"))
        tablaInsumosFisioterapia.Columns.Add("Cantidad", Type.GetType("System.Double"))
    End Sub
    Public Sub agregarTabla(ByVal pNombreTabla As String)
        If IsTablaExistente(pNombreTabla) = False Then
            dtTablaMezclas.Tables.Add(pNombreTabla)
            dtTablaMezclas.Tables(pNombreTabla).Columns.Add("Consecutivo_Infusion", Type.GetType("System.Int32"))
            dtTablaMezclas.Tables(pNombreTabla).Columns.Add("Codigo_Interno", Type.GetType("System.Int32"))
            dtTablaMezclas.Tables(pNombreTabla).Columns.Add("Descripcion", Type.GetType("System.String"))
            dtTablaMezclas.Tables(pNombreTabla).Columns.Add("Dosis", Type.GetType("System.Double"))
            dtTablaMezclas.Tables(pNombreTabla).Columns.Add("Cantidad", Type.GetType("System.Int32"))
        End If
    End Sub
    Public Sub quitarTablas(ByVal pNombreTabla As String)
        If IsTablaExistente(pNombreTabla) Then
            dtTablaMezclas.Tables.RemoveAt(pNombreTabla)
        End If
    End Sub
    Function IsTablaExistente(ByVal pNombreTabla As String)
        If dtTablaMezclas.Tables.Contains(pNombreTabla) Then
            Return True
        End If
        Return False
    End Function
    Public Function nombrarMezclaInfusion(ByVal idConsecutivo As String,
                                          ByVal codigoInterno As String) As String

        Dim nombreTabla As String = ""
        nombreTabla = idConsecutivo & codigoInterno
        Return nombreTabla
    End Function
End Class
