Public Class OrdenMedicaNutricion
    Public Property codigoOrden As Integer
    Public Property peso As Double
    Public Property horaInicial As Integer
    Public Property TasaHidrica As Double
    Public Property FactorAjustes As Double
    Public Property alimentacion As Double
    Public Property medicacion As Double
    Public Property otros As Double
    Public Property tht As Double
    Public Property volumen As Double
    Public Property razon As Double
    Public Property codigoAgua As String
    Public Property cantidadAgua As Double
    Public Property dtNutricion As DataTable
    Sub New()
        dtNutricion = New DataTable
        dtNutricion.Columns.Add("Codigo_Interno", Type.GetType("System.String"))
        dtNutricion.Columns.Add("Descripcion", Type.GetType("System.String"))
        dtNutricion.Columns.Add("Requerimiento", Type.GetType("System.Double"))
        dtNutricion.Columns.Add("Osmolalidad", Type.GetType("System.Double"))
        dtNutricion.Columns.Add("Volumenes", Type.GetType("System.Double"))
        dtNutricion.Columns.Add("Divisor", Type.GetType("System.String"))
        dtNutricion.Columns.Add("Formula", Type.GetType("System.String"))
        dtNutricion.Columns.Add("Tipo", Type.GetType("System.String"))
    End Sub
    Public Function calcularGramos(ByVal indiceFila As Integer) As Double
        Return (dtNutricion.Rows(indiceFila).Item("Volumenes") * dtNutricion.Rows(indiceFila).Item("Divisor")) / ConstantesHC.FACTOR_DIVISION_GRAMOS_NUTRICION
    End Function
    Public Function calcularAminoacidos(ByVal indiceFila As Integer) As Double
        Return dtNutricion.Rows(indiceFila).Item("Volumenes") * ConstantesHC.FACTOR_MULTIPLICACION_AMINOACIDOS
    End Function
    Public Function calcularlipidos(ByVal indiceFila As Integer) As Double
        Return dtNutricion.Rows(indiceFila).Item("Volumenes") * ConstantesHC.FACTOR_MULTIPLICACION_lIPIDOS
    End Function
    Public Overridable Sub cargarNutricion()
        Dim params As New List(Of String)
        params.Add(codigoOrden)
        Dim fila As DataRow
        fila = General.cargarItem(Consultas.ORDEN_MEDICA_NUTRICION_CARGAR, params)
        If Not IsNothing(fila) Then
            peso = fila.Item("peso")
            horaInicial = fila.Item("Hora_inicial")
            TasaHidrica = fila.Item("Tasa_Hidrica")
            FactorAjustes = fila.Item("Fator_ajustes")
            alimentacion = fila.Item("Alimentacion")
            medicacion = fila.Item("Medicacion")
            otros = fila.Item("otros")
            tht = fila.Item("THT")
            volumen = fila.Item("Volumen")
            razon = fila.Item("razon")
            General.llenarTabla(Consultas.ORDEN_MEDICA_NUTRICION_DETALLE_CARGAR, params, dtNutricion)
        Else
            codigoOrden = -1
        End If
    End Sub
    Public Overridable Sub llenarNutricionParaEdicion()
        Dim params As New List(Of String)
        params.Add(codigoOrden)
        General.llenarTabla(ConsultasHC.NUTRICION_POR_DEFECTOS_EDICION, params, dtNutricion)
    End Sub
    Public Overridable Function ultimoPeso(ByVal numeroRegistro As Integer,
                                           ByVal fecha As Date) As Double
        Dim params As New List(Of String)
        params.Add(numeroRegistro)
        params.Add(Format(fecha, Constantes.FORMATO_FECHA_GEN))

        Dim fila As DataRow
        fila = General.cargarItem(ConsultasHC.NUTRICION_ULTIMO_PESO, params)
        Return fila(0)

    End Function
End Class
