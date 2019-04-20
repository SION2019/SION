Public Class Ecocardiograma
    Inherits Eco
    Public Property sqlCargarParametro As String
    Public Property sqlCargarDetalle As String
    Public Property dtParametro As DataTable
    Public Property ventana As String
    Public Property conclusion As String

    Public Sub New()
        dtParametro = New DataTable
        dtParametro.Columns.Add("codigo", Type.GetType("System.String"))
        dtParametro.Columns.Add("Mediciòn", Type.GetType("System.String"))
        dtParametro.Columns.Add("Valores_Pacientes", Type.GetType("System.String")).DefaultValue = String.Empty
        dtParametro.Columns.Add("Valores_Normales", Type.GetType("System.String"))
        titulo = TitulosForm.TITULO_ECO
        sqlCargarParametro = Consultas.CARGAR_PARAMETROS_ECO
        sqlGuardarRegistro = Consultas.ECOCARDIOGRAMA_CREAR
        sqlAnularRegistro = Consultas.ANULAR_ECOCARDIOGRAMA
        sqlCargarDetalle = Consultas.CARGAR_ECOCARDIOGRAMA_D
        formula = "{VISTA_REPORTE_ECO.Codigo_Eco} = $" &
          " AND    {VISTA_PACIENTES.Id_empresa}=" & idEmpresa &
          " AND    {VISTA_REPORTE_ECO.Modulo}= 0"
        area = ConstantesHC.NOMBRE_PDF_ECO
    End Sub

    Public Sub cargarParametros()
        Dim params As New List(Of String)
        params.Add(codigoArea)
        General.llenarTabla(sqlCargarParametro, params, dtParametro)
    End Sub


End Class
