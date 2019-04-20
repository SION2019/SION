Public Class ConfiguracionProcedimientoCups
    Public Property CodigoTipo As String
    Public Property dtProcedimiento As DataTable
    Public Property codigoPunto As Integer
    Public Property sqlCargarProcedimientos As String
    Public Property sqlcargarProcedimientoEstancia As String
    Public Property sqlcargarProcedimientoHemoderivado As String
    Public Property sqlcargarProcedimientoParaclinico As String
    Public Property sqlGrupoProcedimientos As String
    Public Property sqlAreaProcedimientoEstancia As String
    Public Property sqlGrupoProcedimientoHemoderivado As String
    Public Property sqlTipoExamenProcedimientoParaclinico As String
    Public Property sqlCrearProcedimientos As String
    Public Property sqlCrearProcedimientoEstancia As String
    Public Property sqlCrearProcedimientoHemoderivado As String
    Public Property sqlCrearProcedimientoParaclinico As String
    Public Property codigoConfig As Integer
    Public Property usuario As Integer
    Public Sub New()
        codigoPunto = SesionActual.codigoEP
        dtProcedimiento = New DataTable
        sqlCrearProcedimientos = Consultas.PROCEDIMIENTO_CONFIGURACION_CREAR
        sqlCrearProcedimientoEstancia = Consultas.PROCEDIMIENTO_CONFIGURACION_ESTANCIA_CREAR
        sqlCrearProcedimientoHemoderivado = Consultas.PROCEDIMIENTO_CONFIGURACION_HEMODERIVADO_CREAR
        sqlCrearProcedimientoParaclinico = Consultas.PROCEDIMIENTO_CONFIGURACION_PARACLINICOS_CREAR

        sqlCargarProcedimientos = Consultas.PROCEDIMIENTOS_CUPS_LISTAR
        sqlcargarProcedimientoEstancia = Consultas.PROCEDIMIENTOS_ESTANCIA_LISTAR
        sqlcargarProcedimientoHemoderivado = Consultas.PROCEDIMIENTOS_TRANSFUSION_LISTAR
        sqlcargarProcedimientoParaclinico = Consultas.PROCEDIMIENTOS_PARACLINICO_LISTAR

        sqlGrupoProcedimientos = Consultas.GRUPO_PROCEDIMIENTOS_LISTAR
        sqlAreaProcedimientoEstancia = Consultas.LISTA_AREA_SERVICIO_CONF
        sqlGrupoProcedimientoHemoderivado = Consultas.LISTA_GRUPO_SANGUINEO
        sqlTipoExamenProcedimientoParaclinico = Consultas.PROCEDIMIENTOS_CUPS_PARACLINICOS_LISTAR

        dtProcedimiento.Columns.Add("Codigo", Type.GetType("System.String"))
        dtProcedimiento.Columns.Add("Descripcion", Type.GetType("System.String"))
        dtProcedimiento.Columns.Add("Seleccionar", Type.GetType("System.Boolean"))
    End Sub
    Public Sub llenarTabla()
        Dim params As New List(Of String)
        params.Add(CodigoTipo)
        General.llenarTabla(consulta(Constantes.CODIGO_CONFIG_PROCEDIMIENTO), params, dtProcedimiento)
    End Sub

    Public Function consulta() As List(Of String)
        Dim params As New List(Of String)
        Select Case codigoConfig
            Case Constantes.CODIGO_CONFIG_PROCEDIMIENTO '------ Codigo para indentificar las consulta de configuarcion de procedimiento
                params.Add(sqlCargarProcedimientos)
                params.Add(sqlGrupoProcedimientos)
                params.Add(sqlCrearProcedimientos)
            Case Constantes.CODIGO_CONFIG_ESTANCIA '---------Codigo para indentificar las consulta de configuarcion de procedimiento de estancia
                params.Add(sqlcargarProcedimientoEstancia & codigoPunto & ConstantesHC.NOMBRE_PDF_SEPARADOR3)
                params.Add(sqlAreaProcedimientoEstancia)
                params.Add(sqlCrearProcedimientoEstancia)
            Case Constantes.GRUPO_TRANSFUSION  '---------Codigo para indentificar las consulta de configuarcion de procedimiento de Hemoderivado
                params.Add(sqlcargarProcedimientoHemoderivado)
                params.Add(sqlGrupoProcedimientoHemoderivado)
                params.Add(sqlCrearProcedimientoHemoderivado)
            Case Constantes.GRUPO_PARACLINICO  '---------Codigo para indentificar las consulta de configuarcion de procedimiento de Paraclinicos
                params.Add(sqlcargarProcedimientoParaclinico)
                params.Add(sqlTipoExamenProcedimientoParaclinico)
                params.Add(sqlCrearProcedimientoParaclinico)
        End Select
        Return params
    End Function


End Class
