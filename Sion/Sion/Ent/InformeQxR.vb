Public Class InformeQxR
    Inherits InformeQx

    Public Sub New()
        sqlCargarRegistro = Consultas.CARGAR_QX_R
        sqlBuscarRegistro = Consultas.BUSQUEDA_BUSCAR_QX_R
        sqlBuscarPaciente = Consultas.BUSQUEDA_BUSCAR_PACIENTE_QX_R
        sqlCargarPaciente = Consultas.CARGAR_PACIENTE_QX_R
        sqlcargarDiagUltimaEvo = Consultas.EVO_DIAG_ULTIMO_REG_R
        sqlGuardarRegistro = Consultas.QX_CREAR_R
        titulo = TitulosForm.TITULO_INFORME_QX_R
        area = ConstantesHC.NOMBRE_PDF_INFORMEQX_R
        sqlAnularRegistro = Consultas.QX_ANULAR_R
        nombreReporte = area
        moduloreporte = Constantes.REPORTE_AM
        sqlCargarInsumoConfig = Consultas.INSUMO_CONF_AUDITORIA_CARGAR_R
        sqlCargarParaclinicoConfig = Consultas.PARACLINICO_CONF_AUDITORIA_CARGAR_R
    End Sub
    Public Overrides Sub guardarRegistro()
        If String.IsNullOrEmpty(codigo) Then
            codigo = -1
        Else
            editado = 1
        End If
        InformeQuirurgicoBLL.guardarInformeQx(Me)
    End Sub
End Class
