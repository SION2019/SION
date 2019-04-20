Public Class InformeQxRR
    Inherits InformeQx

    Public Sub New()
        sqlCargarRegistro = Consultas.CARGAR_QX_RR
        sqlBuscarRegistro = Consultas.BUSQUEDA_BUSCAR_QX_RR
        sqlBuscarPaciente = Consultas.BUSQUEDA_BUSCAR_PACIENTE_QX_RR
        sqlCargarPaciente = Consultas.CARGAR_PACIENTE_QX_RR
        sqlcargarDiagUltimaEvo = Consultas.EVO_DIAG_ULTIMO_REG_RR
        sqlGuardarRegistro = Consultas.QX_CREAR_RR
        titulo = TitulosForm.TITULO_INFORME_QX_RR
        area = ConstantesHC.NOMBRE_PDF_INFORMEQX_RR
        sqlAnularRegistro = Consultas.QX_ANULAR_RR
        nombreReporte = area
        moduloreporte = Constantes.REPORTE_AF
        sqlCargarInsumoConfig = Consultas.INSUMO_CONF_AUDITORIA_CARGAR_RR
        sqlCargarParaclinicoConfig = Consultas.PARACLINICO_CONF_AUDITORIA_CARGAR_RR
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
