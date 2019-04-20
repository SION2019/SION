Public Class HemodialisisRR
    Inherits Hemodialisis
    Public Property registro As Integer
    Public Sub New()
        area = ConstantesHC.NOMBRE_PDF_HEMODIALISIS_RR
        titulo = "HEMODIALISIS: AUDITORÍA FACTURACIÓN"
        sqlBuscarPaciente = Consultas.BUSQUEDA_BUSCAR_PACIENTE_HEMO_RR
        sqlCargarPaciente = Consultas.CARGAR_PACIENTE_HEMO_RR
        sqlCargarRegistro = Consultas.CARGAR_HEMO_RR
        sqlBuscarRegistro = Consultas.BUSQUEDA_BUSCAR_HEMODIALISI_RR
        sqlGuardarRegistro = Consultas.CREAR_HEMO_RR
        sqlAnularRegistro = Consultas.HEMOD_ANULAR_RR
        moduloreporte = Constantes.REPORTE_AF
        nombreReporte = ConstantesHC.NOMBRE_PDF_TRANSFUSION_RR
        sqlCargarInsumoConfig = Consultas.INSUMO_CONF_AUDITORIA_CARGAR_RR
        sqlCargarParaclinicoConfig = Consultas.PARACLINICO_CONF_AUDITORIA_CARGAR_RR
    End Sub


    Public Overrides Sub guardarRegistro()
        If String.IsNullOrEmpty(codigo) Then
            codigo = -1
        Else
            editado = 1
        End If
        HemodialisiBLL.guardarHemodialisi(Me)
    End Sub
End Class
