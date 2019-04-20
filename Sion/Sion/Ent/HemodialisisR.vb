Public Class HemodialisisR
    Inherits Hemodialisis
    Public Property registro As Integer
    Public Sub New()
        area = ConstantesHC.NOMBRE_PDF_HEMODIALISIS_R
        titulo = "HEMODIALISIS: AUDITORÍA MÉDICA"
        sqlBuscarPaciente = Consultas.BUSQUEDA_BUSCAR_PACIENTE_HEMO_R
        sqlCargarPaciente = Consultas.CARGAR_PACIENTE_HEMO_R
        sqlCargarRegistro = Consultas.CARGAR_HEMO_R
        sqlBuscarRegistro = Consultas.BUSQUEDA_BUSCAR_HEMODIALISI_R
        sqlGuardarRegistro = Consultas.CREAR_HEMO_R
        sqlAnularRegistro = Consultas.HEMOD_ANULAR_R
        moduloreporte = Constantes.REPORTE_AM
        nombreReporte = ConstantesHC.NOMBRE_PDF_TRANSFUSION_R
        sqlCargarInsumoConfig = Consultas.INSUMO_CONF_AUDITORIA_CARGAR_R
        sqlCargarParaclinicoConfig = Consultas.PARACLINICO_CONF_AUDITORIA_CARGAR_R
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
