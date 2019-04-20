Public Class ComiteCTCR
    Inherits ComiteCTC
    Public Sub New()
        titulo = "COMITE TECNICO CIENTIFICO (C.T.C): AUDITORÍA MÉDICA"
        sqlCargarDiagnostico = Consultas.EVO_DIAG_ULTIMO_REG_R
        sqlBuscarRegistro = Consultas.BUSQUEDA_BUSCAR_CTC_R
        sqlCargarRegistro = Consultas.CARGAR_CTC_R
        sqlCargarPaciente = Consultas.CARGAR_PACIENTE_CTC_R
        sqlBuscarPaciente = Consultas.BUSQUEDA_BUSCAR_PACIENTE_CTC_R
        sqlGuardarRegistro = Consultas.CREAR_CTC_R
        sqlAnularRegistro = Consultas.ANULAR_CTC_R
        area = ConstantesHC.NOMBRE_PDF_CTC_R
        moduloReporte = Constantes.REPORTE_AM
        editado = 1
    End Sub
    Public Overrides Sub guardarRegistro()
        If codigo = String.Empty Then
            codigo = -1
        End If
        ComiteTecnicoCientificoBLL.guardarCTC(Me)
    End Sub
End Class
