Public Class ComiteCTCRR
    Inherits ComiteCTC
    Public Sub New()
        sqlCargarDiagnostico = Consultas.EVO_DIAG_ULTIMO_REG_RR
        titulo = "COMITE TECNICO CIENTIFICO (C.T.C): AUDITORÍA FACTURACIÓN"
        sqlBuscarRegistro = Consultas.BUSQUEDA_BUSCAR_CTC_RR
        sqlCargarRegistro = Consultas.CARGAR_CTC_RR
        sqlCargarPaciente = Consultas.CARGAR_PACIENTE_CTC_RR
        sqlBuscarPaciente = Consultas.BUSQUEDA_BUSCAR_PACIENTE_CTC_RR
        sqlGuardarRegistro = Consultas.CREAR_CTC_RR
        sqlAnularRegistro = Consultas.ANULAR_CTC_RR
        area = ConstantesHC.NOMBRE_PDF_CTC_RR
        moduloReporte = Constantes.REPORTE_AF
        editado = 1
    End Sub
    Public Overrides Sub guardarRegistro()
        If codigo = String.Empty Then
            codigo = -1
        End If
        ComiteTecnicoCientificoBLL.guardarCTC(Me)
    End Sub
End Class
