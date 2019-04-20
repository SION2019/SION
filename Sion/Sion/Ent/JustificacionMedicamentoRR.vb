Public Class JustificacionMedicamentoRR
    Inherits JustificacionMedicamento

    Sub New()
        diagnosticoSolicitado = Consultas.JUSTIFICACION_CARGAR_DIAG_RR
        antibioticoTipo = Consultas.JUSTIFICACION_ANT_DETALLE_CARGAR_RR
        anular = Consultas.ANULAR_JUSTIFICACION_ANTIBIOTICO_RR
        medicamento = Consultas.JUSTIFICACION_ANT_MEDICAMENTO_CARGAR_RR
        Antibioticos = Consultas.JUSTIFICACION_ANTIBIOTICO_CARGAR_RR
        diagnosticoEvo = Consultas.DIAGNOSTICO_EVOLUCION_NOPOS_CARGAR_RR
        diagnosticoEvo2 = Consultas.DIAGNOSTICO_EVOLUCION_NOPOS_CARGAR2_RR
        guardarAntibiotico = Consultas.JUSTIFICACION_GUARDAR_RR
        actualizarAntibiotico = Consultas.JUSTIFICACION_ACTUALIZAR_RR
        nombrePDF = ConstantesHC.NOMBRE_PDF_JUSTIFICACION_RR
        consultaBuscar = Consultas.JUSTIFICACION_ANT_REGISTRO_BUSCAR_RR
        cargarBusqueda = Consultas.JUSTIFICACION_ANT_BUSCAR_RR
        titulo = "JUSTIFICACIÓN DE ANTIBIOTICOS: AUDITORÍA FACTURACIÓN"
        nombreReporte = ConstantesHC.NOMBRE_PDF_TRANSFUSION_RR
        moduloreporte = Constantes.REPORTE_AF
    End Sub

    Public Overrides Sub guardarJustificacion()

        If String.IsNullOrEmpty(codigo) Then
            codigo = Constantes.VALOR_PREDETERMINADO
        End If
        AntibioticoJustificacionBLL.guardarAntibiotico(Me)
    End Sub
    Public Overrides Sub anulaJustificacion()
        Dim lista As New List(Of String)
        lista.Add(codigo)
        lista.Add(SesionActual.idUsuario)
        General.ejecutarSQL(anular, lista)
    End Sub
End Class
