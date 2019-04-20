Public Class OrdenMedicaRR
    Inherits OrdenMedica
    Sub New()
        objetoNutricion = New OrdenMedicaNutricionRR
        nombreReporte = ConstantesHC.NOMBRE_PDF_ORDEN_MEDICARR
        moduloReporte = Constantes.REPORTE_AF
        ordenListar = Consultas.ORDEN_MEDICA_LISTAR_RR
        ordenCargar = Consultas.ORDEN_MEDICA_CARGARRR
        estanciaCargar = Consultas.ORDEN_MEDICA_ESTANCIA_CARGARRR
        indicacionCargar = Consultas.ORDEN_MEDICA_INDICACION_CARGARRR
        comidaCargar = Consultas.ORDEN_MEDICA_COMIDA_CARGARRR
        oxigenoCargar = Consultas.ORDEN_MEDICA_OXIGENO_CARGARRR
        oxigenoAutoCargar = Consultas.ORDEN_MEDICA_OXIGENO_AUTO_CARGARRR
        paraclinicoCargar = Consultas.ORDEN_MEDICA_PARACLINICO_CARGARRR
        hemoderivadoCargar = Consultas.ORDEN_MEDICA_HEMODERIVADO_CARGARRR
        glucometriaCargar = Consultas.ORDEN_MEDICA_GLUCOMETRIA_CARGARRR
        procedimientoCargar = Consultas.ORDEN_MEDICA_PROCEDIMIENTO_CARGARRR
        medicamentoCargar = Consultas.ORDEN_MEDICA_MEDICAMENTO_CARGARRR
        impregnacionCargar = Consultas.ORDEN_MEDICA_IMPREGNACION_CARGARRR
        infusionCargar = Consultas.ORDEN_MEDICA_INFUSION_CARGARRR
        mezclaCargar = Consultas.ORDEN_MEDICA_MEZCLA_CARGARRR
        medicamentoAutoCargar = Consultas.ORDEN_MEDICA_MEDICAMENTO_AUTO_CARGARRR
        procedimientoAutoCargar = Consultas.ORDEN_MEDICA_PROCEDIMIENTO_AUTO_CARGARRR
        medicamentoAuditoriaCargar = Consultas.ORDEN_MEDICA_MEDICAMENTO_AUDITORIA_CARGARRR
    End Sub
    Public Overrides Sub guardarSabanaAutomatica()
        Dim objSabanaAplicacion As SabanaAplicacionMed
        objSabanaAplicacion = GeneralHC.fabricaHC.crear(ConstantesHC.IDENTIFICADOR_SABANA_APLICACION & Constantes.CODIGO_MENU_AUDF)
        objSabanaAplicacion.usuario = usuario
        objSabanaAplicacion.nRegistro = registro
        objSabanaAplicacion.guardarReporte2doPlano()
    End Sub
    Public Overrides Function verificarHoraAplicacion(dt As DataTable) As Boolean
        For i = 0 To dt.Rows.Count - 2
            If dt.Rows(i).Item("Inicio").ToString = Constantes.ITEM_dgv_SELECCIONE Then
                Return True
            End If
        Next
        Return False
    End Function

End Class
