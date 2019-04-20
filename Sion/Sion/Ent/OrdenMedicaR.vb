Public Class OrdenMedicaR
    Inherits OrdenMedica
    Sub New()
        objetoNutricion = New OrdenMedicaNutricionR
        nombreReporte = ConstantesHC.NOMBRE_PDF_ORDEN_MEDICAR
        moduloReporte = Constantes.REPORTE_AM
        ordenListar = Consultas.ORDEN_MEDICA_LISTAR_R
        ordenCargar = Consultas.ORDEN_MEDICA_CARGARR
        estanciaCargar = Consultas.ORDEN_MEDICA_ESTANCIA_CARGARR
        indicacionCargar = Consultas.ORDEN_MEDICA_INDICACION_CARGARR
        comidaCargar = Consultas.ORDEN_MEDICA_COMIDA_CARGARR
        oxigenoCargar = Consultas.ORDEN_MEDICA_OXIGENO_CARGARR
        oxigenoAutoCargar = Consultas.ORDEN_MEDICA_OXIGENO_AUTO_CARGARR
        paraclinicoCargar = Consultas.ORDEN_MEDICA_PARACLINICO_CARGARR
        hemoderivadoCargar = Consultas.ORDEN_MEDICA_HEMODERIVADO_CARGARR
        glucometriaCargar = Consultas.ORDEN_MEDICA_GLUCOMETRIA_CARGARR
        procedimientoCargar = Consultas.ORDEN_MEDICA_PROCEDIMIENTO_CARGARR
        medicamentoCargar = Consultas.ORDEN_MEDICA_MEDICAMENTO_CARGARR
        impregnacionCargar = Consultas.ORDEN_MEDICA_IMPREGNACION_CARGARR
        infusionCargar = Consultas.ORDEN_MEDICA_INFUSION_CARGARR
        mezclaCargar = Consultas.ORDEN_MEDICA_MEZCLA_CARGARR
        medicamentoAutoCargar = Consultas.ORDEN_MEDICA_MEDICAMENTO_AUTO_CARGARR
        procedimientoAutoCargar = Consultas.ORDEN_MEDICA_PROCEDIMIENTO_AUTO_CARGARR
        medicamentoAuditoriaCargar = Consultas.ORDEN_MEDICA_MEDICAMENTO_AUDITORIA_CARGARR
    End Sub
    Public Overrides Sub guardarSabanaAutomatica()

        Dim objSabanaAplicacion As SabanaAplicacionMed
        objSabanaAplicacion = GeneralHC.fabricaHC.crear(ConstantesHC.IDENTIFICADOR_SABANA_APLICACION & Constantes.CODIGO_MENU_AUDM)
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
