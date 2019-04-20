Public Class ExamenResultadoRR
    Inherits ExamenResultado
    Sub New()
        titulo = TitulosForm.TITULO_RESULT_RR
        area = ConstantesHC.NOMBRE_PDF_IMAGENOLOGIA_RR
        sqlCargarRegistro = ConsultasHC.RESULTADO_EXAMENES_CARGAR_RR
        sqlCargarRegistroDetalle = ConsultasHC.RESULTADO_EXAMENES_CARGAR_D_RR
        sqlAnularRegistro = ConsultasHC.RESULTADO_EXAMEN_ANULAR_RR
        sqlGuardarRegistro = ConsultasHC.RESULTADO_EXAMEN_GUARDAR_RR
        sqlGuardarRegistroHosp = "SP_RESULTADOS_HOSPITAL_EXAMENES_CREAR_RR"
        editado = 1
        moduloReporte = Constantes.REPORTE_AF
        nombrePDF = ConstantesHC.NOMBRE_PDF_IMAGENOLOGIA_RR
    End Sub
    Public Overrides Sub guardarRegistro()
        Dim objExamen As New ExamenResultadoBLL
        objExamen.guardarResultado(Me)
    End Sub

End Class
