Public Class EcocardiogramaEstres
    Inherits Eco
    Public Sub New()
        titulo = TitulosForm.TITULO_ECO_ESTRES
        sqlGuardarRegistro = Consultas.ECOSTRESS_CREAR
        sqlAnularRegistro = Consultas.ECOSTRESS_ANULAR
        formula = " {VISTA_REPORTE_ECOSTRES.Codigo_Eco} = $" &
                  " AND {VISTA_PACIENTES.Id_empresa}=" & idEmpresa &
                  " AND {VISTA_REPORTE_ECOSTRES.Modulo}= 0"
        area = ConstantesHC.NOMBRE_PDF_ECOSTRES
    End Sub



End Class
