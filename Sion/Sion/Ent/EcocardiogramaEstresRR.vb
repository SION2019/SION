Public Class EcocardiogramaEstresRR
    Inherits EcocardiogramaEstres
    Public Sub New()
        titulo = TitulosForm.TITULO_ECO_ESTRES_RR
        sqlGuardarRegistro = Consultas.ECOSTRESS_CREAR_RR
        sqlAnularRegistro = Consultas.ECOSTRESS_ANULAR_RR
        editado = Constantes.EDITADO
        formula = " {VISTA_REPORTE_ECOSTRES.Codigo_Eco} = $" &
                  " AND {VISTA_PACIENTES.Id_empresa}=" & idEmpresa &
                  " AND {VISTA_REPORTE_ECOSTRES.Modulo}= 2"
        area = ConstantesHC.NOMBRE_PDF_ECOSTRES_RR
    End Sub
End Class
