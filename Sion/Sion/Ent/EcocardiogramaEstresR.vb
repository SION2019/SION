Public Class EcocardiogramaEstresR
    Inherits EcocardiogramaEstres
    Public Sub New()
        titulo = TitulosForm.TITULO_ECO_ESTRES_R
        sqlGuardarRegistro = Consultas.ECOSTRESS_CREAR_R
        sqlAnularRegistro = Consultas.ECOSTRESS_ANULAR_R
        editado = Constantes.EDITADO
        formula = " {VISTA_REPORTE_ECOSTRES.Codigo_Eco} = $" &
                  " AND {VISTA_PACIENTES.Id_empresa}=" & idEmpresa &
                  " AND {VISTA_REPORTE_ECOSTRES.Modulo}= 1"
        area = ConstantesHC.NOMBRE_PDF_ECOSTRES_R
    End Sub

End Class
