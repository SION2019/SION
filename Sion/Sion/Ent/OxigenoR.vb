Public Class OxigenoR
    Inherits Oxigeno
    Sub New()
        modulo = Constantes.REPORTE_AM
        oxigenoCargarGuardados = ConsultasHC.OXIGENO_CARGAR_PRIMERA_VEZ_R
        nombreReporte = ConstantesHC.NOMBRE_PDF_OXIGENO_FISIO_R
        moduloReporte = Constantes.REPORTE_AM
        consultaFechaEgreso = ConsultasHC.OXIGENO_FECHA_EGRESO_R
    End Sub
    'Public Overrides Function getFechaEgreso() As Boolean
    '    Dim result As Boolean = False
    '    Dim params As New List(Of String)
    '    params.Add(nRegistro)
    '    Dim fila As DataRow
    '    fila = General.cargarItem(ConsultasHC.OXIGENO_FECHA_EGRESO_R, params)

    '    If Not IsDBNull(fila(0)) Then
    '        fechaEgreso = fila(0)
    '        result = True
    '    Else
    '        fechaEgreso = Funciones.Fecha(Constantes.FORMATO_FECHA_HORA_NUMERICA)
    '    End If
    '    Return result
    'End Function
End Class
