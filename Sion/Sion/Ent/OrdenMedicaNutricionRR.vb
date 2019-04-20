Public Class OrdenMedicaNutricionRR
    Inherits OrdenMedicaNutricion
    Public Overrides Sub cargarNutricion()
        Dim params As New List(Of String)
        params.Add(codigoOrden)
        Dim fila As DataRow
        fila = General.cargarItem(Consultas.ORDEN_MEDICA_NUTRICION_CARGARRR, params)
        If Not IsNothing(fila) Then
            peso = fila.Item("peso")
            horaInicial = fila.Item("Hora_inicial")
            TasaHidrica = fila.Item("Tasa_Hidrica")
            FactorAjustes = fila.Item("Fator_ajustes")
            alimentacion = fila.Item("Alimentacion")
            medicacion = fila.Item("Medicacion")
            otros = fila.Item("otros")
            tht = fila.Item("THT")
            volumen = fila.Item("Volumen")
            razon = fila.Item("razon")
            General.llenarTabla(Consultas.ORDEN_MEDICA_NUTRICION_DETALLE_CARGARRR, params, dtNutricion)
        Else
            codigoOrden = -1
        End If
    End Sub
    Public Overrides Sub llenarNutricionParaEdicion()
        Dim params As New List(Of String)
        params.Add(codigoOrden)
        General.llenarTabla(ConsultasHC.NUTRICION_POR_DEFECTOS_EDICION_RR, params, dtNutricion)
    End Sub
    Public Overrides Function ultimoPeso(ByVal numeroRegistro As Integer,
                                         ByVal fecha As Date) As Double
        Dim params As New List(Of String)
        params.Add(numeroRegistro)
        params.Add(Format(fecha, Constantes.FORMATO_FECHA_GEN))

        Dim fila As DataRow
        fila = General.cargarItem(ConsultasHC.NUTRICION_ULTIMO_PESO_RR, params)
        Return fila(0)

    End Function
End Class
