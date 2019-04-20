Public Class RemisionR
    Inherits Remision
    Public Sub New()
        nombreReporte = ConstantesHC.NOMBRE_PDF_REMISIONR
        moduloReporte = Constantes.REPORTE_AM
        remisionListar = Consultas.REMISIONES_LISTAR_R
        remisionCargar = Consultas.REMISION_CARGARR
        consultaSolicitudVerificar = Consultas.REMISION_SOLICITUD_VERIFICARR
    End Sub

    Public Overrides Sub guardarRemision()
        Try
            If String.IsNullOrEmpty(codigoOrden) Then
                codigoOrden = Constantes.VALOR_PREDETERMINADO
            End If
            HistoriaClinicaBLL.guardarRemisionR(Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub anularremision()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoremision)
        params.Add(codigoEP)
        General.ejecutarSQL(Consultas.ANULAR_REMISIONR, params)
        consultaVerificar = Consultas.REMISION_VERIFICARR
    End Sub
End Class
