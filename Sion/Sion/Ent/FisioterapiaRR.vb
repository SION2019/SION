Public Class FisioterapiaRR
    Inherits Fisioterapia

    Sub New()
        notaReporte = ConstantesHC.NOMBRE_PDF_NOTA_FISIO_RR
        moduloReporte = Constantes.REPORTE_AF
        terapiaCarga = ConsultasHC.FISIOTERAPIA_FISICA_Y_RESPIRATORIA_LISTAR_RR
        insumoCarga = ConsultasHC.FISIOTERAPIA_INSUMO_LISTAR_RR
        notaCarga = ConsultasHC.FISIOTERAPIA_NOTA_LISTAR_RR
        oxigenoCarga = ConsultasHC.OXIGENO_LISTAR_RR
        nebulizacionCarga = ConsultasHC.NEBULIZACION_LISTAR_RR
        objTerapiaFyR = New TerapiaFisicaYRespiratoriaRR
        objInsumosFisio = New InsumoFisioterapiaRR
        objNotaFisio = New NotaFisioterapiaRR
        objOxigeno = New OxigenoRR
        objnebulizacion = New NebulizacionRR
    End Sub
    Public Overrides Sub guardarDetalleTerapia()
        Try
            HistoriaClinicaDAL.guardarDetalleTerapiaRR(objTerapiaFyR)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub guardarOrdenInsumoFisio()
        Try
            If String.IsNullOrEmpty(objInsumosFisio.codigoOrden) Then
                objInsumosFisio.codigoOrden = Constantes.VALOR_PREDETERMINADO
                HistoriaClinicaDAL.guardarDetalleInsumosFisioRR(objInsumosFisio)
            Else
                objInsumosFisio.editado = ConstantesHC.VALOR_EDITADO
                HistoriaClinicaDAL.actualizarDetalleInsumosFisioRR(objInsumosFisio)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub guardarNotaFisio()
        Try
            If String.IsNullOrEmpty(objNotaFisio.codigoNota) Then
                objNotaFisio.codigoNota = Constantes.VALOR_PREDETERMINADO
                HistoriaClinicaDAL.guardarNotaFisioRR(objNotaFisio)
            Else
                objNotaFisio.editado = ConstantesHC.VALOR_EDITADO
                HistoriaClinicaDAL.actualizarNotaFisioRR(objNotaFisio)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub guardarOxigeno()
        Try
            HistoriaClinicaDAL.guardarFisioterapiaOxigenoRR(objOxigeno)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub guardarNebulizacion()
        Try
            HistoriaClinicaDAL.guardarFisioterapiaNebulizacionRR(objnebulizacion)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
