Public Class FisioterapiaR
    Inherits Fisioterapia

    Sub New()
        notaReporte = ConstantesHC.NOMBRE_PDF_NOTA_FISIO_R
        moduloReporte = Constantes.REPORTE_AM
        terapiaCarga = ConsultasHC.FISIOTERAPIA_FISICA_Y_RESPIRATORIA_LISTAR_R
        insumoCarga = ConsultasHC.FISIOTERAPIA_INSUMO_LISTAR_R
        notaCarga = ConsultasHC.FISIOTERAPIA_NOTA_LISTAR_R
        oxigenoCarga = ConsultasHC.OXIGENO_LISTAR_R
        nebulizacionCarga = ConsultasHC.NEBULIZACION_LISTAR_R
        objTerapiaFyR = New TerapiaFisicaYRespiratoriaR
        objInsumosFisio = New InsumoFisioterapiaR
        objNotaFisio = New NotaFisioterapiaR
        objOxigeno = New OxigenoR
        objnebulizacion = New NebulizacionR
    End Sub
    Public Overrides Sub guardarDetalleTerapia()
        Try
            HistoriaClinicaDAL.guardarDetalleTerapiaR(objTerapiaFyR)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub guardarOrdenInsumoFisio()
        Try
            If String.IsNullOrEmpty(objInsumosFisio.codigoOrden) Then
                objInsumosFisio.codigoOrden = Constantes.VALOR_PREDETERMINADO
                HistoriaClinicaDAL.guardarDetalleInsumosFisioR(objInsumosFisio)
            Else
                objInsumosFisio.editado = ConstantesHC.VALOR_EDITADO
                HistoriaClinicaDAL.actualizarDetalleInsumosFisioR(objInsumosFisio)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub guardarNotaFisio()
        Try
            If String.IsNullOrEmpty(objNotaFisio.codigoNota) Then
                objNotaFisio.codigoNota = Constantes.VALOR_PREDETERMINADO
                HistoriaClinicaDAL.guardarNotaFisioR(objNotaFisio)
            Else
                objNotaFisio.editado = ConstantesHC.VALOR_EDITADO
                HistoriaClinicaDAL.actualizarNotaFisioR(objNotaFisio)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub guardarOxigeno()
        Try
            HistoriaClinicaDAL.guardarFisioterapiaOxigenoR(objOxigeno)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub guardarNebulizacion()
        Try
            HistoriaClinicaDAL.guardarFisioterapiaNebulizacionR(objnebulizacion)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
