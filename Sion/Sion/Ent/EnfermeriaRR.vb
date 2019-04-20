Public Class EnfermeriaRR
    Inherits Enfermeria

    Sub New()
        moduloReporte = Constantes.REPORTE_AF
        nombreReporte = ConstantesHC.NOMBRE_PDF_GLUCOMETRIA_ENFER_RR
        insumoCarga = ConsultasHC.ENFERMERIA_INSUMO_LISTAR_RR
        notaCarga = ConsultasHC.ENFERMERIA_NOTA_LISTAR_RR
        glucomCargar = ConsultasHC.ENFERMERIA_GLUCOMETRIA_LISTA_RR
        examenCargar = ConsultasHC.ENFERMERIA_EXAMENES_LISTA_RR
        examenHemoCargar = ConsultasHC.ENFERMERIA_EXAMENES_HEMODERIVADOS_LISTAR_RR
        objInsumosEnfer = New InsumoEnfermeriaRR
        objNotaEnfer = New NotaEnfermeriaRR
        objGlucomEnfer = New GlucometriaEnfermeriaRR
        objParaclinico = New ParaclinicoEnfermeriaRR
    End Sub
    Public Overrides Sub guardarOrdenInsumoEnfer()
        Try
            If String.IsNullOrEmpty(objInsumosEnfer.codigoOrden) Then
                objInsumosEnfer.codigoOrden = Constantes.VALOR_PREDETERMINADO
                HistoriaClinicaDAL.guardarDetalleInsumosEnferRR(objInsumosEnfer)
            Else
                objInsumosEnfer.editado = ConstantesHC.VALOR_EDITADO
                HistoriaClinicaDAL.actualizarDetalleInsumosEnferRR(objInsumosEnfer)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub guardarNotaEnfer()
        Try
            If String.IsNullOrEmpty(objNotaEnfer.codigoNota) Then
                objNotaEnfer.codigoNota = Constantes.VALOR_PREDETERMINADO
                HistoriaClinicaDAL.guardarNotaEnferRR(objNotaEnfer)
            Else
                objNotaEnfer.editado = ConstantesHC.VALOR_EDITADO
                HistoriaClinicaDAL.actualizarNotaEnferRR(objNotaEnfer)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub guardarGlucometriaEnfe()
        Try
            HistoriaClinicaDAL.guardarGlucometriaAF(Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
