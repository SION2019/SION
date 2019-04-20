Public Class EnfermeriaR
    Inherits Enfermeria

    Sub New()
        moduloReporte = Constantes.REPORTE_AM
        nombreReporte = ConstantesHC.NOMBRE_PDF_GLUCOMETRIA_ENFER_R
        insumoCarga = ConsultasHC.ENFERMERIA_INSUMO_LISTAR_R
        notaCarga = ConsultasHC.ENFERMERIA_NOTA_LISTAR_R
        glucomCargar = ConsultasHC.ENFERMERIA_GLUCOMETRIA_LISTA_R
        examenCargar = ConsultasHC.ENFERMERIA_EXAMENES_LISTA_R
        examenHemoCargar = ConsultasHC.ENFERMERIA_EXAMENES_HEMODERIVADOS_LISTAR_R
        objInsumosEnfer = New InsumoEnfermeriaR
        objNotaEnfer = New NotaEnfermeriaR
        objGlucomEnfer = New GlucometriaEnfermeriaR
        objParaclinico = New ParaclinicoEnfermeriaR
    End Sub

    Public Overrides Sub guardarOrdenInsumoEnfer()
        Try
            If String.IsNullOrEmpty(objInsumosEnfer.codigoOrden) Then
                objInsumosEnfer.codigoOrden = Constantes.VALOR_PREDETERMINADO
                HistoriaClinicaDAL.guardarDetalleInsumosEnferR(objInsumosEnfer)
            Else
                objInsumosEnfer.editado = ConstantesHC.VALOR_EDITADO
                HistoriaClinicaDAL.actualizarDetalleInsumosEnferR(objInsumosEnfer)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub guardarNotaEnfer()
        Try
            If String.IsNullOrEmpty(objNotaEnfer.codigoNota) Then
                objNotaEnfer.codigoNota = Constantes.VALOR_PREDETERMINADO
                HistoriaClinicaDAL.guardarNotaEnferR(objNotaEnfer)
            Else
                objNotaEnfer.editado = ConstantesHC.VALOR_EDITADO
                HistoriaClinicaDAL.actualizarNotaEnferR(objNotaEnfer)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Sub guardarGlucometriaEnfe()
        Try
            HistoriaClinicaDAL.guardarGlucometriaAM(Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
