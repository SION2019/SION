Public Class SabanaAplicacionMedRR
    Inherits SabanaAplicacionMed
    Sub New()
        pMostrarSignos = ConstantesHC.SABANA_SIN_SIGNOS
        sqlDatoPacienteCarga = ConsultasHC.SABANA_PACIENTE_CARGAR_RR
        sqlCodigoSabanaCarga = ConsultasHC.SABANA_CODIGO_CARGAR_RR
        sqlPesoBalanceCarga = ConsultasHC.SABANA_PESO_CARGAR_RR
        sqlDetalleCarga = ConsultasHC.SABANA_SIGNO_VITAL_CARGAR_RR
        sqlDetalleCargaIngreso = ConsultasHC.SABANA_INGRESO_CARGAR_RR
        sqlDetalleCargaEgreso = ConsultasHC.SABANA_EGRESO_CARGAR_RR
        sqlDetalleCargaBalance = ConsultasHC.SABANA_BALANCE_CARGAR_RR
        sqlDetalleCargaAplicacion = ConsultasHC.SABANA_APLICACION_MED_CARGAR_RR
        sqlDetalleCargaGoteo = ConsultasHC.SABANA_GOTEO_CARGAR_RR
        nombreReporte = ConstantesHC.NOMBRE_PDF_SABANA_APLICACION_RR
        moduloReporte = Constantes.REPORTE_AF
    End Sub
    Public Overrides Sub cargarDetalle()
        identificadorPredeterminado = inicialesUsuario
        cargarDetalleSignoVital()
        cargarDetalleIngreso()
        cargarDetalleEgreso()
        cargarDetalleBalance()
        cargarDetalleAplicacion()
        cargarDetalleGoteo()
    End Sub
    Public Overrides Sub guardarDetalle()
        cargarCodigoSabana()
        If String.IsNullOrEmpty(codigoSabana) Then
            codigoSabana = Constantes.VALOR_PREDETERMINADO
            SabanaAplicacionMedDAL.guardarSabanaAF(Me)
        Else
            SabanaAplicacionMedDAL.actualizarSabanaAF(Me)
        End If
    End Sub
    Public Overrides Sub cargarCodigoSabana()
        Try
            Dim params As New List(Of String)
            Dim dr As DataRow
            params.Add(nRegistro)
            params.Add(Format(fechaReal.Date, Constantes.FORMATO_FECHA_GEN))
            dr = General.cargarItem(sqlCodigoSabanaCarga, params)
            If Not IsNothing(dr) Then
                codigoSabana = dr.Item(0)
                editado = dr.Item(1)
            Else
                codigoSabana = ""
                editado = False
            End If
            usuario = SesionActual.idUsuario
            cambiarUsuarioAutomatico()
            cargarIniciales()
        Catch ex As Exception
            general.manejoErrores(ex) 
        End Try

    End Sub
End Class
