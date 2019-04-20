Public Class SabanaAplicacionMedR
    Inherits SabanaAplicacionMed
    Sub New()
        pMostrarSignos = ConstantesHC.SABANA_SIN_SIGNOS
        sqlDatoPacienteCarga = ConsultasHC.SABANA_PACIENTE_CARGAR_R
        sqlCodigoSabanaCarga = ConsultasHC.SABANA_CODIGO_CARGAR_R
        sqlPesoBalanceCarga = ConsultasHC.SABANA_PESO_CARGAR_R
        sqlDetalleCarga = ConsultasHC.SABANA_SIGNO_VITAL_CARGAR_R
        sqlDetalleCargaIngreso = ConsultasHC.SABANA_INGRESO_CARGAR_R
        sqlDetalleCargaEgreso = ConsultasHC.SABANA_EGRESO_CARGAR_R
        sqlDetalleCargaBalance = ConsultasHC.SABANA_BALANCE_CARGAR_R
        sqlDetalleCargaAplicacion = ConsultasHC.SABANA_APLICACION_MED_CARGAR_R
        sqlDetalleCargaGoteo = ConsultasHC.SABANA_GOTEO_CARGAR_R
        nombreReporte = ConstantesHC.NOMBRE_PDF_SABANA_APLICACION_R
        moduloReporte = Constantes.REPORTE_AM
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
    Public Overrides Sub guardarDetalle()
        cargarCodigoSabana()
        If String.IsNullOrEmpty(codigoSabana) Then
            codigoSabana = Constantes.VALOR_PREDETERMINADO
            SabanaAplicacionMedDAL.guardarSabanaAM(Me)
        Else
            SabanaAplicacionMedDAL.actualizarSabanaAM(Me)
        End If
    End Sub
End Class
