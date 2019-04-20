Public Class SolicitudMedicamentosNoPos

    Public Property Codigo As String
    Public Property codigoPuntoActual As String
    Public Property registro As String
    Public Property fechaSolicitud As Date
    Public Property codigoEquivalenciaUno As String
    Public Property codigoEquivalenciaDos As String
    Public Property resumen As String
    Public Property rdsi As Boolean
    Public Property rdno As Boolean
    Public Property respuesta As String
    Public Property efecto As String
    Public Property inicioefecto As String
    Public Property explicarazon As String
    Public Property cbocriterio As String
    Public Property cbocriterio2 As String
    Public Property precaucion As String
    Public Property pelipaciente As String
    Public Property cboposibilidad As String
    Public Property reacciones As String
    Public Property contraindicaciones As String
    Public Property duraciontratamiento As String
    Public Property dosis As String
    Public Property frecuencia As String
    Public Property ndosis As String
    Public Property usuarioInforme As String
    Public Property modulo As String
    Public Property dtDiagnosticos As DataTable
    Public Property dr As String
    Public Property registroMedico As String
    Public Property especialidad As String
    Public Property identificacion As String
    Public Property fechaIngreso As DateTime
    Public Property nombrePaciente As String
    Public Property sexo As String
    Public Property edad As String
    Public Property codigoContrato As String
    Public Property contrato As String
    Public Property entorno As String
    Public Property cama As String
    Public Property datosIngreso As String
    Public Property presentacion As String
    Public Property registroInvima As String
    Public Property consulta As String
    Public Property nombreMedicamento As String
    Public Sub guardarMedicamentos()
        MedicamentoNoPosBLL.guardarMedicamentosNOPOS(modulo, Me, dtDiagnosticos)
    End Sub

    Public Sub imprimirReporte()
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre, moduloreporte, nombreReporte As String

            If modulo = Constantes.HC Then
                nombreReporte = ConstantesHC.NOMBRE_PDF_SOLICITUD_MED_NOSPOS
                moduloreporte = Constantes.REPORTE_HC
            ElseIf modulo = Constantes.AM Then
                nombreReporte = ConstantesHC.NOMBRE_PDF_SOLICITUD_MED_NOSPOS_R
                moduloreporte = Constantes.REPORTE_AM
            ElseIf modulo = Constantes.AF Then
                nombreReporte = ConstantesHC.NOMBRE_PDF_SOLICITUD_MED_NOSPOS_RR
                moduloreporte = Constantes.REPORTE_AF
            Else
                nombreReporte = ""
            End If

            Dim reportes As Object = Nothing
            Dim tipo As Integer
            Dim dt As New DataTable
            General.llenarTablaYdgv(Consultas.SOLICITUD_MED_FORMATO_EPS & registro, dt)
            tipo = dt.Rows(0).Item("Codigo_Reporte").ToString
            Select Case tipo
                Case Constantes.COD_REPORTE_NOPOS_GENERAL
                    reportes = New rptSolicitud_Medicamento
                Case Constantes.COD_REPORTE_NOPOS_AMBUQ
                    reportes = New rptAMBUQ
                Case Constantes.COD_REPORTE_NOPOS_CAJACOPI
                    reportes = New rptCaja_Copi
                Case Constantes.COD_REPORTE_NOPOS_COMPARTA
                    reportes = New rptComparta
                Case Constantes.COD_REPORTE_NOPOS_COOMEVA
                    reportes = New rpt_Coomeva
            End Select

            codigoNombre = Codigo
            nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = Path.GetTempPath() & nombreArchivo
            formula = "{VISTA_SOLICITUD_MED_NOPOS.codigo} = " & codigoNombre & " AND {VISTA_SOLICITUD_MED_NOPOS.Modulo}=" & moduloreporte & ""
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(nombreReporte, Codigo, reportes,
                                  codigoNombre, formula,
                                  nombreReporte, IO.Path.GetTempPath,,)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub cargarMedico()
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(usuarioInforme)
        params.Add(SesionActual.idEmpresa)
        General.llenarTabla(Consultas.EMPLEADO_DATOS_CARGAR, params, dt)
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0).Item("Empleado").ToString
            registroMedico = dt.Rows(0).Item("Registro_Medico").ToString
            especialidad = dt.Rows(0).Item("Descripcion_Especialidad").ToString
        End If
    End Sub

    Public Sub cargarPaciente()
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(Codigo)
        General.llenarTabla(Consultas.PACIENTE_DATOS_CARGAR, params, dt)
        If dt.Rows.Count > 0 Then
            registro = dt.Rows(0).Item("Registro").ToString
            identificacion = dt.Rows(0).Item("Identificacion del paciente")
            nombrePaciente = dt.Rows(0).Item("Paciente").ToString
            sexo = dt.Rows(0).Item("Genero").ToString
            edad = dt.Rows(0).Item("Edad").ToString
            codigoContrato = dt.Rows(0).Item("Codigo contrato").ToString
            contrato = dt.Rows(0).Item("Nombre Eps").ToString
            entorno = dt.Rows(0).Item("Area de servicio").ToString

            datosIngreso = dt.Rows(0).Item("Motivo ingreso").ToString
        End If
    End Sub

    Public Sub cargarPacienteh()
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(Consultas.CARGAR_PACIENTE_NOPOS, params, dt)
        If dt.Rows.Count > 0 Then
            registro = dt.Rows(0).Item("Registro").ToString
            identificacion = dt.Rows(0).Item("Identificacion del paciente")
            nombrePaciente = dt.Rows(0).Item("Paciente").ToString
            sexo = dt.Rows(0).Item("Genero").ToString
            edad = dt.Rows(0).Item("Edad").ToString
            codigoContrato = dt.Rows(0).Item("Codigo contrato").ToString
            contrato = dt.Rows(0).Item("Nombre Eps").ToString
            entorno = dt.Rows(0).Item("Area de servicio").ToString

            datosIngreso = dt.Rows(0).Item("Motivo ingreso").ToString
        End If
    End Sub

    Public Sub cargarPacienteR()
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(Codigo)
        General.llenarTabla(Consultas.PACIENTE_DATOS_CARGAR_R, params, dt)
        If dt.Rows.Count > 0 Then
            registro = dt.Rows(0).Item("Registro").ToString
            identificacion = dt.Rows(0).Item("Identificacion del paciente")
            nombrePaciente = dt.Rows(0).Item("Paciente").ToString
            sexo = dt.Rows(0).Item("Genero").ToString
            edad = dt.Rows(0).Item("Edad").ToString
            codigoContrato = dt.Rows(0).Item("Codigo contrato").ToString
            contrato = dt.Rows(0).Item("Nombre Eps").ToString
            entorno = dt.Rows(0).Item("Area de servicio").ToString

            datosIngreso = dt.Rows(0).Item("Motivo ingreso").ToString
        End If

    End Sub

    Public Sub cargarPacienteRR()
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(Codigo)
        General.llenarTabla(Consultas.PACIENTE_DATOS_CARGAR_RR, params, dt)
        If dt.Rows.Count > 0 Then
            registro = dt.Rows(0).Item("Registro").ToString
            identificacion = dt.Rows(0).Item("Identificacion del paciente")
            nombrePaciente = dt.Rows(0).Item("Paciente").ToString
            sexo = dt.Rows(0).Item("Genero").ToString
            edad = dt.Rows(0).Item("Edad").ToString
            codigoContrato = dt.Rows(0).Item("Codigo contrato").ToString
            contrato = dt.Rows(0).Item("Nombre Eps").ToString
            entorno = dt.Rows(0).Item("Area de servicio").ToString

            datosIngreso = dt.Rows(0).Item("Motivo ingreso").ToString
        End If

    End Sub

    Public Sub cargarDemasRegistro()
        If modulo = Constantes.HC Then
            Dim dt As New DataTable
            Dim params As New List(Of String)
            params.Add(Codigo)
            General.llenarTabla(Consultas.SOLIC_MED_CARGAR_REGISTROS, params, dt)
            If dt.Rows.Count > 0 Then
                resumen = dt.Rows(0).Item("Resumen_hc").ToString
                If dt.Rows(0).Item("Si").ToString Then
                    rdsi = True
                    rdno = False
                Else
                    rdsi = False
                    rdno = True
                End If
                respuesta = dt.Rows(0).Item("RespuestaPOS").ToString
                efecto = dt.Rows(0).Item("Efecto_Deseado").ToString
                inicioefecto = dt.Rows(0).Item("Iniciacion_Efecto").ToString
                explicarazon = dt.Rows(0).Item("Explicar_Razon").ToString
                cbocriterio = dt.Rows(0).Item("Criterio1").ToString
                cbocriterio2 = dt.Rows(0).Item("Criterio2").ToString
                precaucion = dt.Rows(0).Item("PreCoEsTox").ToString
                pelipaciente = dt.Rows(0).Item("PeligroPaciente").ToString
                cboposibilidad = dt.Rows(0).Item("PosibilidadTerapeutica").ToString
                reacciones = dt.Rows(0).Item("PosibilidadTerapeuticaPOS").ToString
                contraindicaciones = dt.Rows(0).Item("PosibilidadTerapeuticaPOScontra").ToString
                duraciontratamiento = dt.Rows(0).Item("Duracion_Tto").ToString
                dosis = dt.Rows(0).Item("Dosis").ToString
                frecuencia = dt.Rows(0).Item("Frecuencia").ToString
                ndosis = dt.Rows(0).Item("Num_Dosis").ToString
                presentacion = dt.Rows(0).Item("Nombre")
                registroInvima = dt.Rows(0).Item("Registro_Sanitario")
                fechaSolicitud = dt.Rows(0).Item("Fecha_Solicitud")
                usuarioInforme = dt.Rows(0).Item("Usuario")
            End If

        ElseIf modulo = Constantes.AM Then
            Dim dt As New DataTable
            Dim params As New List(Of String)
            params.Add(Codigo)
            General.llenarTabla(Consultas.SOLIC_MED_CARGAR_REGISTROS_R, params, dt)
            If dt.Rows.Count > 0 Then
                resumen = dt.Rows(0).Item("Resumen_hc").ToString
                If dt.Rows(0).Item("Si").ToString Then
                    rdsi = True
                    rdno = False
                Else
                    rdsi = False
                    rdno = True
                End If
                respuesta = dt.Rows(0).Item("RespuestaPOS").ToString
                efecto = dt.Rows(0).Item("Efecto_Deseado").ToString
                inicioefecto = dt.Rows(0).Item("Iniciacion_Efecto").ToString
                explicarazon = dt.Rows(0).Item("Explicar_Razon").ToString
                cbocriterio = dt.Rows(0).Item("Criterio1").ToString
                cbocriterio2 = dt.Rows(0).Item("Criterio2").ToString
                precaucion = dt.Rows(0).Item("PreCoEsTox").ToString
                pelipaciente = dt.Rows(0).Item("PeligroPaciente").ToString
                cboposibilidad = dt.Rows(0).Item("PosibilidadTerapeutica").ToString
                reacciones = dt.Rows(0).Item("PosibilidadTerapeuticaPOS").ToString
                contraindicaciones = dt.Rows(0).Item("PosibilidadTerapeuticaPOScontra").ToString
                duraciontratamiento = dt.Rows(0).Item("Duracion_Tto").ToString
                dosis = dt.Rows(0).Item("Dosis").ToString
                frecuencia = dt.Rows(0).Item("Frecuencia").ToString
                ndosis = dt.Rows(0).Item("Num_Dosis").ToString
                presentacion = dt.Rows(0).Item("Nombre")
                registroInvima = dt.Rows(0).Item("Registro_Sanitario")
                fechaSolicitud = dt.Rows(0).Item("Fecha_Solicitud")
                usuarioInforme = dt.Rows(0).Item("Usuario")
            End If

        ElseIf modulo = Constantes.AF Then
            Dim dt As New DataTable
            Dim params As New List(Of String)
            params.Add(Codigo)
            General.llenarTabla(Consultas.SOLIC_MED_CARGAR_REGISTROS_RR, params, dt)
            If dt.Rows.Count > 0 Then
                resumen = dt.Rows(0).Item("Resumen_hc").ToString
                If dt.Rows(0).Item("Si").ToString Then
                    rdsi = True
                    rdno = False
                Else
                    rdsi = False
                    rdno = True
                End If
                respuesta = dt.Rows(0).Item("RespuestaPOS").ToString
                efecto = dt.Rows(0).Item("Efecto_Deseado").ToString
                inicioefecto = dt.Rows(0).Item("Iniciacion_Efecto").ToString
                explicarazon = dt.Rows(0).Item("Explicar_Razon").ToString
                cbocriterio = dt.Rows(0).Item("Criterio1").ToString
                cbocriterio2 = dt.Rows(0).Item("Criterio2").ToString
                precaucion = dt.Rows(0).Item("PreCoEsTox").ToString
                pelipaciente = dt.Rows(0).Item("PeligroPaciente").ToString
                cboposibilidad = dt.Rows(0).Item("PosibilidadTerapeutica").ToString
                reacciones = dt.Rows(0).Item("PosibilidadTerapeuticaPOS").ToString
                contraindicaciones = dt.Rows(0).Item("PosibilidadTerapeuticaPOScontra").ToString
                duraciontratamiento = dt.Rows(0).Item("Duracion_Tto").ToString
                dosis = dt.Rows(0).Item("Dosis").ToString
                frecuencia = dt.Rows(0).Item("Frecuencia").ToString
                ndosis = dt.Rows(0).Item("Num_Dosis").ToString
                presentacion = dt.Rows(0).Item("Nombre")
                registroInvima = dt.Rows(0).Item("Registro_Sanitario")
                fechaSolicitud = dt.Rows(0).Item("Fecha_Solicitud")
                usuarioInforme = dt.Rows(0).Item("Usuario")
            End If
        End If
    End Sub

    Public Sub cargarDiagnosticoSolicitud()
        If modulo = Constantes.HC Then
            consulta = Consultas.DIAGNOSTICO_SOLICITUD_NOPOS_CARGAR
        ElseIf modulo = Constantes.AM Then
            consulta = Consultas.DIAGNOSTICO_SOLICITUD_NOPOS_CARGARR
        ElseIf modulo = Constantes.AF Then
            consulta = Consultas.DIAGNOSTICO_SOLICITUD_NOPOS_CARGARRR
        End If
        Dim params As New List(Of String)
        params.Add(Codigo)
        dtDiagnosticos = New DataTable
        General.llenarTabla(consulta, params, dtDiagnosticos)

    End Sub

    Public Sub cargarDatosAutomaticos()
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(codigoEquivalenciaUno)
        General.llenarTabla(Consultas.SOLIC_MEDI_NOPOS_CARGAR, params, dt)
        If dt.Rows.Count > 0 Then
            codigoEquivalenciaDos = dt.Rows(0).Item("Codigo interno2").ToString
            nombreMedicamento = dt.Rows(0).Item("Descripcion").ToString
            resumen = dt.Rows(0).Item("Justificacion").ToString
            explicarazon = dt.Rows(0).Item("No se seguira usando").ToString
            efecto = dt.Rows(0).Item("Efecto Deseado").ToString
            inicioefecto = dt.Rows(0).Item("Efecto iniciado").ToString
            presentacion = dt.Rows(0).Item("Presentacion").ToString
            duraciontratamiento = dt.Rows(0).Item("Dias tratamiento").ToString
            ndosis = dt.Rows(0).Item("Numero dosis").ToString

            If dt.Rows(0).Item("Existe otro med").ToString = True Then
                rdsi = True
                rdno = False
            Else
                rdsi = False
                rdno = True
            End If
        End If
    End Sub

    Public Sub cargarDiagnosticoEvo()
        If modulo = Constantes.HC Then
            consulta = Consultas.DIAGNOSTICO_EVOLUCION_NOPOS_CARGAR
        ElseIf modulo = Constantes.AM Then
            consulta = Consultas.DIAGNOSTICO_EVOLUCION_NOPOS_CARGAR_R
        ElseIf modulo = Constantes.AF Then
            consulta = Consultas.DIAGNOSTICO_EVOLUCION_NOPOS_CARGAR_RR
        End If
        Dim params As New List(Of String)
        params.Add(registro)
        dtDiagnosticos = New DataTable
        General.llenarTabla(consulta, params, dtDiagnosticos)
        If dtDiagnosticos.Rows.Count = 0 Then
            If modulo = Constantes.HC Then
                consulta = Consultas.DIAGNOSTICO_EVOLUCION_NOPOS_CARGAR2
            ElseIf modulo = Constantes.AM Then
                consulta = Consultas.DIAGNOSTICO_EVOLUCION_NOPOS_CARGAR2_R
            ElseIf modulo = Constantes.AF Then
                consulta = Consultas.DIAGNOSTICO_EVOLUCION_NOPOS_CARGAR2_RR
            End If
            General.llenarTabla(consulta, params, dtDiagnosticos)
        End If
    End Sub
End Class
