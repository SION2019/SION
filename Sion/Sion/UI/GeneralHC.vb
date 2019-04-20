Imports System.Collections
Public Class GeneralHC

    Public Shared Property fabricaHC As Fabrica(Of String, Object)
    Private Function crearObj(Of objEspecificio As {InfoIngresoAdulto, New})() As Object
        ''ESO DE {InfoIngresoAdulto, New} ESTÁ BIEN???
        Return New objEspecificio()
    End Function

    Public Shared Sub cargarHT()

        fabricaHC = New Fabrica(Of String, Object)

        ' instancias de informaciond de ingreso
        fabricaHC.registrar(Of InfoIngresoAdulto)(ConstantesHC.CODIGO_INFO_INGRESO)
        fabricaHC.registrar(Of InfoIngresoAdultoR)(ConstantesHC.CODIGO_INFO_INGRESO_R)
        fabricaHC.registrar(Of InfoIngresoAdultoRR)(ConstantesHC.CODIGO_INFO_INGRESO_RR)
        fabricaHC.registrar(Of InfoIngresoNeonato)(ConstantesHC.CODIGO_INFO_INGRESO_N)
        fabricaHC.registrar(Of InfoIngresoNeonatoR)(ConstantesHC.CODIGO_INFO_INGRESO_N_R)
        fabricaHC.registrar(Of InfoIngresoNeonatoRR)(ConstantesHC.CODIGO_INFO_INGRESO_N_RR)

        ' instancias de evolucion medica 
        fabricaHC.registrar(Of EvolucionMedica)(ConstantesHC.CODIGO_EVOLUCION)
        fabricaHC.registrar(Of EvolucionMedicaR)(ConstantesHC.CODIGO_EVOLUCION_R)
        fabricaHC.registrar(Of EvolucionMedicaRR)(ConstantesHC.CODIGO_EVOLUCION_RR)
        ' instancias de interconsulta medica 
        fabricaHC.registrar(Of InterconsultaMedica)(ConstantesHC.CODIGO_INTERCONSULTA_MEDICA)
        fabricaHC.registrar(Of InterconsultaMedicaR)(ConstantesHC.CODIGO_INTERCONSULTA_MEDICA_R)
        fabricaHC.registrar(Of InterconsultaMedicaRR)(ConstantesHC.CODIGO_INTERCONSULTA_MEDICA_RR)
        ' instancias de remision
        fabricaHC.registrar(Of Remision)(ConstantesHC.CODIGO_REMISION)
        fabricaHC.registrar(Of RemisionR)(ConstantesHC.CODIGO_REMISION_R)
        fabricaHC.registrar(Of RemisionRR)(ConstantesHC.CODIGO_REMISION_RR)
        ' instancias de orden medica 
        fabricaHC.registrar(Of OrdenMedica)(ConstantesHC.CODIGO_ORDEN_MEDICA)
        fabricaHC.registrar(Of OrdenMedicaR)(ConstantesHC.CODIGO_ORDEN_MEDICA_R)
        fabricaHC.registrar(Of OrdenMedicaRR)(ConstantesHC.CODIGO_ORDEN_MEDICA_RR)
        ' instancais de fisioterapia
        fabricaHC.registrar(Of Fisioterapia)(ConstantesHC.CODIGO_FISIOTERAPIA)
        fabricaHC.registrar(Of FisioterapiaR)(ConstantesHC.CODIGO_FISIOTERAPIA_R)
        fabricaHC.registrar(Of FisioterapiaRR)(ConstantesHC.CODIGO_FISIOTERAPIA_RR)

        'Instancias de enfermeria
        fabricaHC.registrar(Of Enfermeria)(ConstantesHC.CODIGO_ENFERMERIA)
        fabricaHC.registrar(Of EnfermeriaR)(ConstantesHC.CODIGO_ENFERMERIA_R)
        fabricaHC.registrar(Of EnfermeriaRR)(ConstantesHC.CODIGO_ENFERMERIA_RR)

        'Instancias de vista previa
        fabricaHC.registrar(Of VistaPrevia)(ConstantesHC.CODIGO_VISTA_PREVIA)
        fabricaHC.registrar(Of VistaPreviaR)(ConstantesHC.CODIGO_VISTA_PREVIA_R)
        fabricaHC.registrar(Of VistaPreviaRR)(ConstantesHC.CODIGO_VISTA_PREVIA_RR)
        fabricaHC.registrar(Of AtencionLaboratorio)(ConstantesHC.CODIGO_VISTA_PREVIA_ATENCION)

        'Instancias de Sabana de aplicacion de medicamentos
        fabricaHC.registrar(Of Prefactura)(ConstantesHC.CODIGO_PREFACTURA)
        fabricaHC.registrar(Of PrefacturaR)(ConstantesHC.CODIGO_PREFACTURA_R)
        fabricaHC.registrar(Of PrefacturaRR)(ConstantesHC.CODIGO_PREFACTURA_RR)

        'Instancias de parametros de ventilacion
        fabricaHC.registrar(Of ParametroVentilacion)(ConstantesHC.CODIGO_PARAMETROS_VENT)
        fabricaHC.registrar(Of ParametroVentilacionR)(ConstantesHC.CODIGO_PARAMETROS_VENT_R)
        fabricaHC.registrar(Of ParametroVentilacionRR)(ConstantesHC.CODIGO_PARAMETROS_VENT_RR)

        'Instancias de sabana de aplicacion de medicamentos
        fabricaHC.registrar(Of SabanaAplicacionMed)(ConstantesHC.CODIGO_SABANA)
        fabricaHC.registrar(Of SabanaAplicacionMedR)(ConstantesHC.CODIGO_SABANA_R)
        fabricaHC.registrar(Of SabanaAplicacionMedRR)(ConstantesHC.CODIGO_SABANA_RR)

        'Instancias de Ecocardiograma
        fabricaHC.registrar(Of Ecocardiograma)(Constantes.CODIGO_ECOCARDIOGRAMA & Constantes.ECO & Constantes.HC)
        fabricaHC.registrar(Of EcocardiogramaHemo)(Constantes.CODIGO_ECOCARDIOGRAMA & Constantes.ECO & Constantes.CODIGO_MENU_HEMO)
        fabricaHC.registrar(Of EcocardiogramaQuir)(Constantes.CODIGO_ECOCARDIOGRAMA & Constantes.ECO & Constantes.CODIGO_MENU_QUIR)
        fabricaHC.registrar(Of EcocardiogramaHosp)(Constantes.CODIGO_ECOCARDIOGRAMA & Constantes.ECO & Constantes.CODIGO_MENU_HOSP)
        fabricaHC.registrar(Of EcocardiogramaUrg)(Constantes.CODIGO_ECOCARDIOGRAMA & Constantes.ECO & Constantes.CODIGO_MENU_URG)
        fabricaHC.registrar(Of EcocardiogramaCExt)(Constantes.CODIGO_ECOCARDIOGRAMA & Constantes.ECO & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of EcocardiogramaR)(Constantes.CODIGO_ECOCARDIOGRAMA & Constantes.ECO & Constantes.AM)
        fabricaHC.registrar(Of EcocardiogramaRR)(Constantes.CODIGO_ECOCARDIOGRAMA & Constantes.ECO & Constantes.AF)

        'Instancias de Ecocardiograma Stres
        fabricaHC.registrar(Of EcocardiogramaEstres)(Constantes.CODIGO_ECO_ESTRES & Constantes.ECOSTRES & Constantes.HC)
        fabricaHC.registrar(Of EcocardiogramaEstresHemo)(Constantes.CODIGO_ECO_ESTRES & Constantes.ECOSTRES & Constantes.CODIGO_MENU_HEMO)
        fabricaHC.registrar(Of EcocardiogramaEstresQuir)(Constantes.CODIGO_ECO_ESTRES & Constantes.ECOSTRES & Constantes.CODIGO_MENU_QUIR)
        fabricaHC.registrar(Of EcocardiogramaEstresHosp)(Constantes.CODIGO_ECO_ESTRES & Constantes.ECOSTRES & Constantes.CODIGO_MENU_HOSP)
        fabricaHC.registrar(Of EcocardiogramaEstresUrg)(Constantes.CODIGO_ECO_ESTRES & Constantes.ECOSTRES & Constantes.CODIGO_MENU_URG)
        fabricaHC.registrar(Of EcocardiogramaEstresCExt)(Constantes.CODIGO_ECO_ESTRES & Constantes.ECOSTRES & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of EcocardiogramaEstresR)(Constantes.CODIGO_ECO_ESTRES & Constantes.ECOSTRES & Constantes.AM)
        fabricaHC.registrar(Of EcocardiogramaEstresRR)(Constantes.CODIGO_ECO_ESTRES & Constantes.ECOSTRES & Constantes.AF)

        'Instancias Hoja de ruta 
        fabricaHC.registrar(Of HojaRuta)(Constantes.CODIGO_HOJA_RUTA & Constantes.HC)
        fabricaHC.registrar(Of HojaRutaR)(Constantes.CODIGO_HOJA_RUTA & Constantes.AM)
        fabricaHC.registrar(Of HojaRutaRR)(Constantes.CODIGO_HOJA_RUTA & Constantes.AF)
        fabricaHC.registrar(Of HojaRutaExterna)(Constantes.CODIGO_HOJA_RUTA & Constantes.EX)
        'Instancias informe qx
        fabricaHC.registrar(Of InformeQx)(Constantes.CODIGO_INFORME_QX & Constantes.HC)
        fabricaHC.registrar(Of InformeQxHemo)(Constantes.CODIGO_INFORME_QX & Constantes.CODIGO_MENU_HEMO)
        fabricaHC.registrar(Of InformeQxQuir)(Constantes.CODIGO_INFORME_QX & Constantes.CODIGO_MENU_QUIR)
        fabricaHC.registrar(Of InformeQxHosp)(Constantes.CODIGO_INFORME_QX & Constantes.CODIGO_MENU_HOSP)
        fabricaHC.registrar(Of InformeQxUrg)(Constantes.CODIGO_INFORME_QX & Constantes.CODIGO_MENU_URG)
        fabricaHC.registrar(Of InformeQxCExt)(Constantes.CODIGO_INFORME_QX & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of InformeQxR)(Constantes.CODIGO_INFORME_QX & Constantes.AM)
        fabricaHC.registrar(Of InformeQxRR)(Constantes.CODIGO_INFORME_QX & Constantes.AF)

        'Instancias Hemodialisis
        fabricaHC.registrar(Of Hemodialisis)(Constantes.CODIGO_HEMODIALISIS & Constantes.HC)
        fabricaHC.registrar(Of HemodialisisHemo)(Constantes.CODIGO_HEMODIALISIS & Constantes.CODIGO_MENU_HEMO)
        fabricaHC.registrar(Of HemodialisisQuir)(Constantes.CODIGO_HEMODIALISIS & Constantes.CODIGO_MENU_QUIR)
        fabricaHC.registrar(Of HemodialisisHosp)(Constantes.CODIGO_HEMODIALISIS & Constantes.CODIGO_MENU_HOSP)
        fabricaHC.registrar(Of HemodialisisUrg)(Constantes.CODIGO_HEMODIALISIS & Constantes.CODIGO_MENU_URG)
        fabricaHC.registrar(Of HemodialisisCExt)(Constantes.CODIGO_HEMODIALISIS & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of HemodialisisR)(Constantes.CODIGO_HEMODIALISIS & Constantes.AM)
        fabricaHC.registrar(Of HemodialisisRR)(Constantes.CODIGO_HEMODIALISIS & Constantes.AF)

        'Instancias Comite
        fabricaHC.registrar(Of ComiteCTC)(Constantes.CODIGO_COMITE & Constantes.HC)
        fabricaHC.registrar(Of ComiteCTCR)(Constantes.CODIGO_COMITE & Constantes.AM)
        fabricaHC.registrar(Of ComiteCTCRR)(Constantes.CODIGO_COMITE & Constantes.AF)

        'Instancias ECO
        fabricaHC.registrar(Of Eco)(Constantes.CODIGO_ECO & Constantes.HC)
        fabricaHC.registrar(Of EcoHemo)(Constantes.CODIGO_ECO & Constantes.CODIGO_MENU_HEMO)
        fabricaHC.registrar(Of EcoQuir)(Constantes.CODIGO_ECO & Constantes.CODIGO_MENU_QUIR)
        fabricaHC.registrar(Of EcoHosp)(Constantes.CODIGO_ECO & Constantes.CODIGO_MENU_HOSP)
        fabricaHC.registrar(Of EcoUrg)(Constantes.CODIGO_ECO & Constantes.CODIGO_MENU_URG)
        fabricaHC.registrar(Of EcoCExt)(Constantes.CODIGO_ECO & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of EcoR)(Constantes.CODIGO_ECO & Constantes.AM)
        fabricaHC.registrar(Of EcoRR)(Constantes.CODIGO_ECO & Constantes.AF)

        'Justificacion antibiotico'
        fabricaHC.registrar(Of JustificacionMedicamento)(Constantes.CODIGO_ANTIBIOTICO & Constantes.HC)
        fabricaHC.registrar(Of JustificacionMedicamentoHemo)(Constantes.CODIGO_ANTIBIOTICO & Constantes.CODIGO_MENU_HEMO)
        fabricaHC.registrar(Of JustificacionMedicamentoQuir)(Constantes.CODIGO_ANTIBIOTICO & Constantes.CODIGO_MENU_QUIR)
        fabricaHC.registrar(Of JustificacionMedicamentoHosp)(Constantes.CODIGO_ANTIBIOTICO & Constantes.CODIGO_MENU_HOSP)
        fabricaHC.registrar(Of JustificacionMedicamentoUrg)(Constantes.CODIGO_ANTIBIOTICO & Constantes.CODIGO_MENU_URG)
        fabricaHC.registrar(Of JustificacionMedicamentoCExt)(Constantes.CODIGO_ANTIBIOTICO & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of JustificacionMedicamentoR)(Constantes.CODIGO_ANTIBIOTICO & Constantes.AM)
        fabricaHC.registrar(Of JustificacionMedicamentoRR)(Constantes.CODIGO_ANTIBIOTICO & Constantes.AF)

        'Epicrisis'
        fabricaHC.registrar(Of Epicrisis)(Constantes.CODIGO_EPICRISIS & Constantes.HC)
        fabricaHC.registrar(Of EpicrisisHemo)(Constantes.CODIGO_EPICRISIS & Constantes.CODIGO_MENU_HEMO)
        fabricaHC.registrar(Of EpicrisisQuir)(Constantes.CODIGO_EPICRISIS & Constantes.CODIGO_MENU_QUIR)
        fabricaHC.registrar(Of EpicrisisHosp)(Constantes.CODIGO_EPICRISIS & Constantes.CODIGO_MENU_HOSP)
        fabricaHC.registrar(Of EpicrisisUrg)(Constantes.CODIGO_EPICRISIS & Constantes.CODIGO_MENU_URG)
        fabricaHC.registrar(Of EpicrisisCExt)(Constantes.CODIGO_EPICRISIS & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of EpicrisisR)(Constantes.CODIGO_EPICRISIS & Constantes.AM)
        fabricaHC.registrar(Of EpicrisisRR)(Constantes.CODIGO_EPICRISIS & Constantes.AF)

        'Informe Cateterismo'
        fabricaHC.registrar(Of CateterismoCardiaco)(Constantes.CODIGO_CATETERISMO & Constantes.HC)
        fabricaHC.registrar(Of CateterismoCardiacoHemo)(Constantes.CODIGO_CATETERISMO & Constantes.CODIGO_MENU_HEMO)
        fabricaHC.registrar(Of CateterismoCardiacoQuir)(Constantes.CODIGO_CATETERISMO & Constantes.CODIGO_MENU_QUIR)
        fabricaHC.registrar(Of CateterismoCardiacoHosp)(Constantes.CODIGO_CATETERISMO & Constantes.CODIGO_MENU_HOSP)
        fabricaHC.registrar(Of CateterismoCardiacoUrg)(Constantes.CODIGO_CATETERISMO & Constantes.CODIGO_MENU_URG)
        fabricaHC.registrar(Of CateterismoCardiacoCExt)(Constantes.CODIGO_CATETERISMO & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of CateterismoCardiacoR)(Constantes.CODIGO_CATETERISMO & Constantes.AM)
        fabricaHC.registrar(Of CateterismoCardiacoRR)(Constantes.CODIGO_CATETERISMO & Constantes.AF)

        'Lista de chequeo'
        fabricaHC.registrar(Of ListaCheck)(Constantes.CODIGO_LISTA_CHEQUEO & Constantes.HC)
        fabricaHC.registrar(Of ListaCheckHemo)(Constantes.CODIGO_LISTA_CHEQUEO & Constantes.CODIGO_MENU_HEMO)
        fabricaHC.registrar(Of ListaCheckQuir)(Constantes.CODIGO_LISTA_CHEQUEO & Constantes.CODIGO_MENU_QUIR)
        fabricaHC.registrar(Of ListaCheckHosp)(Constantes.CODIGO_LISTA_CHEQUEO & Constantes.CODIGO_MENU_HOSP)
        fabricaHC.registrar(Of ListaCheckUrg)(Constantes.CODIGO_LISTA_CHEQUEO & Constantes.CODIGO_MENU_URG)
        fabricaHC.registrar(Of ListaCheckCExt)(Constantes.CODIGO_LISTA_CHEQUEO & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of ListaCheckR)(Constantes.CODIGO_LISTA_CHEQUEO & Constantes.AM)
        fabricaHC.registrar(Of ListaCheckRR)(Constantes.CODIGO_LISTA_CHEQUEO & Constantes.AF)

        'Epicrisis'
        fabricaHC.registrar(Of EstanciaProlongada)(ConstantesHC.CODIGO_ESTANCIA_PROLONGADA)
        fabricaHC.registrar(Of EstanciaProlongadaR)(ConstantesHC.CODIGO_ESTANCIA_PROLONGADA_R)
        fabricaHC.registrar(Of EstanciaProlongadaRR)(ConstantesHC.CODIGO_ESTANCIA_PROLONGADA_RR)
        ' instancias de resultados Laboratorio
        fabricaHC.registrar(Of ParaclinicoLaboratorio)(ConstantesHC.CODIGO_EXAMEN_LAB & Constantes.HC)
        fabricaHC.registrar(Of ParaclinicoLaboratorio)(ConstantesHC.CODIGO_EXAMEN_LAB & Constantes.CODIGO_MENU_HEMO)
        fabricaHC.registrar(Of ParaclinicoLaboratorio)(ConstantesHC.CODIGO_EXAMEN_LAB & Constantes.CODIGO_MENU_QUIR)
        fabricaHC.registrar(Of ParaclinicoLaboratorio)(ConstantesHC.CODIGO_EXAMEN_LAB & Constantes.CODIGO_MENU_HOSP)
        fabricaHC.registrar(Of ParaclinicoLaboratorio)(ConstantesHC.CODIGO_EXAMEN_LAB & Constantes.CODIGO_MENU_URG)
        fabricaHC.registrar(Of ParaclinicoLaboratorio)(ConstantesHC.CODIGO_EXAMEN_LAB & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of ParaclinicoLaboratorioR)(ConstantesHC.CODIGO_EXAMEN_LAB & Constantes.AM)
        fabricaHC.registrar(Of ParaclinicoLaboratorioRR)(ConstantesHC.CODIGO_EXAMEN_LAB & Constantes.AF)
        fabricaHC.registrar(Of AtencionParaclinicoLaboratorio)(ConstantesHC.CODIGO_EXAMEN_LAB & Constantes.LB)
        ' instancias de resultados examenes
        fabricaHC.registrar(Of ExamenResultado)(ConstantesHC.CODIGO_IMAGENOLOGIA & Constantes.HC)
        fabricaHC.registrar(Of ExamenResultadoHemo)(ConstantesHC.CODIGO_IMAGENOLOGIA & Constantes.CODIGO_MENU_HEMO)
        fabricaHC.registrar(Of ExamenResultadoQuir)(ConstantesHC.CODIGO_IMAGENOLOGIA & Constantes.CODIGO_MENU_QUIR)
        fabricaHC.registrar(Of ExamenResultadoHosp)(ConstantesHC.CODIGO_IMAGENOLOGIA & Constantes.CODIGO_MENU_HOSP)
        fabricaHC.registrar(Of ExamenResultadoUrg)(ConstantesHC.CODIGO_IMAGENOLOGIA & Constantes.CODIGO_MENU_URG)
        fabricaHC.registrar(Of ExamenResultadoCExt)(ConstantesHC.CODIGO_IMAGENOLOGIA & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of ExamenResultadoR)(ConstantesHC.CODIGO_IMAGENOLOGIA & Constantes.AM)
        fabricaHC.registrar(Of ExamenResultadoRR)(ConstantesHC.CODIGO_IMAGENOLOGIA & Constantes.AF)
        fabricaHC.registrar(Of Documentacion)(ConstantesHC.CODIGO_IMAGENOLOGIA & Constantes.DC)
        fabricaHC.registrar(Of EstudioLaboratorio)(ConstantesHC.CODIGO_IMAGENOLOGIA & Constantes.LB)
        '--Instancia Imagenologia 
        fabricaHC.registrar(Of EstudioImagenologia)(ConstantesHC.CODIGO_ESTUDIO_IMAG & Constantes.HC)
        fabricaHC.registrar(Of EstudioImagenologiaHemo)(ConstantesHC.CODIGO_ESTUDIO_IMAG & Constantes.CODIGO_MENU_HEMO)
        fabricaHC.registrar(Of EstudioImagenologiaQuir)(ConstantesHC.CODIGO_ESTUDIO_IMAG & Constantes.CODIGO_MENU_QUIR)
        fabricaHC.registrar(Of EstudioImagenologiaHosp)(ConstantesHC.CODIGO_ESTUDIO_IMAG & Constantes.CODIGO_MENU_HOSP)
        fabricaHC.registrar(Of EstudioImagenologiaUrg)(ConstantesHC.CODIGO_ESTUDIO_IMAG & Constantes.CODIGO_MENU_URG)
        fabricaHC.registrar(Of EstudioImagenologiaCExt)(ConstantesHC.CODIGO_ESTUDIO_IMAG & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of EstudioImagenologiaR)(ConstantesHC.CODIGO_ESTUDIO_IMAG & Constantes.AM)
        fabricaHC.registrar(Of EstudioImagenologiaRR)(ConstantesHC.CODIGO_ESTUDIO_IMAG & Constantes.AF)
        fabricaHC.registrar(Of EstudioLaboratorio)(ConstantesHC.CODIGO_ESTUDIO_IMAG & Constantes.LB)
        '--- Solicitud Laboratorio
        fabricaHC.registrar(Of SolicitudLaboratorio)(ConstantesHC.CODIGO_SOLICITUD & Constantes.HC)
        fabricaHC.registrar(Of SolicitudLaboratorioR)(ConstantesHC.CODIGO_SOLICITUD & Constantes.AM)
        fabricaHC.registrar(Of SolicitudLaboratorioRR)(ConstantesHC.CODIGO_SOLICITUD & Constantes.AF)

        '--- facturacion asistencial
        fabricaHC.registrar(Of FacturaAtencionAsistencial)(ConstantesAsis.CODIGO_FACTURA_ASISTENCIAL & Constantes.HC)
        fabricaHC.registrar(Of FacturaAtencionAsistencialR)(ConstantesAsis.CODIGO_FACTURA_ASISTENCIAL & Constantes.AM)
        fabricaHC.registrar(Of FacturaAtencionAsistencialRR)(ConstantesAsis.CODIGO_FACTURA_ASISTENCIAL & Constantes.AF)

        '--- solicitud de laboratorio
        fabricaHC.registrar(Of SolicitudLaboratorio)(ConstantesHC.CODIGO_SOLICITUD_LAB & Constantes.HC)
        fabricaHC.registrar(Of SolicitudLaboratorioAtencion)(ConstantesHC.CODIGO_SOLICITUD_LAB & Constantes.CODIGO_MENU_IMAG)
        fabricaHC.registrar(Of SolicitudLaboratorioAtencion)(ConstantesHC.CODIGO_SOLICITUD_LAB & Constantes.CODIGO_MENU_CEXT)
        fabricaHC.registrar(Of SolicitudLaboratorioAtencion)(ConstantesHC.CODIGO_SOLICITUD_LAB & Constantes.CODIGO_MENU_LAB)
    End Sub

    Public Shared Sub cargarExamenParaclinico()

    End Sub
    Public Shared Function calcularGlasgow(gbReaccion As GroupBox) As Integer
        Dim vItem As RadioButton

        For Each vItem In gbReaccion.Controls
            If (TypeOf vItem Is RadioButton) And vItem.Checked = True Then
                Return CInt(vItem.Text.Substring(0, 1))

            End If
        Next
        Return 0
    End Function

    Public Shared Sub crearTablaSabana(ByRef dtHorasabana As DataTable)
        dtHorasabana = New DataTable
        dtHorasabana.Columns.Add("codigo", Type.GetType("System.Int32"))
        dtHorasabana.Columns.Add("0", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("1", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("2", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("3", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("4", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("5", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("6", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("7", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("8", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("9", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("10", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("11", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("12", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("13", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("14", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("15", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("16", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("17", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("18", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("19", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("20", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("21", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("22", Type.GetType("System.String"))
        dtHorasabana.Columns.Add("23", Type.GetType("System.String"))
    End Sub
    Public Shared Sub crearTablaSabanaGoteo(ByRef dtHorasabana As DataTable)

        dtHorasabana = New DataTable
        dtHorasabana.Columns.Add("codigo", Type.GetType("System.Int32"))
        dtHorasabana.Columns.Add("0", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("1", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("2", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("3", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("4", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("5", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("6", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("7", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("8", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("9", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("10", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("11", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("12", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("13", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("14", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("15", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("16", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("17", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("18", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("19", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("20", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("21", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("22", Type.GetType("System.Double"))
        dtHorasabana.Columns.Add("23", Type.GetType("System.Double"))
    End Sub
    Public Shared Sub enlaceDgvSabana(ByRef dgv As DataGridView, ByRef dt As DataTable)
        If IsNothing(dgv.DataSource) = True Then
            With dgv
                .ReadOnly = False
                .Columns.Item(0).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(0).DataPropertyName = "Codigo"
                .Columns.Item(0).Visible = False
                .Columns.Item(1).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(1).DataPropertyName = "Nombre"
                .Columns.Item(2).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(2).DataPropertyName = "0"
                .Columns.Item(2).ReadOnly = True
                .Columns.Item(3).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(3).DataPropertyName = "1"
                .Columns.Item(3).ReadOnly = True
                .Columns.Item(4).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(4).DataPropertyName = "2"
                .Columns.Item(4).ReadOnly = True
                .Columns.Item(5).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(5).DataPropertyName = "3"
                .Columns.Item(5).ReadOnly = True
                .Columns.Item(6).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(6).DataPropertyName = "4"
                .Columns.Item(6).ReadOnly = True
                .Columns.Item(7).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(7).DataPropertyName = "5"
                .Columns.Item(7).ReadOnly = True
                .Columns.Item(8).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(8).DataPropertyName = "6"
                .Columns.Item(8).ReadOnly = True
                .Columns.Item(9).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(9).DataPropertyName = "7"
                .Columns.Item(9).ReadOnly = True
                .Columns.Item(10).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(10).DataPropertyName = "8"
                .Columns.Item(10).ReadOnly = True
                .Columns.Item(11).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(11).DataPropertyName = "9"
                .Columns.Item(11).ReadOnly = True
                .Columns.Item(12).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(12).DataPropertyName = "10"
                .Columns.Item(12).ReadOnly = True
                .Columns.Item(13).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(13).DataPropertyName = "11"
                .Columns.Item(13).ReadOnly = True
                .Columns.Item(14).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(14).DataPropertyName = "12"
                .Columns.Item(14).ReadOnly = True
                .Columns.Item(15).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(15).DataPropertyName = "13"
                .Columns.Item(15).ReadOnly = True
                .Columns.Item(16).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(16).DataPropertyName = "14"
                .Columns.Item(16).ReadOnly = True
                .Columns.Item(17).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(17).DataPropertyName = "15"
                .Columns.Item(17).ReadOnly = True
                .Columns.Item(18).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(18).DataPropertyName = "16"
                .Columns.Item(18).ReadOnly = True
                .Columns.Item(19).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(19).DataPropertyName = "17"
                .Columns.Item(19).ReadOnly = True
                .Columns.Item(20).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(20).DataPropertyName = "18"
                .Columns.Item(20).ReadOnly = True
                .Columns.Item(21).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(21).DataPropertyName = "19"
                .Columns.Item(21).ReadOnly = True
                .Columns.Item(22).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(22).DataPropertyName = "20"
                .Columns.Item(22).ReadOnly = True
                .Columns.Item(23).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(23).DataPropertyName = "21"
                .Columns.Item(23).ReadOnly = True
                .Columns.Item(24).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(24).DataPropertyName = "22"
                .Columns.Item(24).ReadOnly = True
                .Columns.Item(25).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(25).DataPropertyName = "23"
                .Columns.Item(25).ReadOnly = True
            End With
            If dgv.ColumnCount = 28 Then
                dgv.Columns.Item(26).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv.Columns.Item(26).DataPropertyName = "CodigoOrden"
                dgv.Columns.Item(26).ReadOnly = True
                dgv.Columns.Item(27).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv.Columns.Item(27).DataPropertyName = "Tipo"
                dgv.Columns.Item(27).ReadOnly = True
            ElseIf dgv.ColumnCount = 27 Then
                dgv.Columns.Item(26).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv.Columns.Item(26).DataPropertyName = "ConsecutivoInfusion"
                dgv.Columns.Item(26).ReadOnly = True
            End If
        End If
        dgv.DataSource = dt
        dgv.Columns(dgv.Columns(0).Name).Visible = False

        dgv.DefaultCellStyle.Font = New Font(Constantes.TIPO_LETRA_ELEMENTO, 7)
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub

    Public Shared Sub colorFiLaDgv(ByRef dgv As DataGridView, indiceFila As Integer, color As Color)
        dgv.Rows(indiceFila).DefaultCellStyle.ForeColor = color
    End Sub

    Public Shared Sub enlaceDgvBalance(ByRef dgv As DataGridView, ByRef dt As DataTable)
        If IsNothing(dgv.DataSource) = True Then
            With dgv
                .ReadOnly = False
                .Columns.Item(0).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(0).DataPropertyName = "Codigo"
                .Columns.Item(1).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(1).DataPropertyName = "Nombre"
                .Columns.Item(2).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(2).DataPropertyName = "0 - 6"
                .Columns.Item(2).ReadOnly = True
                .Columns.Item(3).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(3).DataPropertyName = "6 - 12"
                .Columns.Item(3).ReadOnly = True
                .Columns.Item(4).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(4).DataPropertyName = "12 - 18"
                .Columns.Item(4).ReadOnly = True
                .Columns.Item(5).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Item(5).DataPropertyName = "18 - 0"
                .Columns.Item(5).ReadOnly = True
            End With
        End If
        dgv.DataSource = dt
        dgv.Columns(0).Visible = False
        dgv.DefaultCellStyle.Font = New Font(Constantes.TIPO_LETRA_ELEMENTO, 9)
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub

    Public Shared Sub filaDGVSoloLectura(ByRef dgv As DataGridView, pFila As Integer)
        For i = 0 To dgv.ColumnCount - 1
            dgv.Rows(pFila).Cells(i).ReadOnly = True
        Next
    End Sub

    Public Shared Function verificarFechaMaxMinSeleccionParametrosV(dtFecha As DateTimePicker, ByRef objSabana As ParametroVentilacion) As Boolean
        If dtFecha.Value.Date < CDate(objSabana.fechaIngreso).Date Then
            MsgBox(Mensajes.FECHA_MENOR_INGRESO, MsgBoxStyle.Exclamation)
            objSabana.dtHorasabana.Clear()
            Return False
        ElseIf dtFecha.Value.Date > CDate(objSabana.fechaEgreso).Date Then
            MsgBox(Mensajes.FECHA_MAYOR_EPICRISIS, MsgBoxStyle.Exclamation)
            objSabana.dtHorasabana.Clear()
            Return False
        End If
        Return True
    End Function
    Public Shared Function verificarFechaMaxMinSeleccionSabana(dtFecha As DateTimePicker, ByRef objSabana As SabanaAplicacionMed) As Boolean
        If dtFecha.Value.Date < CDate(objSabana.fechaIngreso).Date Then
            MsgBox(Mensajes.FECHA_MENOR_INGRESO, MsgBoxStyle.Exclamation)
            objSabana.dtHorasabana.Clear()
            objSabana.dtHoraSabanaSilverman.Clear()
            objSabana.dtHoraSabanaValoracion.Clear()
            objSabana.dtHoraSabanaIngreso.Clear()
            objSabana.dtHoraSabanaEgreso.Clear()
            objSabana.dtHoraSabanaBalance.Clear()
            objSabana.dtHoraSabanaMedicamento.Clear()
            objSabana.dtHoraSabanaNPT.Clear()
            objSabana.dtHorasabanaGoteo.Clear()
            Return False
        ElseIf dtFecha.Value.Date > CDate(objSabana.fechaEgreso).Date Then
            MsgBox(Mensajes.FECHA_MAYOR_EPICRISIS, MsgBoxStyle.Exclamation)
            objSabana.dtHorasabana.Clear()
            objSabana.dtHoraSabanaSilverman.Clear()
            objSabana.dtHoraSabanaValoracion.Clear()
            objSabana.dtHoraSabanaIngreso.Clear()
            objSabana.dtHoraSabanaEgreso.Clear()
            objSabana.dtHoraSabanaBalance.Clear()
            objSabana.dtHoraSabanaMedicamento.Clear()
            objSabana.dtHoraSabanaNPT.Clear()
            objSabana.dtHorasabanaGoteo.Clear()
            Return False
        End If
        Return True
    End Function

    Public Shared Sub cargarListBox(ByRef lista As ListBox, ByRef tabla As DataTable, display As String, value As String)

        Dim tabla2 As New DataTable
        tabla2 = tabla.Copy
        Dim drFila As DataRow = tabla2.NewRow()
        drFila.Item(value) = "-1"
        drFila.Item(1) = ""
        drFila.Item(display) = " - - - Seleccione - - - "
        tabla2.Rows.InsertAt(drFila, 0)
        lista.DataSource = tabla2
        lista.DisplayMember = display
        lista.ValueMember = value
        lista.SelectedIndex = 0
        lista.Enabled = True
    End Sub


    'Public Shared Sub eliminarInforme(modulo As String, activoAM As Boolean, activoAF As Boolean, nombrePDFHC As String, nombrePDFAM As String,
    '                                  nombrePDFAF As String, verificarAM As String, verificarAF As String)
    '    If modulo = Constantes.HC Then
    '        ftp_reportes.eliminarArchivo(nombrePDFHC)
    '        If activoAM = True And General.getEstadoVF(verificarAM) = False Then
    '            ftp_reportes.eliminarArchivo(nombrePDFAM)
    '        End If
    '        If activoAF = True And General.getEstadoVF(verificarAF) = False Then
    '            ftp_reportes.eliminarArchivo(nombrePDFAF)
    '        End If
    '    ElseIf modulo = Constantes.AM Then
    '        ftp_reportes.eliminarArchivo(nombrePDFAM)
    '        If activoAF = True And General.getEstadoVF(verificarAF) = False Then
    '            ftp_reportes.eliminarArchivo(nombrePDFAF)
    '        End If
    '    ElseIf modulo = Constantes.AF Then
    '        ftp_reportes.eliminarArchivo(nombrePDFAF)
    '    End If

    'End Sub

    Public Shared Function validarFecha(modulo As String, nRegistro As Integer, fechaSeleccionada As Object, fechaIngreso As DateTime, ByVal permiso As List(Of String), Optional ByVal opcionEpicrisis As Boolean = False)
        Dim fecha, fecha2 As DateTime

        fecha = Format(CDate(fechaSeleccionada.VALUE), Constantes.FORMATO_FECHA_HORA_GEN)
        If opcionEpicrisis = False Then
            Dim dtResultado As New DataTable
            Dim consulta As String = ""
            If modulo = Constantes.HC Then
                consulta = Consultas.FECHA_EPICRISIS_VERIFICAR
            ElseIf modulo = Constantes.AM Then
                consulta = Consultas.FECHA_EPICRISIS_VERIFICARR
            ElseIf modulo = Constantes.AF Then
                consulta = Consultas.FECHA_EPICRISIS_VERIFICARRR
            End If
            General.llenarTablaYdgv(consulta & nRegistro, dtResultado)
            If dtResultado.Rows.Count > 0 Then

                fecha2 = Format(CDate(dtResultado.Rows(0).Item("Fecha_Egreso")), Constantes.FORMATO_FECHA_HORA_GEN)
                If fecha > fecha2 Then
                    Return Mensajes.FECHA_MAYOR_EPICRISIS & fecha2
                End If
            End If
        End If

        fecha2 = Format(fechaIngreso, Constantes.FORMATO_FECHA_GEN)
        If CDate(fecha) < CDate(fecha2) Then
            Return Mensajes.FECHA_MENOR_INGRESO
        End If

        Dim diferenciaHoras As Integer

        fecha2 = Format(CDate(Funciones.Fecha(Constantes.FORMATO_FECHA_HORA_NUMERICA)), Constantes.FORMATO_FECHA_HORA_GEN)
        diferenciaHoras = Math.Abs(DateDiff(DateInterval.Hour, fecha, fecha2))

        If modulo = Constantes.HC AndAlso Not SesionActual.tienePermiso(permiso(0)) AndAlso
            (diferenciaHoras > ConstantesHC.CANTIDAD_INTERVALO_HORA_MAXIMA) Then
            Return Mensajes.INTERVALO_MAYOR_HORAS
        End If

        Dim params As New List(Of String)
        params.Add(nRegistro)
        params.Add(Format(fecha, Constantes.FORMATO_FECHA_HORA_GEN))
        params.Add(ConstantesHC.CANTIDAD_INTERMEDIO_HORA_MAXIMA)
        If modulo = Constantes.HC AndAlso Not SesionActual.tienePermiso(permiso(1)) Then
            If General.getEstadoVF(ConsultasHC.ORDEN_MEDICA_INTERMEDIA_VERIFICAR, params) = 1 Then
                Return Mensajes.INTERMEDIO_MAYOR_HORAS
            End If
        End If

        Return ""
    End Function

#Region "Sabana"

    Public Shared Sub RegistroEditar(ByRef dtg As DataGridView, objSabana As Sabana)
        If objSabana.dtHorasabana.Rows.Count > 0 Then
            Dim indiceColumna As Integer = dtg.CurrentCell.ColumnIndex
            Dim fechaServidor As DateTime = Funciones.Fecha(Constantes.FORMATO_FECHA_HORA_NUMERICA)
            Dim horaActual As Integer = Format(fechaServidor, Constantes.FORMATO_HORA_CADENA)
            Dim horaIngreso As Integer = Format(objSabana.fechaIngreso, Constantes.FORMATO_HORA_CADENA)
            Dim horaEgreso As Integer = Format(objSabana.fechaEgreso, Constantes.FORMATO_HORA_CADENA)


            If CDate(objSabana.fechaReal).Date <> CDate(Funciones.Fecha(Constantes.FORMATO_FECHA_HORA_NUMERICA)).Date Then
                If CDate(objSabana.fechaReal).Date = (objSabana.fechaIngreso).Date AndAlso CInt(dtg.DataSource.Columns(indiceColumna).ColumnName) < horaIngreso Then
                    MsgBox(Mensajes.HORA_MENOR_INGRESO, MsgBoxStyle.Exclamation)
                    Exit Sub
                ElseIf CDate(objSabana.fechaReal).Date = (objSabana.fechaEgreso).Date AndAlso CInt(dtg.DataSource.Columns(indiceColumna).ColumnName) > horaEgreso Then
                    MsgBox(Mensajes.HORA_MAYOR_EPICRISIS, MsgBoxStyle.Exclamation)
                    Exit Sub
                End If
                dtg.Columns(indiceColumna).DefaultCellStyle.BackColor = Color.LightSkyBlue
                dtg.Columns(indiceColumna).ReadOnly = False
            End If
            If CDate(objSabana.fechaReal).Date = CDate(Funciones.Fecha(Constantes.FORMATO_FECHA_HORA_NUMERICA)).Date Then
                If CInt(dtg.DataSource.Columns(indiceColumna).ColumnName) > horaActual Then
                    MsgBox(Mensajes.HORA_MAYOR_EPICRISIS, MsgBoxStyle.Exclamation)
                    Exit Sub
                ElseIf CDate(objSabana.fechaReal).Date = (objSabana.fechaIngreso).Date AndAlso CInt(dtg.DataSource.Columns(indiceColumna).ColumnName) < horaIngreso Then
                    MsgBox(Mensajes.HORA_MENOR_INGRESO, MsgBoxStyle.Exclamation)
                    Exit Sub
                ElseIf CDate(objSabana.fechaReal).Date = (objSabana.fechaEgreso).Date AndAlso CInt(dtg.DataSource.Columns(indiceColumna).ColumnName) > horaEgreso Then
                    MsgBox(Mensajes.HORA_MAYOR_EPICRISIS, MsgBoxStyle.Exclamation)
                    Exit Sub
                End If
                dtg.Columns(indiceColumna).DefaultCellStyle.BackColor = Color.LightSkyBlue
                dtg.Columns(indiceColumna).ReadOnly = False
            End If
        End If
    End Sub
    Public Shared Sub RegistroAplicar(ByRef dtg As DataGridView, objSabana As Sabana)
        Dim indiceColumna As Integer = dtg.CurrentCell.ColumnIndex
        Dim indiceFila As Integer = dtg.CurrentCell.RowIndex

        If dtg.Rows(indiceFila).Cells(indiceColumna).Value.ToString = ConstantesHC.IDENTIFICADOR_SABANA_NO_APLICADO Then
            dtg.Rows(indiceFila).Cells(indiceColumna).Value = objSabana.inicialesUsuario
            dtg.Rows(indiceFila).Cells(indiceColumna).Style.BackColor = Color.FromArgb(192, 255, 192)
        End If
    End Sub
    Public Shared Sub RegistroDesaplicar(ByRef dtg As DataGridView, objSabana As Sabana)
        Dim indiceColumna As Integer = dtg.CurrentCell.ColumnIndex
        Dim indiceFila As Integer = dtg.CurrentCell.RowIndex

        If dtg.Rows(indiceFila).Cells(indiceColumna).Style.BackColor = Color.FromArgb(192, 255, 192) Then
            dtg.Rows(indiceFila).Cells(indiceColumna).Value = ConstantesHC.IDENTIFICADOR_SABANA_NO_APLICADO
            dtg.Rows(indiceFila).Cells(indiceColumna).Style.BackColor = Color.LightYellow
        End If
    End Sub

    Public Shared Sub colorRegistroCargar(ByRef dtg As DataGridView, ByRef objSabana As Sabana, idColumna As String)
        Dim fechaServidor As DateTime = Funciones.Fecha(Constantes.FORMATO_FECHA_HORA_NUMERICA)
        Dim horaActual As Integer = Format(fechaServidor, Constantes.FORMATO_HORA_CADENA)
        Dim horaIngreso As Integer = Format(objSabana.fechaIngreso, Constantes.FORMATO_HORA_CADENA)
        Dim horaEgreso As Integer = Format(objSabana.fechaEgreso, Constantes.FORMATO_HORA_CADENA)

        For indiceCol = 0 To 23
            verificarRegistroHoraSabana(dtg, dtg.DataSource, indiceCol, idColumna)
            dtg.Columns(idColumna & indiceCol).ReadOnly = True
        Next
        colorRegistroAdmisionEgreso(dtg, objSabana, idColumna, horaIngreso, horaEgreso)
    End Sub
    Public Shared Sub verificarRegistroHoraSabana(ByRef dtg As DataGridView, dtTabla As DataTable, Columna As String, idColumna As String)
        Dim tieneDatos As Boolean = False
        For I = 0 To dtTabla.Rows.Count - 1
            If Not String.IsNullOrEmpty(dtTabla.Rows(I).Item(Columna).ToString) AndAlso dtTabla.Rows(I).Item(0) <> ConstantesHC.CODIGO_TOTAL_INGRESO AndAlso dtTabla.Rows(I).Item(0) <> ConstantesHC.CODIGO_TOTAL_EGRESO Then
                tieneDatos = True
                Exit For
            End If
        Next
        If tieneDatos = True Then
            dtg.Columns(idColumna & Columna & "").DefaultCellStyle.BackColor = Color.FromArgb(192, 255, 192) ' VERDE
        Else
            dtg.Columns(idColumna & Columna & "").DefaultCellStyle.BackColor = Color.LightYellow
        End If
    End Sub

    Private Shared Sub colorRegistroAdmisionEgreso(ByRef dtg As DataGridView, ByRef objSabana As Sabana, idColumna As String, horaIngreso As Integer, horaEgreso As Integer)
        objSabana.numeroHorastranscurridas = 0
        If CDate(objSabana.fechaReal).Date = (objSabana.fechaIngreso).Date Then
            For hora = 0 To horaIngreso - 1
                dtg.Columns(idColumna & hora).DefaultCellStyle.BackColor = Color.LightGray
                For fila = 0 To dtg.RowCount - 1
                    dtg.Rows(fila).Cells(idColumna & hora).Value = Nothing
                Next
            Next
            objSabana.numeroHorastranscurridas = 24 - horaIngreso
        End If
        If CDate(objSabana.fechaReal).Date = (objSabana.fechaEgreso).Date Then
            verificarRegistroHoraSabana(dtg, dtg.DataSource, horaEgreso, idColumna)
            If dtg.Columns(idColumna & horaEgreso).DefaultCellStyle.BackColor = Color.LightYellow Then
                dtg.Columns(idColumna & horaEgreso).DefaultCellStyle.BackColor = Color.FromArgb(255, 192, 192) ' ROJO
            End If
            dtg.Columns(idColumna & horaEgreso).ReadOnly = False
            For hEgreso = horaEgreso + 1 To 23
                dtg.Columns(idColumna & hEgreso).DefaultCellStyle.BackColor = Color.LightGray
                For fila = 0 To dtg.RowCount - 1
                    dtg.Rows(fila).Cells(idColumna & hEgreso).Value = Nothing
                Next
            Next
            If objSabana.numeroHorastranscurridas > 0 Then
                objSabana.numeroHorastranscurridas = horaEgreso - horaIngreso + 1
            Else
                objSabana.numeroHorastranscurridas = horaEgreso + 1
            End If
        End If
        If Not CDate(objSabana.fechaReal).Date = (objSabana.fechaEgreso).Date AndAlso Not CDate(objSabana.fechaReal).Date = (objSabana.fechaIngreso).Date Then
            objSabana.numeroHorastranscurridas = 24
        End If
    End Sub
    Private Shared Sub colorRegistroAdmisionEgresoAplicacionMed(ByRef dtg As DataGridView, ByRef objSabana As Sabana, idColumna As String, horaIngreso As Integer, horaEgreso As Integer)
        If CDate(objSabana.fechaReal).Date = (objSabana.fechaIngreso).Date Then
            For hora = 0 To horaIngreso - 1

                'dtg.Columns(idColumna & hora).DefaultCellStyle.BackColor = Color.LightGray
                For fila = 0 To dtg.RowCount - 1
                    dtg.Rows(fila).Cells(idColumna & hora).Style.BackColor = Color.LightGray
                    If dtg.Name.ToLower.Contains("goteo") Then
                        dtg.Rows(fila).Cells(idColumna & hora).Value = 0
                    Else
                        dtg.Rows(fila).Cells(idColumna & hora).Value = ""
                    End If
                Next
            Next
        End If
        If CDate(objSabana.fechaReal).Date = (objSabana.fechaEgreso).Date Then
            verificarRegistroHoraSabanaAplicacionMed(dtg, dtg.DataSource, horaEgreso, idColumna)
            If dtg.Columns(idColumna & horaEgreso).DefaultCellStyle.BackColor = Color.LightYellow Then
                dtg.Columns(idColumna & horaEgreso).DefaultCellStyle.BackColor = Color.FromArgb(255, 192, 192) ' ROJO

            End If
            dtg.Columns(idColumna & horaEgreso).ReadOnly = False
            For hEgreso = horaEgreso + 1 To 23

                dtg.Columns(idColumna & hEgreso).ReadOnly = True
                For fila = 0 To dtg.RowCount - 1
                    If dtg.Name.ToLower.Contains("goteo") Then
                        dtg.Rows(fila).Cells(idColumna & hEgreso).Value = 0
                    Else
                        dtg.Rows(fila).Cells(idColumna & hEgreso).Value = ""
                    End If
                    dtg.Rows(fila).Cells(idColumna & hEgreso).Style.BackColor = Color.LightGray
                Next
            Next
        End If
    End Sub

    Public Shared Sub colorRegistroCargarAplicacionMed(ByRef dtg As DataGridView, ByRef objSabana As Sabana, idColumna As String)
        Dim fechaServidor As DateTime = Funciones.Fecha(Constantes.FORMATO_FECHA_HORA_NUMERICA)
        Dim horaActual As Integer = Format(fechaServidor, Constantes.FORMATO_HORA_CADENA)
        Dim horaIngreso As Integer = Format(objSabana.fechaIngreso, Constantes.FORMATO_HORA_CADENA)
        Dim horaEgreso As Integer = Format(objSabana.fechaEgreso, Constantes.FORMATO_HORA_CADENA)

        For indiceCol = 0 To 23
            verificarRegistroHoraSabanaAplicacionMed(dtg, dtg.DataSource, indiceCol, idColumna)
        Next
        If dtg.RowCount > 0 Then
            colorRegistroAdmisionEgresoAplicacionMed(dtg, objSabana, idColumna, horaIngreso, horaEgreso)
        End If
    End Sub
    Public Shared Sub colorRegistroCargarAplicacionGoteo(ByRef dtg As DataGridView, ByRef objSabana As Sabana, idColumna As String)
        Dim fechaServidor As DateTime = Funciones.Fecha(Constantes.FORMATO_FECHA_HORA_NUMERICA)
        Dim horaActual As Integer = Format(fechaServidor, Constantes.FORMATO_HORA_CADENA)
        Dim horaIngreso As Integer = Format(objSabana.fechaIngreso, Constantes.FORMATO_HORA_CADENA)
        Dim horaEgreso As Integer = Format(objSabana.fechaEgreso, Constantes.FORMATO_HORA_CADENA)

        For indiceCol = 0 To 23
            dtg.Columns(idColumna & indiceCol).ReadOnly = True
            verificarRegistroHoraSabanaAplicacionGoteo(dtg, dtg.DataSource, indiceCol, idColumna)

        Next
        If dtg.RowCount > 0 Then
            colorRegistroAdmisionEgresoAplicacionMed(dtg, objSabana, idColumna, horaIngreso, horaEgreso)
        End If
    End Sub
    Public Shared Sub verificarRegistroHoraSabanaAplicacionMed(ByRef dtg As DataGridView, dtTabla As DataTable, Columna As String, idColumna As String)
        For I = 0 To dtTabla.Rows.Count - 1
            Select Case dtTabla.Rows(I).Item(Columna).ToString
                Case ""
                    dtg.Rows(I).Cells(idColumna & Columna & "").Style.BackColor = Color.White
                Case ConstantesHC.IDENTIFICADOR_SABANA_NO_APLICADO
                    dtg.Rows(I).Cells(idColumna & Columna & "").Style.BackColor = Color.LightYellow
                Case ConstantesHC.IDENTIFICADOR_SABANA_MODIFICADO
                    dtg.Rows(I).Cells(idColumna & Columna & "").Style.BackColor = Color.LightSteelBlue
                Case ConstantesHC.IDENTIFICADOR_SABANA_SUSPENDIDO
                    dtg.Rows(I).Cells(idColumna & Columna & "").Style.BackColor = Color.FromArgb(255, 192, 192)
                Case Else
                    dtg.Rows(I).Cells(idColumna & Columna & "").Style.BackColor = Color.FromArgb(192, 255, 192)

            End Select
        Next
    End Sub

    Public Shared Sub verificarRegistroHoraSabanaAplicacionGoteo(ByRef dtg As DataGridView, dtTabla As DataTable, Columna As String, idColumna As String)
        For I = 0 To dtTabla.Rows.Count - 1
            Select Case dtTabla.Rows(I).Item(Columna).ToString
                Case ""
                    dtg.Rows(I).Cells(idColumna & Columna & "").Style.BackColor = Color.White
                Case ConstantesHC.IDENTIFICADOR_SABANA_GOTEO_NO_APLICADO
                    dtg.Rows(I).Cells(idColumna & Columna & "").Style.BackColor = Color.LightYellow
                    dtg.Rows(I).Cells(idColumna & Columna & "").ReadOnly = False
                Case ConstantesHC.IDENTIFICADOR_SABANA_GOTEO_MODIFICADO
                    dtg.Rows(I).Cells(idColumna & Columna & "").Style.BackColor = Color.LightSteelBlue
                Case ConstantesHC.IDENTIFICADOR_SABANA_GOTEO_SUSPENDIDO
                    dtg.Rows(I).Cells(idColumna & Columna & "").Style.BackColor = Color.FromArgb(255, 192, 192)
                Case Else
                    dtg.Rows(I).Cells(idColumna & Columna & "").Style.BackColor = Color.FromArgb(192, 255, 192)
                    dtg.Rows(I).Cells(idColumna & Columna & "").ReadOnly = False
            End Select
        Next
    End Sub

    Friend Shared Function consultaVisado(nRegistro As Integer, Optional opcionConsultar As Integer = 0) As Boolean
        Dim params As New List(Of String)
        params.Add(nRegistro)
        params.Add(opcionConsultar)
        params.Add(SesionActual.idUsuario)

        If General.getEstadoVF(ConsultasHC.CAMBIAR_ESTADO_VISADO, params) Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

End Class
