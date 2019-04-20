Public Class InfoIngresoNeonato
    Inherits InfoIngreso
    Public Property generales As String
    Public Property tParto As String
    Public Property tRupturaM As String
    Public Property induccionParto As String
    Public Property caracteristicasLiquidas As String
    Public Property apgar1 As String
    Public Property apgar5 As String
    Public Property reanimacion As String
    Public Property edadMadre As String
    Public Property antecedentesObstetricos As String
    Public Property fumador As String
    Public Property edadGestacional As String
    Public Property hemoclasificacionMadre As String
    Public Property hemoclasificacionPadre As String
    Public Property control As String
    Public Property medicamentos As String
    Public Property habito As String
    Public Property infeccion As String
    Public Property diabeteG As String
    Public Property diabeteM As String
    Public Property hipertencion As String
    Public Property preclampcia As String
    Public Property enfermedad As String
    Public Property vacunacion As String
    Public Property torch As String
    Public Property hemoclasificacion As String
    Public Property vdrl As String
    Public Property tsh As String
    Public Property glucometria As String

    Public Sub New()
        sqlDetalleCarga = Consultas.INFO_INGRESON_CARGAR
        sqlDetalleImpresionCarga = Consultas.INFOIMPRESION_DIAGNOSTICA_CARGAR
        sqlDetalleRemisionCarga = Consultas.INFOREMISION_DIAGNOSTICA_CARGAR
        nombreReporte = ConstantesHC.NOMBRE_PDF_INGRESO_MEDICO
        objReporte = New rptInfoIngresoNeonato
        vistaReporte = "VISTA_H_INGRESO_NEONATAL"
        moduloReporte = Constantes.REPORTE_HC
    End Sub
    Public Overrides Sub cargarDetalle()
        elementoAEliminar.Clear()
        cargarDetallePaciente()
        cargarDetalleImpresion()
        cargarDetalleRemision()
    End Sub
    Public Sub cargarDetallePaciente()
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(nRegistro)
        General.llenarTabla(sqlDetalleCarga, params, dt)
        estadoAtencion = dt.Rows(0).Item("Descripcion_Estado_Atencion").ToString
        autorizacion = dt.Rows(0).Item("N_Autorizacion").ToString
        institucion = dt.Rows(0).Item("Descripcion_Institucion").ToString
        viaIngreso = dt.Rows(0).Item("Descripcion_Via").ToString
        causaExterna = dt.Rows(0).Item("Descripcion_Causa").ToString
        cama = dt.Rows(0).Item("Descripcion_Cama").ToString
        peso = dt.Rows(0).Item("Peso").ToString
        pesoUltimo = dt.Rows(0).Item("UltimoPeso").ToString
        motivo = dt.Rows(0).Item("motivo_ingreso").ToString
        tParto = dt.Rows(0).Item("t_parto").ToString
        tRupturaM = dt.Rows(0).Item("t_rupt_menbrana").ToString
        induccionParto = dt.Rows(0).Item("indu_parto").ToString
        caracteristicasLiquidas = dt.Rows(0).Item("caract_liquidas").ToString
        apgar1 = dt.Rows(0).Item("apgar1").ToString
        apgar5 = dt.Rows(0).Item("apgar5").ToString
        reanimacion = dt.Rows(0).Item("reanim_nacer").ToString
        edadMadre = dt.Rows(0).Item("edad_madre").ToString
        edadGestacional = dt.Rows(0).Item("edad_gesta").ToString
        fumador = dt.Rows(0).Item("fum").ToString
        antecedentesObstetricos = dt.Rows(0).Item("ant_obstetrico").ToString
        hemoclasificacionMadre = dt.Rows(0).Item("hemoc_maternal").ToString
        hemoclasificacionPadre = dt.Rows(0).Item("hemoc_paternal").ToString
        control = dt.Rows(0).Item("control_prenatal").ToString
        medicamentos = dt.Rows(0).Item("medc_durat_embarazo").ToString
        habito = dt.Rows(0).Item("habito").ToString
        infeccion = dt.Rows(0).Item("infecc_embarazo").ToString
        diabeteG = dt.Rows(0).Item("diabete_gest").ToString
        diabeteM = dt.Rows(0).Item("diabete_mellitus").ToString
        hipertencion = dt.Rows(0).Item("hipertenc_gesta").ToString
        preclampcia = dt.Rows(0).Item("preclampcia").ToString
        enfermedad = dt.Rows(0).Item("enferm_tiroidea").ToString
        vacunacion = dt.Rows(0).Item("vacunacion").ToString
        torch = dt.Rows(0).Item("torch").ToString
        hemoclasificacion = dt.Rows(0).Item("hemoclasificacion").ToString
        vdrl = dt.Rows(0).Item("vdrl").ToString
        tsh = dt.Rows(0).Item("tsh").ToString
        glucometria = dt.Rows(0).Item("glucometria").ToString
        generales = dt.Rows(0).Item("Generales").ToString
        signosVitales = dt.Rows(0).Item("Sig_Vitales").ToString
        cabezaYCuello = dt.Rows(0).Item("Cab_Cuello").ToString
        torax = dt.Rows(0).Item("Torax").ToString
        cardio = dt.Rows(0).Item("Card_Pulm").ToString
        abdomen = dt.Rows(0).Item("Abdom").ToString
        gentoUrinario = dt.Rows(0).Item("Genturinario").ToString
        extremidades = dt.Rows(0).Item("Extrem").ToString
        sistemaNervioso = dt.Rows(0).Item("S_Nerv_Central").ToString
        analisis = dt.Rows(0).Item("Analiisis").ToString
        pronostico = dt.Rows(0).Item("Pronosticos").ToString
        usuarioNombre = dt.Rows(0).Item("Usuario")
        usuarioReal = dt.Rows(0).Item("usuarioInforme")
        codEps = dt.Rows(0).Item("Id_Tercero_EPS")
        usuarioCreacion = dt.Rows(0).Item("usuarioCreacion")
    End Sub
    Public Sub cargarDetalleImpresion()
        Dim params As New List(Of String)
        params.Add(nRegistro)
        General.llenarTabla(sqlDetalleImpresionCarga, params, dtDiagImpresion)
    End Sub

    Public Sub cargarDetalleRemision()
        Dim params As New List(Of String)
        params.Add(nRegistro)
        General.llenarTabla(sqlDetalleRemisionCarga, params, dtDiagRemision)
    End Sub


    Public Overrides Sub guardarDetalle()
        HistoriaClinicaBLL.guardarInfoIngresoN(Me)
    End Sub

End Class
