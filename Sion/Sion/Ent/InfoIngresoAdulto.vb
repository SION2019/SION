Public Class InfoIngresoAdulto
    Inherits InfoIngreso
    Public Property antecedentes As String
    Public Property antecedentesQuirurgicos As String
    Public Property antecedentesTransfusionales As String
    Public Property antecedentesAlergicos As String
    Public Property antecedentesTraumaticos As String
    Public Property antecedentesToxicos As String
    Public Property antecedentesFamiliares As String
    Public Property revisionSistema As String
    Public Property paraclinico As String


    Public Sub New()
        sqlDetalleCarga = Consultas.INFO_INGRESO_CARGAR
        sqlDetalleImpresionCarga = Consultas.INFOIMPRESION_DIAGNOSTICA_CARGAR
        sqlDetalleRemisionCarga = Consultas.INFOREMISION_DIAGNOSTICA_CARGAR
        nombreReporte = ConstantesHC.NOMBRE_PDF_INGRESO_MEDICO
        objReporte = New rptInfoIngresoAdulto
        vistaReporte = "VISTA_H_INGRESO"
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
        EstadoAtencion = dt.Rows(0).Item("Descripcion_Estado_Atencion").ToString
        Autorizacion = dt.Rows(0).Item("N_Autorizacion").ToString
        Institucion = dt.Rows(0).Item("Descripcion_Institucion").ToString
        ViaIngreso = dt.Rows(0).Item("Descripcion_Via").ToString
        CausaExterna = dt.Rows(0).Item("Descripcion_Causa").ToString
        Cama = dt.Rows(0).Item("Descripcion_Cama").ToString
        peso = IIf(String.IsNullOrEmpty(dt.Rows(0).Item("Peso").ToString), 0, dt.Rows(0).Item("Peso").ToString)
        pesoUltimo = dt.Rows(0).Item("UltimoPeso").ToString
        motivo = dt.Rows(0).Item("motivo_ingreso").ToString
        antecedentes = dt.Rows(0).Item("Ant_Medico").ToString
        antecedentesQuirurgicos = dt.Rows(0).Item("Ant_Quirurgico").ToString
        antecedentesTraumaticos = dt.Rows(0).Item("Ant_Traumaticos").ToString
        antecedentesTransfusionales = dt.Rows(0).Item("Ant_Transfuncionales").ToString
        antecedentesAlergicos = dt.Rows(0).Item("Ant_Alergicos").ToString
        antecedentesToxicos = dt.Rows(0).Item("Ant_Toxicos").ToString
        antecedentesFamiliares = dt.Rows(0).Item("Ant_Familiares").ToString
        revisionSistema = dt.Rows(0).Item("Revision_Sistema").ToString
        signosVitales = dt.Rows(0).Item("Sig_Vitales").ToString
        cabezaYCuello = dt.Rows(0).Item("Cab_Cuello").ToString
        torax = dt.Rows(0).Item("Torax").ToString
        cardio = dt.Rows(0).Item("Card_Pulm").ToString
        abdomen = dt.Rows(0).Item("Abdom").ToString
        gentoUrinario = dt.Rows(0).Item("Genturinario").ToString
        extremidades = dt.Rows(0).Item("Extrem").ToString
        sistemaNervioso = dt.Rows(0).Item("S_Nerv_Central").ToString
        paraclinico = dt.Rows(0).Item("Paraclinicos").ToString
        analisis = dt.Rows(0).Item("Analisis").ToString
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
        HistoriaClinicaBLL.guardarInfoIngreso(Me)
    End Sub




End Class
