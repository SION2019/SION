Public Class FormatoIngreso
    Public Property usuario As Integer
    Public Property codigo As String
    Public Property anamnesis As String
    Public Property medico As String
    Public Property quirurgico As String
    Public Property traumatico As String
    Public Property alergico As String
    Public Property toxico As String
    Public Property antFamiliares As String
    Public Property revisionSistema As String
    Public Property signoVitales As String
    Public Property cabezaCuello As String
    Public Property torax As String
    Public Property cardioPulmonar As String
    Public Property abdomen As String
    Public Property gestionUrinaria As String
    Public Property extremidades As String
    Public Property nervioCentral As String
    Public Property paraclinico As String
    Public Property tipoParto As String
    Public Property rubtura As String
    Public Property inducionParto As String
    Public Property apgar1 As String
    Public Property apgar5 As String
    Public Property reanimacionNacer As String
    Public Property edadMadre As String
    Public Property edadGestacional As String
    Public Property fum As String
    Public Property obstetrico As String
    Public Property hemograsificacionMadre As String
    Public Property hemograsificacionPadre As String
    Public Property controlPrenatal As String
    Public Property medDuranteEmbarazo As String
    Public Property habitos As String
    Public Property infeccionEmb As String
    Public Property diabGestional As String
    Public Property diabMellitud As String
    Public Property hiperGestional As String
    Public Property preclampcia As String
    Public Property vacunacion As String
    Public Property tiroidea As String
    Public Property torch As String
    Public Property hemogracificacionNcido As String
    Public Property VDRL As String
    Public Property TSH As String
    Public Property glucometria As String
    Public Property generales As String
    Public Property pronostico As String
    Public Property analisis As String
    Public Property transfuncionales As String
    Public Property caracteristicaLiquido As String
    Public Property codigoArea As String
    Public Property codigoGenero As String
    Public Property dtDiagPrincipales As DataTable
    Public Property dtDiagImpresion As DataTable
    Public Property sqlGuardarFormato As String
    Public Property sqldiagnosticoFormato As String
    Public Property sqlBuscarRegistro As String
    Public Property sqlCargarRegistro As String
    Public Property sqlCargarDiagnoticoPrincipal As String
    Public Property sqlCargarDiagnoticoImpresion As String
    Public Property sqlCargarParaclinicos As String
    Public Property sqlCargarMedicamentos As String
    Public Property dtParaclinico As DataTable
    Public Property dtMedicamento As DataTable
    Public Sub New()
        dtDiagPrincipales = New DataTable
        dtDiagPrincipales.Columns.Add("Código", Type.GetType("System.String"))
        dtDiagPrincipales.Columns.Add("Descripción", Type.GetType("System.String"))
        dtDiagPrincipales.Columns.Add("Estado", Type.GetType("System.String"))

        dtDiagImpresion = New DataTable
        dtDiagImpresion.Columns.Add("Código", Type.GetType("System.String"))
        dtDiagImpresion.Columns.Add("Descripción", Type.GetType("System.String"))
        dtDiagImpresion.Columns.Add("Estado", Type.GetType("System.String"))
        dtDiagImpresion.Columns.Add("H.C", Type.GetType("System.Boolean")).DefaultValue = 0

        dtParaclinico = New DataTable
        dtParaclinico.Columns.Add("Codigo", Type.GetType("System.String"))
        dtParaclinico.Columns.Add("Procedimiento", Type.GetType("System.String"))
        dtParaclinico.Columns.Add("Cant", Type.GetType("System.Int32")).DefaultValue = 1
        dtParaclinico.Columns.Add("Estado", Type.GetType("System.String"))
        dtParaclinico.Columns.Add("H.C", Type.GetType("System.Boolean")).DefaultValue = 0

        dtMedicamento = New DataTable
        dtMedicamento.Columns.Add("Codigo", Type.GetType("System.String"))
        dtMedicamento.Columns.Add("Medicamento", Type.GetType("System.String"))
        dtMedicamento.Columns.Add("Estado", Type.GetType("System.String"))
        dtMedicamento.Columns.Add("H.C", Type.GetType("System.Boolean")).DefaultValue = 0

        sqldiagnosticoFormato = Consultas.CONSULTAR_DIAGNOSTICO_FORMATO
        sqlGuardarFormato = Consultas.FORMATO_INGRESO_CREAR
        sqlBuscarRegistro = Consultas.FORMATO_INGRESO_BUSCAR
        sqlCargarRegistro = Consultas.FORMATO_INGRESO_CARGAR
        sqlCargarDiagnoticoPrincipal = Consultas.FORMATO_INGRESO_DIAGNOSTICO_PRINCIPAL
        sqlCargarDiagnoticoImpresion = Consultas.FORMATO_INGRESO_DIAGNOSTICO_IMPRESION
        sqlCargarParaclinicos = Consultas.FORMATO_INGRESO_PARACLINICO
        sqlCargarMedicamentos = Consultas.FORMATO_INGRESO_MEDICAMENTO
    End Sub

End Class
