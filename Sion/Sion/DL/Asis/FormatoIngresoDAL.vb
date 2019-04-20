Public Class FormatoIngresoDAL
    Public Shared Function guardarFormatoIngreso(objFormato As FormatoIngreso)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = objFormato.sqlGuardarFormato
                    comando.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objFormato.codigo
                    comando.Parameters.Add(New SqlParameter("@Codigo_Servicio", SqlDbType.Int)).Value = objFormato.codigoArea
                    comando.Parameters.Add(New SqlParameter("@Codigo_Genero", SqlDbType.Int)).Value = objFormato.codigoGenero
                    comando.Parameters.Add(New SqlParameter("@Anamnesis", SqlDbType.NVarChar)).Value = objFormato.anamnesis
                    comando.Parameters.Add(New SqlParameter("@Medico", SqlDbType.NVarChar)).Value = objFormato.medico
                    comando.Parameters.Add(New SqlParameter("@Quirurgico", SqlDbType.NVarChar)).Value = objFormato.quirurgico
                    comando.Parameters.Add(New SqlParameter("@Traumatico", SqlDbType.NVarChar)).Value = objFormato.traumatico
                    comando.Parameters.Add(New SqlParameter("@Transfucionales", SqlDbType.NVarChar)).Value = objFormato.transfuncionales
                    comando.Parameters.Add(New SqlParameter("@Alergico", SqlDbType.NVarChar)).Value = objFormato.alergico
                    comando.Parameters.Add(New SqlParameter("@Toxico", SqlDbType.NVarChar)).Value = objFormato.toxico
                    comando.Parameters.Add(New SqlParameter("@Ant_Familiares", SqlDbType.NVarChar)).Value = objFormato.antFamiliares
                    comando.Parameters.Add(New SqlParameter("@Revision_Sistema", SqlDbType.NVarChar)).Value = objFormato.revisionSistema
                    comando.Parameters.Add(New SqlParameter("@Signo_Vitales", SqlDbType.NVarChar)).Value = objFormato.signoVitales
                    comando.Parameters.Add(New SqlParameter("@Cabeza_Cuello", SqlDbType.NVarChar)).Value = objFormato.cabezaCuello
                    comando.Parameters.Add(New SqlParameter("@Torax", SqlDbType.NVarChar)).Value = objFormato.torax
                    comando.Parameters.Add(New SqlParameter("@Cardio_Pulmonar", SqlDbType.NVarChar)).Value = objFormato.cardioPulmonar
                    comando.Parameters.Add(New SqlParameter("@Abdomen", SqlDbType.NVarChar)).Value = objFormato.abdomen
                    comando.Parameters.Add(New SqlParameter("@Gestion_Urinaria", SqlDbType.NVarChar)).Value = objFormato.gestionUrinaria
                    comando.Parameters.Add(New SqlParameter("@Extremidades", SqlDbType.NVarChar)).Value = objFormato.extremidades
                    comando.Parameters.Add(New SqlParameter("@Nervio_Central", SqlDbType.NVarChar)).Value = objFormato.nervioCentral
                    comando.Parameters.Add(New SqlParameter("@Paraclinico", SqlDbType.NVarChar)).Value = objFormato.paraclinico
                    comando.Parameters.Add(New SqlParameter("@Analisis", SqlDbType.NVarChar)).Value = objFormato.analisis
                    comando.Parameters.Add(New SqlParameter("@Pronostico", SqlDbType.NVarChar)).Value = objFormato.pronostico
                    comando.Parameters.Add(New SqlParameter("@Tipo_Parto", SqlDbType.NVarChar)).Value = objFormato.tipoParto
                    comando.Parameters.Add(New SqlParameter("@Ruptura_Memb", SqlDbType.NVarChar)).Value = objFormato.rubtura
                    comando.Parameters.Add(New SqlParameter("@Induccion_Parto", SqlDbType.NVarChar)).Value = objFormato.inducionParto
                    comando.Parameters.Add(New SqlParameter("@Apgar1", SqlDbType.NVarChar)).Value = objFormato.apgar1
                    comando.Parameters.Add(New SqlParameter("@Apgar5", SqlDbType.NVarChar)).Value = objFormato.apgar5
                    comando.Parameters.Add(New SqlParameter("@Reanimacion_Nacer", SqlDbType.NVarChar)).Value = objFormato.reanimacionNacer
                    comando.Parameters.Add(New SqlParameter("@Edad_Madre", SqlDbType.NVarChar)).Value = objFormato.edadMadre
                    comando.Parameters.Add(New SqlParameter("@Edad_Gestacional", SqlDbType.NVarChar)).Value = objFormato.edadGestacional
                    comando.Parameters.Add(New SqlParameter("@Fum", SqlDbType.NVarChar)).Value = objFormato.fum
                    comando.Parameters.Add(New SqlParameter("@Obstetrico", SqlDbType.NVarChar)).Value = objFormato.obstetrico
                    comando.Parameters.Add(New SqlParameter("@Hemo_Madre", SqlDbType.NVarChar)).Value = objFormato.hemograsificacionMadre
                    comando.Parameters.Add(New SqlParameter("@Hemo_Padre", SqlDbType.NVarChar)).Value = objFormato.hemograsificacionPadre
                    comando.Parameters.Add(New SqlParameter("@Control_Prenatal", SqlDbType.NVarChar)).Value = objFormato.controlPrenatal
                    comando.Parameters.Add(New SqlParameter("@Med_Durante_Emb", SqlDbType.NVarChar)).Value = objFormato.medDuranteEmbarazo
                    comando.Parameters.Add(New SqlParameter("@Habitos", SqlDbType.NVarChar)).Value = objFormato.habitos
                    comando.Parameters.Add(New SqlParameter("@Infecciones_Emb", SqlDbType.NVarChar)).Value = objFormato.infeccionEmb
                    comando.Parameters.Add(New SqlParameter("@Diab_Gestional", SqlDbType.NVarChar)).Value = objFormato.diabGestional
                    comando.Parameters.Add(New SqlParameter("@Diab_Mellitud", SqlDbType.NVarChar)).Value = objFormato.diabMellitud
                    comando.Parameters.Add(New SqlParameter("@Hiper_Gestional", SqlDbType.NVarChar)).Value = objFormato.hiperGestional
                    comando.Parameters.Add(New SqlParameter("@Preclampcia", SqlDbType.NVarChar)).Value = objFormato.preclampcia
                    comando.Parameters.Add(New SqlParameter("@Vacunacion", SqlDbType.NVarChar)).Value = objFormato.vacunacion
                    comando.Parameters.Add(New SqlParameter("@Tiroidea", SqlDbType.NVarChar)).Value = objFormato.tiroidea
                    comando.Parameters.Add(New SqlParameter("@Torch", SqlDbType.NVarChar)).Value = objFormato.torch
                    comando.Parameters.Add(New SqlParameter("@Hemo_Menor", SqlDbType.NVarChar)).Value = objFormato.hemogracificacionNcido
                    comando.Parameters.Add(New SqlParameter("@VDRL", SqlDbType.NVarChar)).Value = objFormato.VDRL
                    comando.Parameters.Add(New SqlParameter("@TSH", SqlDbType.NVarChar)).Value = objFormato.TSH
                    comando.Parameters.Add(New SqlParameter("@Glucometria", SqlDbType.NVarChar)).Value = objFormato.glucometria
                    comando.Parameters.Add(New SqlParameter("@Generales", SqlDbType.NVarChar)).Value = objFormato.generales
                    comando.Parameters.Add(New SqlParameter("@Caract_Liquido", SqlDbType.NVarChar)).Value = objFormato.caracteristicaLiquido
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.NVarChar)).Value = objFormato.usuario
                    comando.Parameters.Add(New SqlParameter("@tablaDiagPrincipal", SqlDbType.Structured)).Value = objFormato.dtDiagPrincipales
                    comando.Parameters.Add(New SqlParameter("@tablaDiagImpresion", SqlDbType.Structured)).Value = objFormato.dtDiagImpresion
                    comando.Parameters.Add(New SqlParameter("@tablaMedicamento", SqlDbType.Structured)).Value = objFormato.dtMedicamento
                    comando.Parameters.Add(New SqlParameter("@tablaParaclinico", SqlDbType.Structured)).Value = objFormato.dtParaclinico
                    objFormato.codigo = CType(comando.ExecuteScalar, String)
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objFormato
    End Function
End Class
