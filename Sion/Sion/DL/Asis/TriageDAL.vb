Imports System.Data.SqlClient

Public Class TriageDAL
    Public Sub guardarTriage(ByRef objTriage As Triage)
        Try
            Dim diag As String = ""
            Dim sNerv As String = ""
            Dim a = Chr(124)

            If objTriage.codigoTriage <> "" AndAlso objTriage.sisNervCen.Contains(Chr(124)) Then
                objTriage.sisNervCen = Trim(objTriage.sisNervCen.Split(Chr(124))(1))
            End If
            If objTriage.tblParametros.Rows.Count > 0 Then
                diag = objTriage.tblParametros.Rows(0).Item("Descripcion") & ": " & objTriage.tblParametros.Rows(0).Item("Resultado")
            End If
            For indice = 1 To objTriage.tblParametros.Rows.Count - 1 Step 1

                If Constantes.PESO_TRIAGE <> objTriage.tblParametros.Rows(indice).Item("Codigo_Parametro") Then
                    If indice <= 6 Then
                        diag += ", " & objTriage.tblParametros.Rows(indice).Item("Descripcion") & ": " & objTriage.tblParametros.Rows(indice).Item("Resultado")
                    Else
                        sNerv += ", " & objTriage.tblParametros.Rows(indice).Item("Descripcion") & ": " & objTriage.tblParametros.Rows(indice).Item("Resultado")
                    End If
                End If
            Next

            Dim fila As DataRow() = objTriage.tblParametros.Select("Codigo_Parametro = '" & Constantes.PESO_TRIAGE & "'")
            Dim peso As String = fila(0).Item(2)

            Using sentencia = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    sentencia.Connection = FormPrincipal.cnxion
                    sentencia.Transaction = trnsccion
                    sentencia.CommandType = CommandType.StoredProcedure

                    sentencia.Parameters.Clear()
                    sentencia.CommandText = "[SP_TRIAGE_CREAR]"
                    sentencia.Parameters.AddWithValue("@pN_registro", objTriage.registro)
                    sentencia.Parameters.AddWithValue("@pCodigoNivelTriage", objTriage.nivelTriage)
                    sentencia.Parameters.AddWithValue("@pEsNeonatal", objTriage.esNeonatal)
                    sentencia.Parameters.AddWithValue("@pFecha", objTriage.fecha)
                    sentencia.Parameters.AddWithValue("@pCodigoAreatraslado", objTriage.codigoAreaServicio)
                    sentencia.Parameters.AddWithValue("@pTblParametros", objTriage.tblParametros)
                    sentencia.Parameters.AddWithValue("@pUsuario", SesionActual.idUsuario)
                    sentencia.Parameters.AddWithValue("@PESO", peso)
                    sentencia.Parameters.AddWithValue("@MOTIVO", objTriage.motivoConsulta)
                    sentencia.Parameters.AddWithValue("@ANTECEDENTES", objTriage.medicos)
                    sentencia.Parameters.AddWithValue("@ANTECEDENTES_QUI", objTriage.quirurgicos)
                    sentencia.Parameters.AddWithValue("@ANTECEDENTES_TRANS", objTriage.trasnfucionales)
                    sentencia.Parameters.AddWithValue("@ANTECEDENTES_ALER", objTriage.alergicos)
                    sentencia.Parameters.AddWithValue("@ANTECEDENTES_TRAU", objTriage.traumaticos)
                    sentencia.Parameters.AddWithValue("@ANTECEDENTES_TOX", objTriage.toxico)
                    sentencia.Parameters.AddWithValue("@ANTECEDENTES_FAM", objTriage.anteFamiliares)
                    sentencia.Parameters.AddWithValue("@REVISION_SIST", DBNull.Value)
                    sentencia.Parameters.AddWithValue("@SIG_VITALES", diag)
                    sentencia.Parameters.AddWithValue("@CAB_CUELLO", objTriage.cabCue)
                    sentencia.Parameters.AddWithValue("@TORAX", objTriage.torax)
                    sentencia.Parameters.AddWithValue("@CARDIO", objTriage.cardioPulmonar)
                    sentencia.Parameters.AddWithValue("@ABDOMEN", objTriage.abdomen)
                    sentencia.Parameters.AddWithValue("@GENTURINARIO", objTriage.genitoUrinario)
                    sentencia.Parameters.AddWithValue("@EXTREM", objTriage.extremidades)
                    sentencia.Parameters.AddWithValue("@S_NERV", objTriage.sisNervCen)
                    sentencia.Parameters.AddWithValue("@PARACLINICO", DBNull.Value)
                    sentencia.Parameters.AddWithValue("@ANALISIS", DBNull.Value)
                    sentencia.Parameters.AddWithValue("@PRONOSTICO", DBNull.Value)
                    sentencia.Parameters.AddWithValue("@CODIGO_EP", SesionActual.codigoEP)
                    sentencia.Parameters.AddWithValue("@GENERALES", sNerv) ''--> se tomara para enviar la variable 
                    sentencia.Parameters.AddWithValue("@T_PARTO", DBNull.Value)
                    sentencia.Parameters.AddWithValue("@T_RUPT_M", DBNull.Value)
                    sentencia.Parameters.AddWithValue("@INDU_PARTO", DBNull.Value)
                    sentencia.Parameters.AddWithValue("@CARACT_LIQUIDAS", DBNull.Value)
                    sentencia.Parameters.AddWithValue("@APGAR1", DBNull.Value)
                    sentencia.Parameters.AddWithValue("@APGAR5", DBNull.Value)
                    sentencia.Parameters.AddWithValue("@REANIM", DBNull.Value)
                    sentencia.Parameters.AddWithValue("@EDAD_MADRE", objTriage.edadMadre)
                    sentencia.Parameters.AddWithValue("@ANT_OBST", objTriage.obstetricos)
                    sentencia.Parameters.AddWithValue("@FUM", objTriage.fum)
                    sentencia.Parameters.AddWithValue("@EDAD_GESTA", objTriage.edadGestacional)
                    sentencia.Parameters.AddWithValue("@HEMOC_M", objTriage.hemM)
                    sentencia.Parameters.AddWithValue("@HEMOC_P", objTriage.hemP)
                    sentencia.Parameters.AddWithValue("@CONTROL", objTriage.controlPrenatal)
                    sentencia.Parameters.AddWithValue("@MEDC", objTriage.medDurantembarazo)
                    sentencia.Parameters.AddWithValue("@HABITO", objTriage.habitos)
                    sentencia.Parameters.AddWithValue("@INFECC", objTriage.infeccEmb)
                    sentencia.Parameters.AddWithValue("@DIABETE_G", objTriage.diabGestacional)
                    sentencia.Parameters.AddWithValue("@DIABETE_M", objTriage.diabMellitus)
                    sentencia.Parameters.AddWithValue("@HIPERTENCION", objTriage.hiperGestacional)
                    sentencia.Parameters.AddWithValue("@PRECLAMPCIA", objTriage.preclancia)
                    sentencia.Parameters.AddWithValue("@ENFER", objTriage.enfermTiroidea)
                    sentencia.Parameters.AddWithValue("@VACUNACION", objTriage.vacunacion)
                    sentencia.Parameters.AddWithValue("@TORCH", objTriage.torch)
                    sentencia.Parameters.AddWithValue("@HEMOCLASIFICACION", objTriage.hemoclasificacion)
                    sentencia.Parameters.AddWithValue("@VDRL", objTriage.vdrl)
                    sentencia.Parameters.AddWithValue("@TSH", objTriage.tsh)
                    sentencia.Parameters.AddWithValue("@GLUCOMETRIA", objTriage.glucometria)
                    sentencia.Parameters.AddWithValue("@vCodigoTriage", If(IsNothing(objTriage.codigoTriage), DBNull.Value, objTriage.codigoTriage))
                    objTriage.codigoTriage = sentencia.ExecuteScalar()

                    For indiceActual = 0 To objTriage.tblDiagnostico.Rows.Count - 2
                        sentencia.Parameters.Clear()
                        sentencia.CommandText = "[SP_INGRESO_DIAG_CREAR]"
                        sentencia.Parameters.AddWithValue("@N_Registro", objTriage.registro)
                        sentencia.Parameters.AddWithValue("@Codigo_Cie", objTriage.tblDiagnostico.Rows(indiceActual).Item("Codigo"))
                        sentencia.Parameters.AddWithValue("@CODIGO_EP", SesionActual.codigoEP)
                        sentencia.Parameters.AddWithValue("@usuario", SesionActual.idUsuario)
                        sentencia.Parameters.AddWithValue("@observacion", "")
                        sentencia.ExecuteNonQuery()
                    Next


                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        trnsccion.Rollback()
                    End Try
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
