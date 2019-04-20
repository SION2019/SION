Public Class ManualTarifario

    Public Property codigo As String
    Public Property codigoCups As String
    Public Property codigoIss As String
    Public Property codigoSoat As String
    Public Property usuario As Integer

    Public Sub guardarManual()
        ManualTarifarioDAL.guardar(Me)
    End Sub

    Public Function validar_codigo(ByVal codigomanual As Integer, ByVal codigocup As String) As Boolean
        Dim formManual As New Form_MANUAL_TARIFARIO
        If formManual.GroupBox2.Enabled = True Then
            Try
                Using consulta_verificacion As New SqlCommand("SP_MANUAL_TARIFARIO_CONSULTAR_CODIGO_EXISTENTE @CODIGO='" & codigocup & "',@codigomanual='" & codigomanual & "'", FormPrincipal.cnxion)
                    Using lector_ver = consulta_verificacion.ExecuteReader
                        If lector_ver.HasRows Then

                            Return True
                        Else
                            Return False
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function
End Class
