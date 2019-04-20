Imports System.Data.SqlClient
Public Class FchaTmpo
    '**********************************************************************************************************************************************************
    '*  Por favor no borrar estos comentarios                                                                                                                 *
    '*                                                                                                                                                        *
    '*  NOMBRRE:        FCHATMPO                                                                                                                              *
    '*  FECHA:          27 DE SEPTIEMBRE DE 2012                                                                                                              *
    '*  AUTOR:          JAIDY BOTELLO ROCHA                                                                                                                   *
    '*  DESCRIPCION:    Esta clase devuelve la fecha y/o la hora del servidor  de SQL SERVER. La clase contiene dos funciones de acceso publico.              *
    '*                  Una para obtener la hora: Hora, y otra para obtener la fecha: Fecha. Ambas funciones devuelven valores tipo String                    *
    '*                  Es necesario enviarles a las funciones unos parametros que se relacionan mas adelantes, y que permiten escoger el formato             *
    '*                  para la fecha y/o la hora. Si el parametro no es correcto las funciones devuelve el valor "0" como string                             *
    '*                                                                                                                                                        *
    '*                  A continuacion los parametros validos para la clase                                                                                   *
    '*                                                                                                                                                        *
    '*                  Ejemplo para la fecha 27 de Septiembre de 2012             Ejemplo para la hora 5:32 PM con 55 segundos                               *
    '*                     Ej                 Parametros                               Ej                 Parametros                                          *
    '*                  09/27/12                1 o 101                            17:32:55                8 o 108                                            *
    '*                  12.09.27                2 o 102                            17:32:55:583            14 o 114                                           *
    '*                  27/09/12                3 o 103                                                                                                       *
    '*                  27.09.12                4 o 104                                                                                                       *
    '*                  27-09-12                5 o 105                                                                                                       *
    '*                  27 Sep 12               6 o 106                            Ejemplo para la Fecha y hora completa                                      *
    '*                  Sep 27, 12              7 o 107                                    Ej                        Parametros                               *
    '*                  09-27-12                10 o 108                           Sep 27 2012  5:32:55:583PM          9 o 109                                *
    '*                  12/09/27                11 o 111                           27 Sep 2012 17:32:55:583            13 o 113                               *
    '*                  120927                  12 o 112                           2012-11-22 21:17:40                  20                                                                        *
    '*                 2012-10-04                     23                            11/22/2012 9:16:02 PM
    '                                                                                                12/1/2012 10:19:25 PM                                           *
    '**********************************************************************************************************************************************************

    Public Shared Function Fecha(ByVal frmto As Integer) As String
        Dim fcha As String
        If Validar(frmto) Then
            fcha = solicita("SELECT CONVERT(varchar,GETDATE()," & 20 & ") AS Fecha", "Fecha")
            fcha = Format(CDate(fcha), Constantes.FORMATO_FECHA_HORA_GEN)
            Return fcha
        Else
            Return "0"
        End If
    End Function

    Public Shared Function Hora(ByVal frmto As Integer) As String
        Dim hra As String
        If Validar(frmto) Then
            hra = solicita("SELECT CONVERT(varchar,GETDATE()," & frmto.ToString & ") AS Hora", "Hora")
            Return hra
        Else
            Return "0"
        End If
    End Function

    Private Shared Function solicita(ByVal cdna As String, ByVal itm As String) As String
        'Dim cnslta As SqlCommand
        'Dim rsltdo As SqlDataReader
        'Dim FormPrincipal.cnxion As New ConeccionBD
        Dim vlor As String = ""
        'FormPrincipal.cnxion.conectarbd()
        Try

            'FormPrincipal.Timer1.Stop()
            Using cnslta = New SqlCommand(cdna, FormPrincipal.cnxion)
                Using rsltdo = cnslta.ExecuteReader()
                    If rsltdo.HasRows Then
                        rsltdo.Read()
                        vlor = rsltdo.Item(itm).ToString
                    End If
                End Using
            End Using

        Catch ex As Exception
            general.manejoErrores(ex)
            'Finally
            '    'If rsltdo.IsClosed = False Then
            '    '    rsltdo.Close()
            '    'End If
            '    'FormPrincipal.cnxion.desconectarbd()
            '    FormPrincipal.Timer1.Start()
        End Try
        Return vlor
    End Function

    Private Shared Function Validar(ByVal frmto As Integer) As Boolean
        If (0 <= frmto < 20) Or (99 < frmto < 115) Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
