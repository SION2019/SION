Imports System.IO
Imports System.Drawing.Imaging
Public Class FormFirmaDigital
    Dim sign As Bitmap
    Dim ok As Bitmap
    Dim clear As Bitmap
    Dim please As Bitmap
    Dim lcdX As Integer
    Dim lcdY As Integer
    Dim lcdSize As Integer
    Dim screen As Integer
    Dim data As String
    Dim data2 As String
    Public nRegistro As String
    Dim direct, consulta As String
    Dim idForm As String
    Dim vFormPadre As Object
    Dim captura As String
    Dim estadoActivo As Boolean
    Dim ruta As String
    Public Sub iniciarForm(ByRef vFormPadre As Object)
        Me.vFormPadre = vFormPadre
    End Sub

    Public Sub crearImagen(idFormulario As String, registro As Integer)
        idForm = idFormulario
        nRegistro = registro

        If Not IsNothing(ruta) Then
            SigPlusNET1.ImportSigFile(ruta)
        End If

        If idForm = Constantes.ID_ADMISIONES_A Then
            consulta = Consultas.FIRMA_ACOMPAÑANTE
        ElseIf idForm = Constantes.ID_ADMISIONES_R Then
            consulta = Consultas.FIRMA_RESPONSABLE
        ElseIf idForm = Constantes.ID_EPICRISIS Then
            consulta = Consultas.FIRMA_RETIRO
        End If

    End Sub
    Private Sub FormFirmaDigital_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        salirFirmero()
    End Sub

    Private Sub SigPlusNET1_PenUp(sender As Object, e As EventArgs) Handles SigPlusNET1.PenUp
        activarFirmero()
    End Sub

    Private Sub FormFirmaDigital_Load(sender As Object, e As EventArgs) Handles Me.Load
        ruta = Path.GetTempPath & nRegistro & "_" & idForm & ".SIG"
        activarEventoFirma()
    End Sub


    Private Sub Btenviar_Click(sender As Object, e As EventArgs) Handles Btenviar.Click
        Dim picture As New PictureBox
        Try
            If idForm <> Constantes.ID_EMPLEADOS Then
                FirmaDAL.guardarFirma(nRegistro, File.ReadAllBytes(ruta), consulta, IIf(idForm = Constantes.ID_EPICRISIS, True, False))
                MsgBox(Mensajes.GUARDADO)

                If idForm = Constantes.ID_ADMISIONES_A Then
                    vFormPadre.cargarFirmaAcomp(nRegistro)
                ElseIf idForm = Constantes.ID_ADMISIONES_R Then
                    vFormPadre.cargarFirmaRespon(nRegistro)
                ElseIf idForm = Constantes.ID_EPICRISIS Then
                    vFormPadre.cargarFirma(nRegistro)
                End If

            Else
                ruta = Path.GetTempPath & nRegistro & "_" & idForm & ".png"
                SigPlusNET1.ImportSigFile(ruta)
                picture.Load("ruta")
                vFormPadre.pictureFirma = picture
            End If
            Me.Close()
        Catch ex As Exception
            General.manejoErrores(ex) 
        End Try

    End Sub
    Private Sub activarEventoFirma()
        'The following code will write BMP images out to the LCD 1X5 screen

        sign = New System.Drawing.Bitmap(Constantes.RUTA_SISTEMA & "\Sign.bmp")
        ok = New System.Drawing.Bitmap(Constantes.RUTA_SISTEMA & "\OK.bmp")
        clear = New System.Drawing.Bitmap(Constantes.RUTA_SISTEMA & "\CLEAR.bmp")
        please = New System.Drawing.Bitmap(Constantes.RUTA_SISTEMA & "\please.bmp")

        SigPlusNET1.SetTabletState(1)  'Turns tablet on to collect signature
        SigPlusNET1.LCDRefresh(0, 0, 0, 240, 64)
        SigPlusNET1.SetTranslateBitmapEnable(False)

        'Images sent to the background
        SigPlusNET1.LCDSendGraphic(1, 2, 0, 20, sign)
        SigPlusNET1.LCDSendGraphic(1, 2, 207, 4, ok)
        SigPlusNET1.LCDSendGraphic(1, 2, 15, 4, clear)

        lcdX = 240
        lcdY = 64

        SigPlusNET1.ClearTablet()

        SigPlusNET1.LCDSetWindow(0, 0, 1, 1)
        SigPlusNET1.SetSigWindow(1, 0, 0, 1, 1) 'Sets the area where ink is permitted in the SigPlus object
        SigPlusNET1.SetLCDCaptureMode(2)   'Sets mode so ink will not disappear after a few seconds

        estadoActivo = True
        activarFirmero()
    End Sub
    Private Sub salirFirmero()
        SigPlusNET1.LCDRefresh(0, 0, 0, 240, 64)
        SigPlusNET1.LCDSetWindow(0, 0, 240, 64)
        SigPlusNET1.SetSigWindow(1, 0, 0, 240, 64)
        SigPlusNET1.KeyPadClearHotSpotList()
        SigPlusNET1.SetLCDCaptureMode(1)
        SigPlusNET1.SetTabletState(0)
    End Sub
    Private Sub activarFirmero()
        If estadoActivo = True Then
            SigPlusNET1.ClearSigWindow(1)
            SigPlusNET1.LCDRefresh(1, 16, 45, 50, 15) 'Refresh LCD at 'Continue' to indicate to user that this option has been sucessfully chosen
            SigPlusNET1.LCDRefresh(2, 0, 0, 240, 64) 'Brings the background image already loaded into foreground
            SigPlusNET1.ClearTablet()
            SigPlusNET1.KeyPadClearHotSpotList()
            SigPlusNET1.KeyPadAddHotSpot(2, 1, 10, 5, 53, 17) 'For CLEAR button
            SigPlusNET1.KeyPadAddHotSpot(3, 1, 197, 5, 19, 17) 'For OK button
            SigPlusNET1.LCDSetWindow(0, 22, 238, 40)
            SigPlusNET1.SetSigWindow(1, 0, 22, 240, 40) 'Sets the area where ink is permitted in the SigPlus object
            SigPlusNET1.SetLCDCaptureMode(2)
            estadoActivo = False
        End If

        If SigPlusNET1.KeyPadQueryHotSpot(2) > 0 Then 'if the CLEAR hotspot is tapped, then...
            SigPlusNET1.ClearSigWindow(1)
            SigPlusNET1.LCDRefresh(1, 10, 0, 53, 17) 'Refresh LCD at 'CLEAR' to indicate to user that this option has been sucessfully chosen
            SigPlusNET1.LCDRefresh(2, 0, 0, 240, 64) 'Brings the background image already loaded into foreground
            SigPlusNET1.ClearTablet()
        End If

        If SigPlusNET1.KeyPadQueryHotSpot(3) > 0 Then 'if the OK hotspot is tapped, then...
            SigPlusNET1.ClearSigWindow(1)

            SigPlusNET1.ExportSigFile(ruta)

            SigPlusNET1.LCDRefresh(1, 210, 3, 14, 14) 'Refresh LCD at 'OK' to indicate to user that this option has been sucessfully chosen

            If SigPlusNET1.NumberOfTabletPoints > 0 Then
                SigPlusNET1.LCDRefresh(0, 0, 0, 240, 64)
                Dim f As Font = New System.Drawing.Font("Arial", 9.0F, System.Drawing.FontStyle.Regular)
                SigPlusNET1.LCDWriteString(0, 2, 35, 25, f, "se capturo la firma con exito.")
                System.Threading.Thread.Sleep(2000)
                Btenviar.Enabled = True
            Else
                SigPlusNET1.LCDRefresh(0, 0, 0, 240, 64)
                SigPlusNET1.LCDSendGraphic(0, 2, 4, 20, please)
                System.Threading.Thread.Sleep(750)
                SigPlusNET1.ClearTablet()
                SigPlusNET1.LCDRefresh(2, 0, 0, 240, 64)
                SigPlusNET1.SetLCDCaptureMode(2)   'Sets mode so ink will not disappear after a few seconds
            End If
        End If

        SigPlusNET1.ClearSigWindow(1)
    End Sub
End Class