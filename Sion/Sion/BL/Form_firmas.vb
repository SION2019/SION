
Imports System.IO
Imports System.Drawing.Imaging

Public Class Form_firmas
    Dim g As Graphics
    Dim image As Bitmap
    Private startX As Integer = 0
    Private startY As Integer = 0
    Private gBitmap As Bitmap
    Private isMouseDown As Boolean = False
    Dim linea As New Pen(Color.Black, 2)
    Public nRegistro As String
    Dim direct, consulta As String
    Dim idForm As String
    Dim vFormPadre As Object

    Public Sub iniciarForm(ByRef vFormPadre As Object)
        Me.vFormPadre = vFormPadre
    End Sub

    Public Sub crearImagen(idFormulario As String, registro As Integer, pImagen As PictureBox)

        If IsNothing(pImagen.Image) Then
            gBitmap = New Bitmap(Pictcaptura.Width, Pictcaptura.Height)
            g = Graphics.FromImage(gBitmap)
            g.Clear(Color.White)
            Pictcaptura.Image = gBitmap
        Else
            gBitmap = New Bitmap(pImagen.Width, pImagen.Height)
            g = Graphics.FromImage(gBitmap)
            Pictcaptura.Image = pImagen.Image
        End If

        idForm = idFormulario
        nRegistro = registro

        If idForm = Constantes.ID_ADMISIONES_A Then
            consulta = Consultas.FIRMA_ACOMPAÑANTE
        ElseIf idForm = Constantes.ID_ADMISIONES_R Then
            consulta = Consultas.FIRMA_RESPONSABLE
        ElseIf idForm = Constantes.ID_EPICRISIS Then
            consulta = Consultas.FIRMA_RETIRO
        ElseIf idForm = Constantes.ID_EMPLEADOS Then
        End If


    End Sub
    Private Sub Pictcaptura_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Pictcaptura.MouseDown
        isMouseDown = True
        startX = e.X
        startY = e.Y
    End Sub
    Private Sub Pictcaptura_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Pictcaptura.MouseMove

        If Pictcaptura.Enabled = True Then
            If isMouseDown = True Then

                Dim g As Graphics = Graphics.FromImage(gBitmap)
                g.DrawLine(linea, startX, startY, e.X, e.Y)

                Pictcaptura.Image = gBitmap

                startX = e.X
                startY = e.Y

            End If
        End If
    End Sub
    Private Sub Pictcaptura_MouseUp(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles Pictcaptura.MouseUp
        isMouseDown = False
    End Sub
    Private Sub Btborra_Click(sender As System.Object, e As System.EventArgs) Handles Btborra.Click
        Pictcaptura.Image = Nothing
        gBitmap = New Bitmap(Pictcaptura.Width, Pictcaptura.Height)
        g = Graphics.FromImage(gBitmap)
        g.Clear(Color.White)
        Me.Pictcaptura.Image = gBitmap
        crearImagen(idForm, nRegistro, Pictcaptura)
    End Sub
    Private Sub Form_firmas_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        crearImagen(idForm, nRegistro, Pictcaptura)
    End Sub
    Private Sub Btenviar_Click(sender As System.Object, e As System.EventArgs) Handles Btenviar.Click
        If idForm = Constantes.ID_EMPLEADOS Then
            vFormPadre.tomarImagen(Pictcaptura.Image)
            Me.Close()
        Else
            guardarDocumentos()
        End If

    End Sub
    Private Sub guardarDocumentos()

        Dim ruta As String
        Dim direccion As Object
        direccion = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\Firmas-" & Format(Date.Now, "yyyyMMdd")
        If Not File.Exists(direccion) Then
            Directory.CreateDirectory(direccion)
        End If
        ruta = direccion & "\" & idForm & ".png"
        Pictcaptura.Image.Save(ruta, ImageFormat.Png)

        Dim arrFile() As Byte
        Dim ms As MemoryStream
        ms = New MemoryStream
        Pictcaptura.Image.Save(ms, Imaging.ImageFormat.Png)
        arrFile = ms.GetBuffer
        Try
            FirmaDAL.guardarFirma(nRegistro, arrFile, consulta, IIf(idForm = Constantes.ID_EPICRISIS, True, False))
            MsgBox(Mensajes.GUARDADO)

            If idForm = Constantes.ID_ADMISIONES_A Then
                vFormPadre.cargarFirmaAcomp(nRegistro)
            ElseIf idForm = Constantes.ID_ADMISIONES_R Then
                vFormPadre.cargarFirmaRespon(nRegistro)
            ElseIf idForm = Constantes.ID_EPICRISIS Then
                vFormPadre.cargarFirma(nRegistro)
            End If
            Me.Close()
        Catch ex As Exception
            general.manejoErrores(ex) 
        End Try
    End Sub
End Class