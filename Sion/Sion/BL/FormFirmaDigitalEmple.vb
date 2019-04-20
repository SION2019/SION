
Public Class FormFirmaDigitalEmple
    Public Property formEmpleado As Form_empleado
    Private Sub Btenviar_Click(sender As Object, e As EventArgs) Handles btEnviar.Click
        Dim myimage As Image
        Dim ruta As String = System.IO.Path.GetTempPath & Replace(Replace(CStr(Format(DateTime.Now, Constantes.FORMATO_FECHA_HORA_GEN)), " ", ""), ":", "")
        SigPlusNET1.SetImageXSize(500)
        SigPlusNET1.SetImageYSize(150)
        SigPlusNET1.SetJustifyMode(5)
        myimage = SigPlusNET1.GetSigImage()
        myimage.Save(ruta, System.Drawing.Imaging.ImageFormat.Jpeg)
        SigPlusNET1.SetJustifyMode(0)
        formEmpleado.pictureFirma.Load(ruta)
        Dispose()
    End Sub
    Private Sub FormFirmaDigitalEmple_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim conectado As Boolean
        SigPlusNET1.SetTabletState(1)
        conectado = SigPlusNET1.TabletConnectQuery
    End Sub
    Private Sub btlimpiar_Click(sender As Object, e As EventArgs) Handles btlimpiar.Click
        SigPlusNET1.ClearTablet()
    End Sub
    Private Sub FormFirmaDigitalEmple_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        SigPlusNET1.SetTabletState(0)
    End Sub

End Class