'Imports System.IO

'Public Class ConvertirImagen
'    Public Shared Function Image2Bytes(ByVal img As Image) As Byte()
'        Dim sTemp As String = Path.GetTempFileName()
'        Dim fs As New FileStream(sTemp, FileMode.OpenOrCreate, FileAccess.ReadWrite)
'        img.Save(fs, System.Drawing.Imaging.ImageFormat.Jpeg)
'        fs.Position = 0
'        Dim imgLength As Integer = CInt(fs.Length)
'        Dim bytes(0 To imgLength - 1) As Byte
'        fs.Read(bytes, 0, imgLength)
'        fs.Close()
'        Return bytes
'    End Function

'    Public Shared Function Bytes2Image(ByVal bytes() As Byte) As Image
'        If bytes Is Nothing Then Return Nothing
'        '
'        Dim ms As New MemoryStream(bytes)
'        Dim bm As Bitmap = Nothing
'        Try
'            bm = New Bitmap(ms)
'        Catch ex As Exception
'            System.Diagnostics.Debug.WriteLine(ex.Message)
'        End Try
'        Return bm
'    End Function
'End Class
