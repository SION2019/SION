<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormFirmaDigitalEmple
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.SigPlusNET1 = New Topaz.SigPlusNET()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btEnviar = New System.Windows.Forms.Button()
        Me.btlimpiar = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SigPlusNET1
        '
        Me.SigPlusNET1.BackColor = System.Drawing.Color.White
        Me.SigPlusNET1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SigPlusNET1.ForeColor = System.Drawing.Color.Black
        Me.SigPlusNET1.Location = New System.Drawing.Point(3, 17)
        Me.SigPlusNET1.Name = "SigPlusNET1"
        Me.SigPlusNET1.Size = New System.Drawing.Size(524, 164)
        Me.SigPlusNET1.TabIndex = 1
        Me.SigPlusNET1.Text = "SigPlusNET1"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.SigPlusNET1)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 60)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(530, 184)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Firma Digital"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Celer.My.Resources.Resources.signature_icon__1_
        Me.PictureBox1.Location = New System.Drawing.Point(5, 5)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(40, 40)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 80025
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(50, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(106, 16)
        Me.Label1.TabIndex = 80023
        Me.Label1.Text = "FIRMA EMPLEADO"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.SandyBrown
        Me.Label2.Location = New System.Drawing.Point(46, 21)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(492, 20)
        Me.Label2.TabIndex = 80024
        Me.Label2.Text = "_____________________________________________________________________"
        '
        'btEnviar
        '
        Me.btEnviar.Image = Global.Celer.My.Resources.Resources.Ok_icon1
        Me.btEnviar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btEnviar.Location = New System.Drawing.Point(158, 247)
        Me.btEnviar.Name = "btEnviar"
        Me.btEnviar.Size = New System.Drawing.Size(78, 32)
        Me.btEnviar.TabIndex = 80026
        Me.btEnviar.Text = "Enviar"
        Me.btEnviar.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btEnviar.UseVisualStyleBackColor = True
        '
        'btlimpiar
        '
        Me.btlimpiar.Image = Global.Celer.My.Resources.Resources.Close_icon
        Me.btlimpiar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btlimpiar.Location = New System.Drawing.Point(260, 247)
        Me.btlimpiar.Name = "btlimpiar"
        Me.btlimpiar.Size = New System.Drawing.Size(78, 32)
        Me.btlimpiar.TabIndex = 80027
        Me.btlimpiar.Text = "Limpiar"
        Me.btlimpiar.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btlimpiar.UseVisualStyleBackColor = True
        '
        'FormFirmaDigitalEmple
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(543, 289)
        Me.Controls.Add(Me.btlimpiar)
        Me.Controls.Add(Me.btEnviar)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.MaximumSize = New System.Drawing.Size(559, 323)
        Me.MinimumSize = New System.Drawing.Size(559, 323)
        Me.Name = "FormFirmaDigitalEmple"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SigPlusNET1 As Topaz.SigPlusNET
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents btEnviar As Button
    Friend WithEvents btlimpiar As Button
End Class
