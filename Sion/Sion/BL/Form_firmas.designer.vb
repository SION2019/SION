<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form_firmas
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
        Me.Pictcaptura = New System.Windows.Forms.PictureBox()
        Me.Btenviar = New System.Windows.Forms.Button()
        Me.Btborra = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        CType(Me.Pictcaptura, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Pictcaptura
        '
        Me.Pictcaptura.BackColor = System.Drawing.Color.White
        Me.Pictcaptura.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Pictcaptura.Location = New System.Drawing.Point(66, 12)
        Me.Pictcaptura.Name = "Pictcaptura"
        Me.Pictcaptura.Size = New System.Drawing.Size(420, 193)
        Me.Pictcaptura.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Pictcaptura.TabIndex = 0
        Me.Pictcaptura.TabStop = False
        '
        'Btenviar
        '
        Me.Btenviar.Image = Global.Celer.My.Resources.Resources.sign_check_icon
        Me.Btenviar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Btenviar.Location = New System.Drawing.Point(183, 271)
        Me.Btenviar.Name = "Btenviar"
        Me.Btenviar.Size = New System.Drawing.Size(78, 32)
        Me.Btenviar.TabIndex = 80018
        Me.Btenviar.Text = "Guardar"
        Me.Btenviar.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Btenviar.UseVisualStyleBackColor = True
        '
        'Btborra
        '
        Me.Btborra.Image = Global.Celer.My.Resources.Resources.sign_error_icon
        Me.Btborra.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Btborra.Location = New System.Drawing.Point(281, 271)
        Me.Btborra.Name = "Btborra"
        Me.Btborra.Size = New System.Drawing.Size(78, 32)
        Me.Btborra.TabIndex = 80019
        Me.Btborra.Text = "Borrar"
        Me.Btborra.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Btborra.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Celer.My.Resources.Resources.signature_icon__1_
        Me.PictureBox1.Location = New System.Drawing.Point(12, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(40, 40)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 80022
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(58, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 16)
        Me.Label1.TabIndex = 80020
        Me.Label1.Text = "FIRMA"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.DarkSeaGreen
        Me.Label2.Location = New System.Drawing.Point(54, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(492, 20)
        Me.Label2.TabIndex = 80021
        Me.Label2.Text = "_____________________________________________________________________"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Pictcaptura)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 54)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(553, 211)
        Me.GroupBox1.TabIndex = 80023
        Me.GroupBox1.TabStop = False
        '
        'Form_firmas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(559, 306)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Btborra)
        Me.Controls.Add(Me.Btenviar)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(575, 345)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(575, 345)
        Me.Name = "Form_firmas"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.Pictcaptura, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Pictcaptura As System.Windows.Forms.PictureBox
    Friend WithEvents Btenviar As Button
    Friend WithEvents Btborra As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents GroupBox1 As GroupBox
End Class
