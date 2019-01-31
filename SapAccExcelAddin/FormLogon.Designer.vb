<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormLogon
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.UserName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Password = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ButtonLogon = New System.Windows.Forms.Button()
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.Client = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Language = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Destination = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'UserName
        '
        Me.UserName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.UserName.Location = New System.Drawing.Point(96, 73)
        Me.UserName.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.UserName.MaxLength = 12
        Me.UserName.Name = "UserName"
        Me.UserName.Size = New System.Drawing.Size(203, 22)
        Me.UserName.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 80)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 17)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "User"
        '
        'Password
        '
        Me.Password.Location = New System.Drawing.Point(96, 106)
        Me.Password.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Password.Name = "Password"
        Me.Password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.Password.Size = New System.Drawing.Size(203, 22)
        Me.Password.TabIndex = 2
        Me.Password.UseSystemPasswordChar = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(17, 114)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 17)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Password"
        '
        'ButtonLogon
        '
        Me.ButtonLogon.Location = New System.Drawing.Point(16, 179)
        Me.ButtonLogon.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonLogon.Name = "ButtonLogon"
        Me.ButtonLogon.Size = New System.Drawing.Size(72, 31)
        Me.ButtonLogon.TabIndex = 4
        Me.ButtonLogon.Text = "Logon"
        Me.ButtonLogon.UseVisualStyleBackColor = True
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Location = New System.Drawing.Point(96, 180)
        Me.ButtonCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(72, 31)
        Me.ButtonCancel.TabIndex = 5
        Me.ButtonCancel.Text = "Cancel"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'Client
        '
        Me.Client.Location = New System.Drawing.Point(96, 41)
        Me.Client.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Client.MaxLength = 3
        Me.Client.Name = "Client"
        Me.Client.Size = New System.Drawing.Size(44, 22)
        Me.Client.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(17, 50)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(43, 17)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Client"
        '
        'Language
        '
        Me.Language.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.Language.Location = New System.Drawing.Point(96, 138)
        Me.Language.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Language.MaxLength = 2
        Me.Language.Name = "Language"
        Me.Language.Size = New System.Drawing.Size(44, 22)
        Me.Language.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(17, 147)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 17)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Language"
        '
        'Destination
        '
        Me.Destination.BackColor = System.Drawing.SystemColors.Control
        Me.Destination.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.Destination.Enabled = False
        Me.Destination.Location = New System.Drawing.Point(20, 10)
        Me.Destination.Margin = New System.Windows.Forms.Padding(4)
        Me.Destination.MaxLength = 12
        Me.Destination.Name = "Destination"
        Me.Destination.Size = New System.Drawing.Size(279, 22)
        Me.Destination.TabIndex = 10
        Me.Destination.TabStop = False
        Me.Destination.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'FormLogon
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(325, 223)
        Me.Controls.Add(Me.Destination)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Language)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Client)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonLogon)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Password)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.UserName)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "FormLogon"
        Me.Text = "SAP-Logon"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents UserName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Password As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ButtonLogon As System.Windows.Forms.Button
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents Client As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Language As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Destination As System.Windows.Forms.TextBox
End Class
