<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DailySheetOperationsEmailCredentials
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
        Me.DsOpsEmailAddressTB = New System.Windows.Forms.TextBox()
        Me.DsOpsEmailPasswordTB = New System.Windows.Forms.TextBox()
        Me.EngrMyInfoDispSurnameLab = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DsOpsCredentialsDoneBut = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DsOpsEmpNoTB = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'DsOpsEmailAddressTB
        '
        Me.DsOpsEmailAddressTB.Location = New System.Drawing.Point(12, 83)
        Me.DsOpsEmailAddressTB.Name = "DsOpsEmailAddressTB"
        Me.DsOpsEmailAddressTB.Size = New System.Drawing.Size(213, 20)
        Me.DsOpsEmailAddressTB.TabIndex = 0
        '
        'DsOpsEmailPasswordTB
        '
        Me.DsOpsEmailPasswordTB.Location = New System.Drawing.Point(12, 122)
        Me.DsOpsEmailPasswordTB.Name = "DsOpsEmailPasswordTB"
        Me.DsOpsEmailPasswordTB.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.DsOpsEmailPasswordTB.Size = New System.Drawing.Size(213, 20)
        Me.DsOpsEmailPasswordTB.TabIndex = 1
        Me.DsOpsEmailPasswordTB.UseSystemPasswordChar = True
        '
        'EngrMyInfoDispSurnameLab
        '
        Me.EngrMyInfoDispSurnameLab.AutoSize = True
        Me.EngrMyInfoDispSurnameLab.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EngrMyInfoDispSurnameLab.Location = New System.Drawing.Point(9, 67)
        Me.EngrMyInfoDispSurnameLab.Name = "EngrMyInfoDispSurnameLab"
        Me.EngrMyInfoDispSurnameLab.Size = New System.Drawing.Size(88, 13)
        Me.EngrMyInfoDispSurnameLab.TabIndex = 2
        Me.EngrMyInfoDispSurnameLab.Text = "Email Address"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 106)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Password"
        '
        'DsOpsCredentialsDoneBut
        '
        Me.DsOpsCredentialsDoneBut.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DsOpsCredentialsDoneBut.Location = New System.Drawing.Point(260, 122)
        Me.DsOpsCredentialsDoneBut.Name = "DsOpsCredentialsDoneBut"
        Me.DsOpsCredentialsDoneBut.Size = New System.Drawing.Size(112, 33)
        Me.DsOpsCredentialsDoneBut.TabIndex = 4
        Me.DsOpsCredentialsDoneBut.Text = "Done"
        Me.DsOpsCredentialsDoneBut.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(9, 28)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Employee ID"
        '
        'DsOpsEmpNoTB
        '
        Me.DsOpsEmpNoTB.Location = New System.Drawing.Point(12, 44)
        Me.DsOpsEmpNoTB.Name = "DsOpsEmpNoTB"
        Me.DsOpsEmpNoTB.Size = New System.Drawing.Size(213, 20)
        Me.DsOpsEmpNoTB.TabIndex = 5
        '
        'DailySheetOperationsEmailCredentials
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(384, 158)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DsOpsEmpNoTB)
        Me.Controls.Add(Me.DsOpsCredentialsDoneBut)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.EngrMyInfoDispSurnameLab)
        Me.Controls.Add(Me.DsOpsEmailPasswordTB)
        Me.Controls.Add(Me.DsOpsEmailAddressTB)
        Me.Name = "DailySheetOperationsEmailCredentials"
        Me.Text = "LEMS Email Credentials"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DsOpsEmailAddressTB As System.Windows.Forms.TextBox
    Friend WithEvents DsOpsEmailPasswordTB As System.Windows.Forms.TextBox
    Friend WithEvents EngrMyInfoDispSurnameLab As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DsOpsCredentialsDoneBut As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DsOpsEmpNoTB As System.Windows.Forms.TextBox
End Class
