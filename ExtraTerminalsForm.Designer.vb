<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Extra_Terminals_Form
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
        Me.CDSAllDailySheetDGV = New System.Windows.Forms.DataGridView()
        Me.Label99 = New System.Windows.Forms.Label()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.cdsRADRONDGVColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cdsRADRONDGVColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cdsRADRONDGVColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cdsRADRONDGVColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.MgmtCdsManUpdBut = New System.Windows.Forms.Button()
        Me.MgmtCdsEmailBut = New System.Windows.Forms.Button()
        Me.MgmtCdsSaveBut = New System.Windows.Forms.Button()
        Me.MgmtCdsPrintBut = New System.Windows.Forms.Button()
        Me.MgmtDSDateLab = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.CDSTAllAddEngrPanelLB = New System.Windows.Forms.ListBox()
        Me.CDSTT4AddEngrPanel = New System.Windows.Forms.Panel()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.CDSAllDailySheetDGV, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.CDSTT4AddEngrPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'CDSAllDailySheetDGV
        '
        Me.CDSAllDailySheetDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.CDSAllDailySheetDGV.Location = New System.Drawing.Point(15, 209)
        Me.CDSAllDailySheetDGV.Name = "CDSAllDailySheetDGV"
        Me.CDSAllDailySheetDGV.RowHeadersVisible = False
        Me.CDSAllDailySheetDGV.Size = New System.Drawing.Size(1165, 245)
        Me.CDSAllDailySheetDGV.TabIndex = 213
        '
        'Label99
        '
        Me.Label99.AutoSize = True
        Me.Label99.Font = New System.Drawing.Font("Verdana", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label99.Location = New System.Drawing.Point(636, 78)
        Me.Label99.Name = "Label99"
        Me.Label99.Size = New System.Drawing.Size(152, 16)
        Me.Label99.TabIndex = 212
        Me.Label99.Text = "RAD/RON AIRCRAFT"
        '
        'DataGridView3
        '
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.cdsRADRONDGVColumn1, Me.cdsRADRONDGVColumn2, Me.cdsRADRONDGVColumn3, Me.cdsRADRONDGVColumn4})
        Me.DataGridView3.Location = New System.Drawing.Point(639, 97)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.RowHeadersVisible = False
        Me.DataGridView3.Size = New System.Drawing.Size(541, 106)
        Me.DataGridView3.TabIndex = 211
        '
        'cdsRADRONDGVColumn1
        '
        Me.cdsRADRONDGVColumn1.HeaderText = "GND Time"
        Me.cdsRADRONDGVColumn1.Name = "cdsRADRONDGVColumn1"
        '
        'cdsRADRONDGVColumn2
        '
        Me.cdsRADRONDGVColumn2.HeaderText = "Ship"
        Me.cdsRADRONDGVColumn2.Name = "cdsRADRONDGVColumn2"
        '
        'cdsRADRONDGVColumn3
        '
        Me.cdsRADRONDGVColumn3.HeaderText = "AMTs"
        Me.cdsRADRONDGVColumn3.Name = "cdsRADRONDGVColumn3"
        '
        'cdsRADRONDGVColumn4
        '
        Me.cdsRADRONDGVColumn4.HeaderText = "Comments"
        Me.cdsRADRONDGVColumn4.Name = "cdsRADRONDGVColumn4"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.MgmtCdsManUpdBut)
        Me.GroupBox1.Controls.Add(Me.MgmtCdsEmailBut)
        Me.GroupBox1.Controls.Add(Me.MgmtCdsSaveBut)
        Me.GroupBox1.Controls.Add(Me.MgmtCdsPrintBut)
        Me.GroupBox1.Location = New System.Drawing.Point(390, 86)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(240, 117)
        Me.GroupBox1.TabIndex = 210
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Daily Sheet Operations"
        '
        'MgmtCdsManUpdBut
        '
        Me.MgmtCdsManUpdBut.Location = New System.Drawing.Point(12, 24)
        Me.MgmtCdsManUpdBut.Name = "MgmtCdsManUpdBut"
        Me.MgmtCdsManUpdBut.Size = New System.Drawing.Size(104, 38)
        Me.MgmtCdsManUpdBut.TabIndex = 200
        Me.MgmtCdsManUpdBut.Text = "Manual Entry"
        Me.MgmtCdsManUpdBut.UseVisualStyleBackColor = True
        '
        'MgmtCdsEmailBut
        '
        Me.MgmtCdsEmailBut.Location = New System.Drawing.Point(122, 68)
        Me.MgmtCdsEmailBut.Name = "MgmtCdsEmailBut"
        Me.MgmtCdsEmailBut.Size = New System.Drawing.Size(104, 38)
        Me.MgmtCdsEmailBut.TabIndex = 201
        Me.MgmtCdsEmailBut.Text = "Email"
        Me.MgmtCdsEmailBut.UseVisualStyleBackColor = True
        '
        'MgmtCdsSaveBut
        '
        Me.MgmtCdsSaveBut.Location = New System.Drawing.Point(12, 68)
        Me.MgmtCdsSaveBut.Name = "MgmtCdsSaveBut"
        Me.MgmtCdsSaveBut.Size = New System.Drawing.Size(104, 38)
        Me.MgmtCdsSaveBut.TabIndex = 200
        Me.MgmtCdsSaveBut.Text = "Save"
        Me.MgmtCdsSaveBut.UseVisualStyleBackColor = True
        '
        'MgmtCdsPrintBut
        '
        Me.MgmtCdsPrintBut.Location = New System.Drawing.Point(122, 24)
        Me.MgmtCdsPrintBut.Name = "MgmtCdsPrintBut"
        Me.MgmtCdsPrintBut.Size = New System.Drawing.Size(104, 38)
        Me.MgmtCdsPrintBut.TabIndex = 0
        Me.MgmtCdsPrintBut.Text = "Print"
        Me.MgmtCdsPrintBut.UseVisualStyleBackColor = True
        '
        'MgmtDSDateLab
        '
        Me.MgmtDSDateLab.AutoSize = True
        Me.MgmtDSDateLab.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MgmtDSDateLab.Location = New System.Drawing.Point(76, 76)
        Me.MgmtDSDateLab.Name = "MgmtDSDateLab"
        Me.MgmtDSDateLab.Size = New System.Drawing.Size(50, 18)
        Me.MgmtDSDateLab.TabIndex = 209
        Me.MgmtDSDateLab.Text = "Date"
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.Location = New System.Drawing.Point(-67, 130)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(37, 18)
        Me.Label28.TabIndex = 208
        Me.Label28.Text = "For"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.Location = New System.Drawing.Point(12, 36)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(179, 36)
        Me.Label25.TabIndex = 207
        Me.Label25.Text = "Current Daily Sheet" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ALL TERMINALS"
        '
        'CDSTAllAddEngrPanelLB
        '
        Me.CDSTAllAddEngrPanelLB.FormattingEnabled = True
        Me.CDSTAllAddEngrPanelLB.Location = New System.Drawing.Point(133, 6)
        Me.CDSTAllAddEngrPanelLB.Name = "CDSTAllAddEngrPanelLB"
        Me.CDSTAllAddEngrPanelLB.Size = New System.Drawing.Size(213, 147)
        Me.CDSTAllAddEngrPanelLB.TabIndex = 0
        '
        'CDSTT4AddEngrPanel
        '
        Me.CDSTT4AddEngrPanel.Controls.Add(Me.Button11)
        Me.CDSTT4AddEngrPanel.Controls.Add(Me.Button9)
        Me.CDSTT4AddEngrPanel.Controls.Add(Me.Button1)
        Me.CDSTT4AddEngrPanel.Controls.Add(Me.CDSTAllAddEngrPanelLB)
        Me.CDSTT4AddEngrPanel.Location = New System.Drawing.Point(818, 255)
        Me.CDSTT4AddEngrPanel.Name = "CDSTT4AddEngrPanel"
        Me.CDSTT4AddEngrPanel.Size = New System.Drawing.Size(349, 159)
        Me.CDSTT4AddEngrPanel.TabIndex = 214
        '
        'Button11
        '
        Me.Button11.Location = New System.Drawing.Point(3, 98)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(124, 40)
        Me.Button11.TabIndex = 210
        Me.Button11.Text = "Done"
        Me.Button11.UseVisualStyleBackColor = True
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(3, 52)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(124, 40)
        Me.Button9.TabIndex = 209
        Me.Button9.Text = "Clear All"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(3, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(124, 40)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Add Selected Engineer"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 78)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 18)
        Me.Label1.TabIndex = 215
        Me.Label1.Text = "For"
        '
        'Extra_Terminals_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1170, 466)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CDSTT4AddEngrPanel)
        Me.Controls.Add(Me.CDSAllDailySheetDGV)
        Me.Controls.Add(Me.Label99)
        Me.Controls.Add(Me.DataGridView3)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.MgmtDSDateLab)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.Label25)
        Me.Name = "Extra_Terminals_Form"
        Me.Text = "Extra_Terminals_Form"
        CType(Me.CDSAllDailySheetDGV, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.CDSTT4AddEngrPanel.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CDSAllDailySheetDGV As System.Windows.Forms.DataGridView
    Friend WithEvents Label99 As System.Windows.Forms.Label
    Friend WithEvents DataGridView3 As System.Windows.Forms.DataGridView
    Friend WithEvents cdsRADRONDGVColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cdsRADRONDGVColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cdsRADRONDGVColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cdsRADRONDGVColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents MgmtCdsManUpdBut As System.Windows.Forms.Button
    Friend WithEvents MgmtCdsEmailBut As System.Windows.Forms.Button
    Friend WithEvents MgmtCdsSaveBut As System.Windows.Forms.Button
    Friend WithEvents MgmtCdsPrintBut As System.Windows.Forms.Button
    Friend WithEvents MgmtDSDateLab As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents CDSTAllAddEngrPanelLB As System.Windows.Forms.ListBox
    Friend WithEvents CDSTT4AddEngrPanel As System.Windows.Forms.Panel
    Friend WithEvents Button11 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
