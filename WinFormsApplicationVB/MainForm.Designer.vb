<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ConverterWindow
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ConverterWindow))
        Me.chbxExportOld2Excel = New System.Windows.Forms.CheckBox()
        Me.chbxExportNew2Excel = New System.Windows.Forms.CheckBox()
        Me.pbLoadOldWord = New System.Windows.Forms.Button()
        Me.pbLoadNewWord = New System.Windows.Forms.Button()
        Me.pbCompareAndExport = New System.Windows.Forms.Button()
        Me.lblOldWord = New System.Windows.Forms.Label()
        Me.lblNewWord = New System.Windows.Forms.Label()
        Me.pbClean = New System.Windows.Forms.Button()
        Me.pbLoadPDBExcel = New System.Windows.Forms.Button()
        Me.lblLoadPDBExcel = New System.Windows.Forms.Label()
        Me.pbCompareWithPDBAndExport = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'chbxExportOld2Excel
        '
        Me.chbxExportOld2Excel.AutoSize = True
        Me.chbxExportOld2Excel.Location = New System.Drawing.Point(164, 20)
        Me.chbxExportOld2Excel.Name = "chbxExportOld2Excel"
        Me.chbxExportOld2Excel.Size = New System.Drawing.Size(97, 17)
        Me.chbxExportOld2Excel.TabIndex = 11
        Me.chbxExportOld2Excel.Text = "extractToExcel"
        Me.chbxExportOld2Excel.UseVisualStyleBackColor = True
        '
        'chbxExportNew2Excel
        '
        Me.chbxExportNew2Excel.AutoSize = True
        Me.chbxExportNew2Excel.Location = New System.Drawing.Point(164, 47)
        Me.chbxExportNew2Excel.Name = "chbxExportNew2Excel"
        Me.chbxExportNew2Excel.Size = New System.Drawing.Size(97, 17)
        Me.chbxExportNew2Excel.TabIndex = 12
        Me.chbxExportNew2Excel.Text = "extractToExcel"
        Me.chbxExportNew2Excel.UseVisualStyleBackColor = True
        '
        'pbLoadOldWord
        '
        Me.pbLoadOldWord.Location = New System.Drawing.Point(9, 15)
        Me.pbLoadOldWord.MaximumSize = New System.Drawing.Size(150, 23)
        Me.pbLoadOldWord.MinimumSize = New System.Drawing.Size(150, 23)
        Me.pbLoadOldWord.Name = "pbLoadOldWord"
        Me.pbLoadOldWord.Size = New System.Drawing.Size(150, 23)
        Me.pbLoadOldWord.TabIndex = 0
        Me.pbLoadOldWord.Text = "LoadOldWord"
        Me.pbLoadOldWord.UseVisualStyleBackColor = True
        '
        'pbLoadNewWord
        '
        Me.pbLoadNewWord.Location = New System.Drawing.Point(9, 44)
        Me.pbLoadNewWord.MaximumSize = New System.Drawing.Size(150, 23)
        Me.pbLoadNewWord.MinimumSize = New System.Drawing.Size(150, 23)
        Me.pbLoadNewWord.Name = "pbLoadNewWord"
        Me.pbLoadNewWord.Size = New System.Drawing.Size(150, 23)
        Me.pbLoadNewWord.TabIndex = 1
        Me.pbLoadNewWord.Text = "LoadNewWord"
        Me.pbLoadNewWord.UseVisualStyleBackColor = True
        '
        'pbCompareAndExport
        '
        Me.pbCompareAndExport.Location = New System.Drawing.Point(9, 73)
        Me.pbCompareAndExport.MaximumSize = New System.Drawing.Size(150, 23)
        Me.pbCompareAndExport.MinimumSize = New System.Drawing.Size(150, 23)
        Me.pbCompareAndExport.Name = "pbCompareAndExport"
        Me.pbCompareAndExport.Size = New System.Drawing.Size(150, 23)
        Me.pbCompareAndExport.TabIndex = 2
        Me.pbCompareAndExport.Text = "CompareAndExport"
        Me.pbCompareAndExport.UseVisualStyleBackColor = True
        '
        'lblOldWord
        '
        Me.lblOldWord.AutoSize = True
        Me.lblOldWord.Location = New System.Drawing.Point(258, 20)
        Me.lblOldWord.Name = "lblOldWord"
        Me.lblOldWord.Size = New System.Drawing.Size(16, 13)
        Me.lblOldWord.TabIndex = 3
        Me.lblOldWord.Text = " : "
        '
        'lblNewWord
        '
        Me.lblNewWord.AutoSize = True
        Me.lblNewWord.Location = New System.Drawing.Point(258, 49)
        Me.lblNewWord.Name = "lblNewWord"
        Me.lblNewWord.Size = New System.Drawing.Size(16, 13)
        Me.lblNewWord.TabIndex = 4
        Me.lblNewWord.Text = " : "
        '
        'pbClean
        '
        Me.pbClean.Location = New System.Drawing.Point(9, 102)
        Me.pbClean.MaximumSize = New System.Drawing.Size(150, 23)
        Me.pbClean.MinimumSize = New System.Drawing.Size(150, 23)
        Me.pbClean.Name = "pbClean"
        Me.pbClean.Size = New System.Drawing.Size(150, 23)
        Me.pbClean.TabIndex = 5
        Me.pbClean.Text = "Clean"
        Me.pbClean.UseVisualStyleBackColor = True
        '
        'pbLoadPDBExcel
        '
        Me.pbLoadPDBExcel.Location = New System.Drawing.Point(9, 131)
        Me.pbLoadPDBExcel.MaximumSize = New System.Drawing.Size(150, 23)
        Me.pbLoadPDBExcel.MinimumSize = New System.Drawing.Size(150, 23)
        Me.pbLoadPDBExcel.Name = "pbLoadPDBExcel"
        Me.pbLoadPDBExcel.Size = New System.Drawing.Size(150, 23)
        Me.pbLoadPDBExcel.TabIndex = 7
        Me.pbLoadPDBExcel.Text = "LoadPDBExcel"
        Me.pbLoadPDBExcel.UseCompatibleTextRendering = True
        Me.pbLoadPDBExcel.UseVisualStyleBackColor = True
        '
        'lblLoadPDBExcel
        '
        Me.lblLoadPDBExcel.AutoSize = True
        Me.lblLoadPDBExcel.Location = New System.Drawing.Point(161, 136)
        Me.lblLoadPDBExcel.Name = "lblLoadPDBExcel"
        Me.lblLoadPDBExcel.Size = New System.Drawing.Size(16, 13)
        Me.lblLoadPDBExcel.TabIndex = 9
        Me.lblLoadPDBExcel.Text = " : "
        '
        'pbCompareWithPDBAndExport
        '
        Me.pbCompareWithPDBAndExport.Enabled = False
        Me.pbCompareWithPDBAndExport.Location = New System.Drawing.Point(9, 160)
        Me.pbCompareWithPDBAndExport.MaximumSize = New System.Drawing.Size(150, 23)
        Me.pbCompareWithPDBAndExport.MinimumSize = New System.Drawing.Size(150, 23)
        Me.pbCompareWithPDBAndExport.Name = "pbCompareWithPDBAndExport"
        Me.pbCompareWithPDBAndExport.Size = New System.Drawing.Size(150, 23)
        Me.pbCompareWithPDBAndExport.TabIndex = 10
        Me.pbCompareWithPDBAndExport.Text = "CompareWithPDBAndExport"
        Me.pbCompareWithPDBAndExport.UseVisualStyleBackColor = True
        '
        'ConverterWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(738, 201)
        Me.Controls.Add(Me.chbxExportNew2Excel)
        Me.Controls.Add(Me.chbxExportOld2Excel)
        Me.Controls.Add(Me.pbCompareWithPDBAndExport)
        Me.Controls.Add(Me.lblLoadPDBExcel)
        Me.Controls.Add(Me.pbLoadPDBExcel)
        Me.Controls.Add(Me.pbClean)
        Me.Controls.Add(Me.lblNewWord)
        Me.Controls.Add(Me.lblOldWord)
        Me.Controls.Add(Me.pbCompareAndExport)
        Me.Controls.Add(Me.pbLoadNewWord)
        Me.Controls.Add(Me.pbLoadOldWord)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(854, 240)
        Me.MinimumSize = New System.Drawing.Size(300, 240)
        Me.Name = "ConverterWindow"
        Me.RightToLeftLayout = True
        Me.Text = "Converter"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents pbLoadOldWord As Button
    Friend WithEvents pbLoadNewWord As Button
    Friend WithEvents pbCompareAndExport As Button
    Friend WithEvents lblOldWord As Label
    Friend WithEvents lblNewWord As Label
    Friend WithEvents pbClean As Button
    Friend WithEvents pbLoadPDBExcel As Button
    Friend WithEvents lblLoadPDBExcel As Label
    Friend WithEvents pbCompareWithPDBAndExport As Button
    Friend WithEvents chbxExportOld2Excel As CheckBox
    Friend WithEvents chbxExportNew2Excel As CheckBox
End Class
