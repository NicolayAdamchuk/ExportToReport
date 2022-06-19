<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormExportSchedule
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.BindingsTree = New System.Windows.Forms.TreeView()
        Me.ComboBoxScheduler = New System.Windows.Forms.ComboBox()
        Me.ButtonExport = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ButtonCancel
        '
        Me.ButtonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonCancel.Location = New System.Drawing.Point(598, 369)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(154, 40)
        Me.ButtonCancel.TabIndex = 0
        Me.ButtonCancel.Text = "Cancel"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'BindingsTree
        '
        Me.BindingsTree.Indent = 38
        Me.BindingsTree.Location = New System.Drawing.Point(639, 31)
        Me.BindingsTree.Name = "BindingsTree"
        Me.BindingsTree.Size = New System.Drawing.Size(138, 187)
        Me.BindingsTree.TabIndex = 1
        '
        'ComboBoxScheduler
        '
        Me.ComboBoxScheduler.FormattingEnabled = True
        Me.ComboBoxScheduler.Location = New System.Drawing.Point(52, 67)
        Me.ComboBoxScheduler.Name = "ComboBoxScheduler"
        Me.ComboBoxScheduler.Size = New System.Drawing.Size(301, 28)
        Me.ComboBoxScheduler.TabIndex = 2
        '
        'ButtonExport
        '
        Me.ButtonExport.Location = New System.Drawing.Point(52, 359)
        Me.ButtonExport.Name = "ButtonExport"
        Me.ButtonExport.Size = New System.Drawing.Size(228, 39)
        Me.ButtonExport.TabIndex = 3
        Me.ButtonExport.Text = "Export"
        Me.ButtonExport.UseVisualStyleBackColor = True
        '
        'FormExportSchedule
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.CausesValidation = False
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.ButtonExport)
        Me.Controls.Add(Me.ComboBoxScheduler)
        Me.Controls.Add(Me.BindingsTree)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Name = "FormExportSchedule"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FormExportSchedule"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents BindingsTree As System.Windows.Forms.TreeView
    Friend WithEvents ComboBoxScheduler As System.Windows.Forms.ComboBox
    Friend WithEvents ButtonExport As System.Windows.Forms.Button
End Class
