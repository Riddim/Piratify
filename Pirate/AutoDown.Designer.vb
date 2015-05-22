<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AutoDown
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
        Dim DataSourceColumnBinder1 As XPTable.Models.DataSourceColumnBinder = New XPTable.Models.DataSourceColumnBinder()
        Dim DragDropRenderer1 As XPTable.Renderers.DragDropRenderer = New XPTable.Renderers.DragDropRenderer()
        Me.tmDown = New XPTable.Models.TableModel()
        Me.Table1 = New XPTable.Models.Table()
        Me.cmDown = New XPTable.Models.ColumnModel()
        Me.Down_Excel2 = New System.Windows.Forms.Button()
        CType(Me.Table1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Table1
        '
        Me.Table1.BorderColor = System.Drawing.Color.Black
        Me.Table1.ColumnModel = Me.cmDown
        Me.Table1.DataMember = Nothing
        Me.Table1.DataSourceColumnBinder = DataSourceColumnBinder1
        DragDropRenderer1.ForeColor = System.Drawing.Color.Red
        Me.Table1.DragDropRenderer = DragDropRenderer1
        Me.Table1.Location = New System.Drawing.Point(12, 35)
        Me.Table1.Name = "Table1"
        Me.Table1.Size = New System.Drawing.Size(692, 284)
        Me.Table1.TabIndex = 0
        Me.Table1.TableModel = Me.tmDown
        Me.Table1.Text = "Table1"
        Me.Table1.UnfocusedBorderColor = System.Drawing.Color.Black
        '
        'Down_Excel2
        '
        Me.Down_Excel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Down_Excel2.Location = New System.Drawing.Point(623, 3)
        Me.Down_Excel2.Name = "Down_Excel2"
        Me.Down_Excel2.Size = New System.Drawing.Size(81, 26)
        Me.Down_Excel2.TabIndex = 9
        Me.Down_Excel2.Text = "Down excel"
        Me.Down_Excel2.UseVisualStyleBackColor = True
        '
        'AutoDown
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(716, 331)
        Me.Controls.Add(Me.Down_Excel2)
        Me.Controls.Add(Me.Table1)
        Me.Name = "AutoDown"
        Me.Text = "AutoDown"
        CType(Me.Table1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tmDown As XPTable.Models.TableModel
    Friend WithEvents Table1 As XPTable.Models.Table
    Friend WithEvents cmDown As XPTable.Models.ColumnModel
    Friend WithEvents Down_Excel2 As System.Windows.Forms.Button
End Class
