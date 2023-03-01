<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class WOCUserControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnPrev = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnUpdateWoCList = New System.Windows.Forms.Button()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.btnHighlight = New System.Windows.Forms.Button()
        Me.TabControlMain = New System.Windows.Forms.TabControl()
        Me.tpWOC = New System.Windows.Forms.TabPage()
        Me.lvMatched = New System.Windows.Forms.ListView()
        Me.colHdrMatched = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.colHdrBookMark = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.tpMatched = New System.Windows.Forms.TabPage()
        Me.dgvWOCList = New System.Windows.Forms.DataGridView()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.TabControlMain.SuspendLayout()
        Me.tpWOC.SuspendLayout()
        Me.tpMatched.SuspendLayout()
        CType(Me.dgvWOCList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnPrev
        '
        Me.btnPrev.Location = New System.Drawing.Point(4, 131)
        Me.btnPrev.Name = "btnPrev"
        Me.btnPrev.Size = New System.Drawing.Size(64, 23)
        Me.btnPrev.TabIndex = 2
        Me.btnPrev.Text = "Previous"
        Me.btnPrev.UseVisualStyleBackColor = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(128, 131)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(64, 23)
        Me.btnNext.TabIndex = 3
        Me.btnNext.Text = "Next"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnUpdateWoCList
        '
        Me.btnUpdateWoCList.Location = New System.Drawing.Point(3, 53)
        Me.btnUpdateWoCList.Name = "btnUpdateWoCList"
        Me.btnUpdateWoCList.Size = New System.Drawing.Size(189, 23)
        Me.btnUpdateWoCList.TabIndex = 4
        Me.btnUpdateWoCList.Text = "Update WoC List"
        Me.btnUpdateWoCList.UseVisualStyleBackColor = True
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(8, 11)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(163, 24)
        Me.lblTitle.TabIndex = 6
        Me.lblTitle.Text = "Words of Concern"
        '
        'btnHighlight
        '
        Me.btnHighlight.Location = New System.Drawing.Point(4, 92)
        Me.btnHighlight.Name = "btnHighlight"
        Me.btnHighlight.Size = New System.Drawing.Size(107, 23)
        Me.btnHighlight.TabIndex = 7
        Me.btnHighlight.Text = "Highlight Matches"
        Me.btnHighlight.UseVisualStyleBackColor = True
        '
        'TabControlMain
        '
        Me.TabControlMain.Controls.Add(Me.tpWOC)
        Me.TabControlMain.Controls.Add(Me.tpMatched)
        Me.TabControlMain.Location = New System.Drawing.Point(4, 176)
        Me.TabControlMain.Name = "TabControlMain"
        Me.TabControlMain.SelectedIndex = 0
        Me.TabControlMain.Size = New System.Drawing.Size(363, 440)
        Me.TabControlMain.TabIndex = 8
        '
        'tpWOC
        '
        Me.tpWOC.Controls.Add(Me.lvMatched)
        Me.tpWOC.Location = New System.Drawing.Point(4, 22)
        Me.tpWOC.Name = "tpWOC"
        Me.tpWOC.Padding = New System.Windows.Forms.Padding(3)
        Me.tpWOC.Size = New System.Drawing.Size(355, 414)
        Me.tpWOC.TabIndex = 0
        Me.tpWOC.Text = "Matched Items"
        Me.tpWOC.UseVisualStyleBackColor = True
        '
        'lvMatched
        '
        Me.lvMatched.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colHdrMatched, Me.colHdrBookMark})
        Me.lvMatched.HideSelection = False
        Me.lvMatched.Location = New System.Drawing.Point(6, 34)
        Me.lvMatched.Name = "lvMatched"
        Me.lvMatched.Size = New System.Drawing.Size(336, 377)
        Me.lvMatched.TabIndex = 0
        Me.lvMatched.UseCompatibleStateImageBehavior = False
        Me.lvMatched.View = System.Windows.Forms.View.Details
        '
        'colHdrMatched
        '
        Me.colHdrMatched.Text = "Matched"
        Me.colHdrMatched.Width = 163
        '
        'colHdrBookMark
        '
        Me.colHdrBookMark.Text = "Bookmark"
        Me.colHdrBookMark.Width = 255
        '
        'tpMatched
        '
        Me.tpMatched.BackColor = System.Drawing.Color.Transparent
        Me.tpMatched.Controls.Add(Me.dgvWOCList)
        Me.tpMatched.Location = New System.Drawing.Point(4, 22)
        Me.tpMatched.Name = "tpMatched"
        Me.tpMatched.Padding = New System.Windows.Forms.Padding(3)
        Me.tpMatched.Size = New System.Drawing.Size(355, 414)
        Me.tpMatched.TabIndex = 1
        Me.tpMatched.Text = "Words of Concern"
        '
        'dgvWOCList
        '
        Me.dgvWOCList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvWOCList.Location = New System.Drawing.Point(3, 34)
        Me.dgvWOCList.Name = "dgvWOCList"
        Me.dgvWOCList.Size = New System.Drawing.Size(336, 377)
        Me.dgvWOCList.TabIndex = 0
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(117, 91)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(75, 23)
        Me.btnClear.TabIndex = 9
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'WOCUserControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.TabControlMain)
        Me.Controls.Add(Me.btnHighlight)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.btnUpdateWoCList)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnPrev)
        Me.Name = "WOCUserControl"
        Me.Size = New System.Drawing.Size(504, 656)
        Me.TabControlMain.ResumeLayout(False)
        Me.tpWOC.ResumeLayout(False)
        Me.tpMatched.ResumeLayout(False)
        CType(Me.dgvWOCList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnPrev As Windows.Forms.Button
    Friend WithEvents btnNext As Windows.Forms.Button
    Friend WithEvents btnUpdateWoCList As Windows.Forms.Button
    Friend WithEvents lblTitle As Windows.Forms.Label
    Friend WithEvents btnHighlight As Windows.Forms.Button
    Friend WithEvents TabControlMain As Windows.Forms.TabControl
    Friend WithEvents tpWOC As Windows.Forms.TabPage
    Friend WithEvents tpMatched As Windows.Forms.TabPage
    Friend WithEvents dgvWOCList As Windows.Forms.DataGridView
    Friend WithEvents btnClear As Windows.Forms.Button
    Friend WithEvents lvMatched As Windows.Forms.ListView
    Friend WithEvents colHdrMatched As Windows.Forms.ColumnHeader
    Friend WithEvents colHdrBookMark As Windows.Forms.ColumnHeader
End Class
