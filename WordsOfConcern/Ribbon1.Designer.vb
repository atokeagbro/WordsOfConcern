Partial Class rbnWordOfConcern
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(rbnWordOfConcern))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.tabWordOfConcern = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.btnOpenTaskPane = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.tabWordOfConcern.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'tabWordOfConcern
        '
        Me.tabWordOfConcern.Groups.Add(Me.Group2)
        Me.tabWordOfConcern.Label = "Words of Concern"
        Me.tabWordOfConcern.Name = "tabWordOfConcern"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnOpenTaskPane)
        Me.Group2.Label = "Words of Concern "
        Me.Group2.Name = "Group2"
        '
        'btnOpenTaskPane
        '
        Me.btnOpenTaskPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnOpenTaskPane.Image = CType(resources.GetObject("btnOpenTaskPane.Image"), System.Drawing.Image)
        Me.btnOpenTaskPane.Label = "Open WoC Task Pane"
        Me.btnOpenTaskPane.Name = "btnOpenTaskPane"
        Me.btnOpenTaskPane.ShowImage = True
        '
        'rbnWordOfConcern
        '
        Me.Name = "rbnWordOfConcern"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.tabWordOfConcern)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.tabWordOfConcern.ResumeLayout(False)
        Me.tabWordOfConcern.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents tabWordOfConcern As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnOpenTaskPane As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As rbnWordOfConcern
        Get
            Return Me.GetRibbon(Of rbnWordOfConcern)()
        End Get
    End Property
End Class
