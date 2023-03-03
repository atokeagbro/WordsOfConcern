Imports Microsoft.Office.Tools.Ribbon

Public Class rbnWordOfConcern

    Private Sub btnOpenTaskPane_Click(sender As Object, e As RibbonControlEventArgs) Handles btnOpenTaskPane.Click

        ' Create an instance of the user control
        Dim wocUserControl As New WOCUserControl

        ' Add the user control to the custom task panes collection
        Dim myTaskPane As Microsoft.Office.Tools.CustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(wocUserControl, "Words of Concern User Control")

        ' Set the width and dock position of the task pane
        myTaskPane.Width = 300
        myTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight

        ' Show the task pane
        myTaskPane.Visible = True

    End Sub

End Class
