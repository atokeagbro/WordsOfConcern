Imports System.Windows.Forms

Public Class ThisAddIn
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ' Add a handler for the DocumentBeforeClose event
        AddHandler Globals.ThisAddIn.Application.DocumentBeforeClose, AddressOf Application_DocumentBeforeClose
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' Remove the handler for the DocumentBeforeClose event
        RemoveHandler Globals.ThisAddIn.Application.DocumentBeforeClose, AddressOf Application_DocumentBeforeClose
    End Sub

    Private Sub Application_DocumentBeforeClose(Doc As Word.Document, ByRef Cancel As Boolean) Handles Application.DocumentBeforeClose
        ' Check if there are any bookmarks with the bkmTemp_ prefix
        Dim hasTempBookmarks As Boolean = False
        For Each bookmark As Word.Bookmark In Doc.Bookmarks
            If bookmark.Name.StartsWith("bkmTemp_") Then
                hasTempBookmarks = True
                Exit For
            End If
        Next bookmark


        ' If there are temp bookmarks, ask the user if they want to remove them
        If hasTempBookmarks Then
            Dim result As DialogResult = MessageBox.Show("The document contains Words of Concern bookmarks. Do you want to remove them before closing?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                MSearchManager.DeleteBookMarks()
                MSearchManager.RemoveHighLights()
                Doc.Save()
            End If
        End If


    End Sub
End Class
