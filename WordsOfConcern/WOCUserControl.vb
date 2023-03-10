Imports System.Data
Imports Microsoft.Office.Interop.Word
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Windows.Media.Animation
Imports System.IO
Imports System.Collections.Specialized

Public Class WOCUserControl

    Private Sub btnHighlight_Click(sender As Object, e As EventArgs) Handles btnHighlight.Click
        TabControlMain.SelectedIndex = 0
        Call MSearchManager.HighlightWords(wordsGrid:=dgvWOCList, color:=Word.WdColorIndex.wdYellow, highlightedWordsListView:=lvMatched)
        SelectFirstItemIfMatchesExist()
    End Sub


    Private Sub SelectFirstItemIfMatchesExist()
        If lvMatched.Items.Count > 0 Then
            ' Select the first item
            lvMatched.Items(0).Selected = True
        Else
            ' Display a message box
            MessageBox.Show("No matches were found.")
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ResetMatched()
    End Sub

    Private Sub ResetMatched()
        MSearchManager.DeleteBookMarks()
        MSearchManager.RemoveHighLights()
        lvMatched.Items.Clear()
    End Sub


    Private Sub btnPrev_Click(sender As Object, e As EventArgs) Handles btnPrev.Click
        ' Get the currently selected item(s)
        Dim selectedIndices As ListView.SelectedIndexCollection = lvMatched.SelectedIndices

        ' Check if there are any selected items
        If selectedIndices.Count > 0 Then
            ' Get the index of the first selected item
            Dim selectedIndex As Integer = selectedIndices(0)

            ' Check if the selected item is not the first item
            If selectedIndex > 0 Then
                ' Deselect all items
                lvMatched.SelectedIndices.Clear()

                ' Select the previous item
                lvMatched.Items(selectedIndex - 1).Selected = True

                ' Get the bookmark name from the second column
                Dim bookmarkName As String = lvMatched.Items(selectedIndex - 1).SubItems(1).Text

                ' Select the bookmark in the Word document
                If Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Exists(bookmarkName) Then
                    Dim range As Word.Range = Globals.ThisAddIn.Application.ActiveDocument.Range
                    range.GoTo(WdGoToItem.wdGoToBookmark, , , bookmarkName)

                    Dim bookmarkRange As Word.Range = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks(bookmarkName).Range
                    bookmarkRange.Select()
                End If
            End If
        End If
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        ' Get the currently selected item(s)
        Dim selectedIndices As ListView.SelectedIndexCollection = lvMatched.SelectedIndices

        ' Check if there are any selected items
        If selectedIndices.Count > 0 Then
            ' Get the index of the last selected item
            Dim selectedIndex As Integer = selectedIndices(selectedIndices.Count - 1)

            ' Check if the selected item is not the last item
            If selectedIndex < lvMatched.Items.Count - 1 Then
                ' Deselect all items
                lvMatched.SelectedIndices.Clear()

                ' Select the next item
                lvMatched.Items(selectedIndex + 1).Selected = True

                ' Get the bookmark name from the second column
                Dim bookmarkName As String = lvMatched.Items(selectedIndex + 1).SubItems(1).Text

                ' Select the bookmark in the Word document
                If Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Exists(bookmarkName) Then
                    Dim range As Word.Range = Globals.ThisAddIn.Application.ActiveDocument.Range
                    range.GoTo(WdGoToItem.wdGoToBookmark, , , bookmarkName)

                    Dim bookmarkRange As Word.Range = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks(bookmarkName).Range
                    bookmarkRange.Select()
                End If
            End If
        End If
    End Sub

    Private Sub WOCUserControl_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If Me.Visible Then
            MDataManager.LoadWoCGridFromSettings(dgvCurrent:=dgvWOCList)
        End If


    End Sub


    Private Sub lvMatches_DoubleClick(sender As Object, e As EventArgs) Handles lvMatched.DoubleClick
        ' Get the selected item and the bookmark name in the second column.
        Dim selectedItem As ListViewItem = lvMatched.SelectedItems(0)
        Dim bookmarkName As String = selectedItem.SubItems(1).Text

        ' Select the bookmark in the active document.
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        If doc.Bookmarks.Exists(bookmarkName) Then
            Dim bookmarkRange As Word.Range = doc.Bookmarks(bookmarkName).Range
            bookmarkRange.Select()
        Else
            MessageBox.Show("The specified bookmark was not found.")
        End If
    End Sub


    Private Sub WOCUserControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TabControlMain.SelectedIndex = 1
        LoadWoCGridFromSettings(dgvWOCList)
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Dim stringCollection As New StringCollection()
        Try

            ' Read the selected CSV file
            Using reader As New StreamReader(OpenFileDialog1.FileName)
                'Read the file line by line
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    ' add the line to the StringCollection.
                    stringCollection.Add(line)
                End While

                My.Settings.TableData.Clear()
                My.Settings.TableData = stringCollection
                My.Settings.Save()
            End Using
        Catch ex As Exception
            ' Handle any errors that occur while reading the file
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnUpdateWoCList_Click(sender As Object, e As EventArgs) Handles btnUpdateWoCList.Click

        ' Show the OpenFileDialog to select a CSV file
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            TabControlMain.SelectedIndex = 1
            ' The FileOk event will handle reading the file and updating the settings variable
            MDataManager.LoadWoCGridFromSettings(dgvWOCList)
        End If

    End Sub
End Class
