Imports System.Data
Imports Microsoft.Office.Interop.Word
Imports System.Windows.Forms

Public Class WOCUserControl


    Private Sub btnUpdateWoCList_Click(sender As Object, e As EventArgs) Handles btnUpdateWoCList.Click
        MDataManager.LoadSourceGrid(dgvCurrent:=dgvWOCList)
    End Sub


    Private Sub btnHighlight_Click(sender As Object, e As EventArgs) Handles btnHighlight.Click
        Call MSearchManager.HighlightWords(wordsGrid:=dgvWOCList, color:=Word.WdColorIndex.wdYellow, highlightedWordsListView:=lvMatched)
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        MSearchManager.DeleteBookMarks()
        MSearchManager.RemoveHighLights()
    End Sub


End Class
