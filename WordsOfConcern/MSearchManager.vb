Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Xml
Imports System.Data.Linq
Imports Microsoft.Office.Interop.Word

Module MSearchManager

    Friend msBookmarkNamePrefix As String = "bkmTemp_"

    Friend Sub HighlightWords(ByRef wordsGrid As DataGridView, ByRef color As Word.WdColorIndex, ByRef highlightedWordsListView As ListView)
        ' Get the range of the document
        Dim range As Word.Range = Globals.ThisAddIn.Application.ActiveDocument.Content

        ' Create a list to store the highlighted words and their locations
        Dim highlightedWords As New List(Of String)

        ' Loop through the rows of the DataGridView and highlight each word
        For Each row As DataGridViewRow In wordsGrid.Rows
            Dim word As String = row.Cells(1).Value

            ' Generate temporary bookmark names based on the search text
            Dim bookmarkNamePrefix As String = "bkmTemp"
            Dim bookmarkNameRegex As New Regex("[^\w\d]") ' Regular expression to remove non-alphanumeric characters

            Do While range.Find.Execute(FindText:=word)
                ' Highlight the found word
                range.HighlightColorIndex = color

                ' Add the highlighted word and its location to the list
                Dim start As Integer = range.Start
                Dim [end] As Integer = range.End
                Dim text As String = range.Text

                Dim bookmarkName As String = GenerateTempBookmarkName(text)

                Dim item As New ListViewItem(bookmarkName)

                item.SubItems.Add(bookmarkName)
                item.SubItems.Add(text)
                highlightedWordsListView.Items.Add(item)

                ' Move the range to the end of the highlighted word
                range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
            Loop

            ' Reset the range to the beginning of the document
            range.Start = Globals.ThisAddIn.Application.ActiveDocument.Content.Start
            range.End = Globals.ThisAddIn.Application.ActiveDocument.Content.End
        Next

    End Sub

    Friend Sub RemoveHighLights()
        Dim rngTemp As Range = Globals.ThisAddIn.Application.ActiveDocument.Range(Start:=0, End:=0)
        With rngTemp.Find
            .ClearFormatting()
            .Highlight = True
            With .Replacement
                .ClearFormatting()
                .Highlight = False
            End With
            .Execute(Replace:=WdReplace.wdReplaceAll, Forward:=True, FindText:="", ReplaceWith:="", Format:=True)
        End With
    End Sub

    Friend Function GenerateTempBookmarkName(ByRef word As String) As String

        Dim bookmarkNameRegex As New Regex("[^\w\d]") ' Regular expression to remove non-alphanumeric characters
        Dim bookmarkName As String = msBookmarkNamePrefix & bookmarkNameRegex.Replace(word, "")

        Return bookmarkName
    End Function

    Friend Sub DeleteBookMarks()
        ' Convert the Bookmarks collection to a sequence of Bookmark objects using Cast
        Dim bookmarkSequence = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Cast(Of Bookmark)()

        ' Use LINQ to filter the sequence based on the prefix
        Dim filteredBookmarks = From bm In bookmarkSequence
                                Where bm.Name.StartsWith(msBookmarkNamePrefix)
                                Select bm
        ' Loop through the filtered bookmarks and delete them
        If filteredBookmarks IsNot Nothing Then
            For Each bm As Bookmark In filteredBookmarks
                bm.Delete()
            Next
        End If
    End Sub

End Module
