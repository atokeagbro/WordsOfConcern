Imports System.Collections.Specialized
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports System.Xml.Serialization
Imports System.IO

Module MDataManager
    Private LocalWordsOfConcern As New StringCollection()

    Friend Sub LoadWoCGridFromSettings(ByRef dgvCurrent As DataGridView)
        Dim stringCollection As StringCollection = My.Settings.TableData

        Dim dataTable As New System.Data.DataTable
        dataTable.Columns.Add("Phrase_ID", GetType(String))
        dataTable.Columns.Add("Phrase_of_Concern", GetType(String))

        For Each item As String In stringCollection
            Dim values() As String = item.Split(","c)
            dataTable.Rows.Add(values)
        Next item

        dgvCurrent.DataSource = dataTable

    End Sub

End Module
