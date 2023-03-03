Imports System.Collections.Specialized
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports System.Xml.Serialization

Module MDataManager
    Private LocalWordsOfConcern As New StringCollection()

    'ReadOnly mSourceFilePath As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "FreddieMacPhrasesOfConcern.xlsx")
    ReadOnly mSourceFilePath As String = "C:\Users\aoa82\OneDrive - ProsperSpark\ProsperSpark\BBG\Proof of Concepts\Words of Concern Base\FreddieMacPhrasesOfConcern.xlsx"
    ReadOnly mDBConn As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mSourceFilePath + ";Extended Properties=Excel 12.0;"
    ReadOnly mSQLSelectData As String = "SELECT * FROM [Sheet1$]"


    ''' <summary>
    ''' Whenever we update the data from the outside, update our internal version as well.
    ''' </summary>
    ''' <param name="dgvCurrent"></param>
    Friend Sub LoadSourceGrid(ByRef dgvCurrent As DataGridView)

        Dim MyConnection As OleDbConnection = New System.Data.OleDb.OleDbConnection(mDBConn)
        MyConnection.Open()

        Dim MyCommand As OleDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter(mSQLSelectData, MyConnection)

        Dim dataSet As System.Data.DataSet = New System.Data.DataSet
        MyCommand.Fill(dataSet)

        MyConnection.Close()

        Dim dataTable As System.Data.DataTable = dataSet.Tables(0)
        dgvCurrent.DataSource = dataTable

        If dataTable IsNot Nothing Then
            If dataTable.Rows.Count >= 1 Then
                ' Persist the data in the application settings.
                Dim xmlString As String = dataSet.GetXml()
                PersistToTableData(xmlString)
            End If
        End If

    End Sub

    ''' <summary>
    ''' Whenever the external data is updated, we store the latest version internally.
    ''' Later on if the external data is not available we can still use our last internal stored version.
    ''' </summary>
    ''' <param name="data"></param>
    Private Sub PersistToTableData(ByRef data As String)


        ' Serialize the data to a string.
        Dim serializer As New Xml.Serialization.XmlSerializer(GetType(String))
        Dim stringWriter As New IO.StringWriter()
        serializer.Serialize(stringWriter, data)
        Dim serializedData As String = stringWriter.ToString()

        ' Store the serialized data in the application settings.
        My.Settings.TableData = serializedData
        My.Settings.Save()

    End Sub

    ''' <summary>
    ''' Transform a datatable into a two-dimensional array.l
    ''' </summary>
    ''' <param name="dataTable"></param>
    ''' <returns></returns>
    Private Function ConvertTableToArray(dataTable) As String(,)
        Dim rowsCount As Integer = dataTable.Rows.Count
        Dim colsCount As Integer = dataTable.Columns.Count

        Dim data(rowsCount, colsCount - 1) As String

        For i As Integer = 0 To rowsCount - 1
            For j As Integer = 0 To colsCount - 1
                data(i, j) = dataTable.Rows(i)(j).ToString()
            Next
        Next

        Return data
    End Function

    ''' <summary>
    ''' Extract the Words of Concern from local persistant storage.
    ''' </summary>
    ''' <returns></returns>
    Private Function ExtractTableFromSettings() As String(,)
        ' Retrieve the serialized data from the application settings.
        Dim serializedData As String = My.Settings.TableData

        ' Deserialize the data back into a two-dimensional array.
        Dim serializer As New Xml.Serialization.XmlSerializer(GetType(String(,)))
        Dim stringReader As New IO.StringReader(serializedData)
        Dim data As String(,) = CType(serializer.Deserialize(stringReader), String(,))

        Return data
    End Function

    ''' <summary>
    '''    Populate the Word of Concern Datagrid view from a two dimensional array.
    ''' </summary>
    ''' <param name="dgvCurrent"></param>
    ''' <param name="data"></param>
    Private Sub ArrayToDataGrid(ByRef dgvCurrent As DataGridView, ByRef data As String(,))
        dgvCurrent.ColumnCount = data.GetLength(1)
        dgvCurrent.RowCount = data.GetLength(0)

        For i As Integer = 0 To data.GetLength(0) - 1
            For j As Integer = 0 To data.GetLength(1) - 1
                dgvCurrent(j, i).Value = data(i, j)
            Next
        Next

    End Sub

End Module
