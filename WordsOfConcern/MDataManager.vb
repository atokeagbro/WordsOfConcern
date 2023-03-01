Imports System.Windows.Forms

Module MDataManager

    'ReadOnly mSourceFilePath As String = "C:\Users\aoa82\OneDrive - ProsperSpark\ProsperSpark\BBG\Proof of Concepts\Words of Concern Base\PhrasesOfConcern.xlsx"
    ReadOnly mSourceFilePath As String = "C:\Users\aoa82\OneDrive - ProsperSpark\ProsperSpark\BBG\Proof of Concepts\Words of Concern Base\VBdotNET Tests\SourceDataMainTest.xlsx"
    ReadOnly mDBConn As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mSourceFilePath + ";Extended Properties=Excel 12.0;"
    ReadOnly mSQLSelectData As String = "SELECT * FROM [Sheet1$]"


    Friend Sub LoadSourceGrid(ByRef dgvCurrent As DataGridView)
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter

            MyConnection = New System.Data.OleDb.OleDbConnection(mDBConn)
            MyCommand = New System.Data.OleDb.OleDbDataAdapter(mSQLSelectData, MyConnection)

            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            dgvCurrent.DataSource = dataSet.Tables(0)

            MyConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub
End Module
