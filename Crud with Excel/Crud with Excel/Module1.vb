Imports System.Data.OleDb

Module Module1

    Public CONN As OleDbConnection
    Public CMD As OleDbCommand
    Public DS As New DataSet
    Public DA As OleDbDataAdapter
    Public RD As OleDbDataReader
    Public DT As DataTable

    Sub Koneksi()
        'string koneksi ke data excel
        CONN = New OleDbConnection("provider=Microsoft.ace.OLEDB.12.0;data source=D:\N3PTUNUS\PKL\D_B.xlsx;Extended Properties=Excel 8.0;")
        CONN.Open()
    End Sub
End Module