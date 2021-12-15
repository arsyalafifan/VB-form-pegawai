Imports System.Data.Odbc
Imports System.Data
Imports System.Data.OleDb

Module Module1
    Public x As Integer
    Public s As String = ""
    Public t As String = ""
    Public olecmd As OleDbCommand
    Public olerdr As OleDbDataReader
    Public CNN As OleDbConnection
    Public Koneksi As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" &
        Application.StartupPath & "/ADOGaji.accdb"
End Module
