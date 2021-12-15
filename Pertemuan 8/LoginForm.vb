Imports System.Data.OleDb

Public Class LoginForm
    Private Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click
        CNN = New OleDbConnection(Koneksi)
        If CNN.State <> ConnectionState.Closed Then CNN.Close()
        CNN.Open()
        olecmd = New OleDbCommand("SELECT * FROM login WHERE username = '" & TxtUsername.Text & "' and password = '" & TxtPassword.Text & "'", CNN)

        olerdr = olecmd.ExecuteReader

        If (olerdr.Read()) Then
            Form1.Show()
            Me.Hide()
        Else
            MsgBox("Maaf, Username dan Password yang anda masukkan tidak terdaftar", MsgBoxStyle.OkOnly, "Login Gagal")
            TxtUsername.Text = ""
            TxtPassword.Text = ""
            TxtUsername.Focus()
        End If
    End Sub
End Class