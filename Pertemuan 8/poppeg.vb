Imports System.Data
Imports System.Data.OleDb

Public Class poppeg
    Public colNIP, colNama, colBgn, colTglLhr, colAlamat, colPend, colStatus, colEmail, colNoTelp, colAgama As String


    Dim cnn As OleDbConnection

    Dim cmmd As OleDbCommand


    Dim dReader As OleDbDataReader


    Private Sub ClearList()
        While Val(Counter.Text) > 0
            ListView1.Items(0).Remove()
            Counter.Text = Val(Counter.Text) - 1
        End While
    End Sub

    Private Sub Pilih()
        Try
            colNIP = ListView1.SelectedItems(0).SubItems(0).Text.ToString
            colNama = ListView1.SelectedItems(0).SubItems(1).Text.ToString
            colBgn = ListView1.SelectedItems(0).SubItems(2).Text.ToString
            colTglLhr = ListView1.SelectedItems(0).SubItems(3).Text.ToString
            colAlamat = ListView1.SelectedItems(0).SubItems(4).Text.ToString
            colPend = ListView1.SelectedItems(0).SubItems(5).Text.ToString
            colStatus = ListView1.SelectedItems(0).SubItems(6).Text.ToString
            colEmail = ListView1.SelectedItems(0).SubItems(7).Text.ToString
            colNoTelp = ListView1.SelectedItems(0).SubItems(8).Text.ToString
            colAgama = ListView1.SelectedItems(0).SubItems(9).Text.ToString
            Me.Close()
        Catch ex As Exception
            MsgBox("Pilih Salah Satu Data", MsgBoxStyle.Information)
        End Try
    End Sub

    Private Sub ListData()
        Call ClearList()

        Dim sqlx As String
        Dim x As Integer

        sqlx = "select nip, nama, bagian, tgl_lahir, alamat, pendidikan, status, email, no_telp, agama from pegawai where nama like '%" & Trim(TxtNama.Text) & "%'order by nama asc"
        cnn = New OleDbConnection(Koneksi)
        If cnn.State <> ConnectionState.Closed Then cnn.Close()
        cnn.Open()
        cmmd = New OleDbCommand(sqlx, cnn)
        dReader = cmmd.ExecuteReader

        Try
            While dReader.Read = True
                x = Val(Counter.Text)
                Counter.Text = Str(Val(Counter.Text) + 1) & " Record"

                With ListView1
                    .Items.Add("")
                    .Items(ListView1.Items.Count - 1).SubItems.Add("")
                    .Items(ListView1.Items.Count - 1).SubItems.Add("")
                    .Items(ListView1.Items.Count - 1).SubItems.Add("")
                    .Items(ListView1.Items.Count - 1).SubItems.Add("")
                    .Items(ListView1.Items.Count - 1).SubItems.Add("")
                    .Items(ListView1.Items.Count - 1).SubItems.Add("")
                    .Items(ListView1.Items.Count - 1).SubItems.Add("")
                    .Items(ListView1.Items.Count - 1).SubItems.Add("")
                    .Items(ListView1.Items.Count - 1).SubItems.Add("")
                    .Items(ListView1.Items.Count - 1).SubItems.Add("")
                    .Items(x).SubItems(0).Text = dReader.GetString(0)
                    .Items(x).SubItems(1).Text = dReader.GetString(1)
                    .Items(x).SubItems(2).Text = dReader.GetString(2)
                    .Items(x).SubItems(3).Text = dReader.GetDateTime(3)
                    .Items(x).SubItems(4).Text = dReader.GetString(4)
                    .Items(x).SubItems(5).Text = dReader.GetString(5)
                    .Items(x).SubItems(6).Text = dReader.GetString(6)
                    .Items(x).SubItems(7).Text = dReader.GetString(7)
                    .Items(x).SubItems(8).Text = dReader.GetString(8)
                    .Items(x).SubItems(9).Text = dReader.GetString(9)
                End With

            End While

        Finally
            dReader.Close()
        End Try
        cnn.Close()

    End Sub
    Private Sub poppeg_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call ListData()
    End Sub

    Private Sub BtnOk_Click(sender As Object, e As EventArgs) Handles BtnOk.Click
        Call Pilih()
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        Call Pilih()
    End Sub

    Private Sub TxtNama_TextChanged(sender As Object, e As EventArgs) Handles TxtNama.TextChanged
        Call ListData()
    End Sub

End Class