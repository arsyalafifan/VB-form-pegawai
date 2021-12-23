Imports System.Data
Imports System.Data.OleDb

Public Class Form1

    Sub Bersih()
        TxtNIP.Text = ""
        TxtNama.Text = ""
        TxtAlamat.Text = ""
        CmbBgn.Text = ""
        CmbPend.Text = ""
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CmbBgn.Items.Add("Manager")
        CmbBgn.Items.Add("Project Manager")
        CmbBgn.Items.Add("Supervisor")
        CmbBgn.Items.Add("Staff IT")
        CmbBgn.Items.Add("Software Engineer")

        CmbPend.Items.Add("SD")
        CmbPend.Items.Add("SMP")
        CmbPend.Items.Add("SMA")
        CmbPend.Items.Add("S1")
        CmbPend.Items.Add("S2")
        CmbPend.Items.Add("S3")
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
        Dim kawin As String

        If TxtNIP.Text = "" Or TxtNama.Text = "" Then
            MsgBox("Isi Data Dengan Benar", MsgBoxStyle.Exclamation, "Kesalahan")
        End If

        If RBSingle.Checked Then
            kawin = "Belum Kawin"
        Else
            kawin = "Kawin"
        End If

        CNN = New OleDbConnection(Koneksi)
        If CNN.State <> ConnectionState.Closed Then CNN.Close()
        CNN.Open()
        olecmd = New OleDbCommand("Insert into pegawai (nip, nama, bagian, tgl_lahir, alamat, pendidikan, status) values ('" & TxtNIP.Text & "','" & TxtNama.Text & "','" & CmbBgn.Text & "','" & DTtgl.Value & "','" & TxtAlamat.Text & "','" & CmbPend.Text & "','" & kawin & "')", CNN)

        x = olecmd.ExecuteNonQuery

        If x = 1 Then
            MsgBox("Data Berhasil Disimpan", MsgBoxStyle.Information, "Informasi")
            Call Bersih()
            TxtNIP.Focus()
        Else
            MsgBox("Salah Menyimpan Data", MsgBoxStyle.Information, "Informasi")
        End If

    End Sub

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles BtnExit.Click
        Close()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        Call Bersih()
        BtnSave.Enabled = True
        BtnEdit.Enabled = False
        BtnDelete.Enabled = False
        TxtNIP.Enabled = True
        TxtNIP.Focus()
    End Sub

    Private Sub TxtNIP_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtNIP.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TxtNama.Focus()
        End If
    End Sub

    Private Sub TxtNama_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtNama.KeyPress
        If Asc(e.KeyChar) = 13 Then
            CmbBgn.Focus()
        End If
    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        Dim kawin As String

        If TxtNIP.Text = "" Or TxtNama.Text = "" Then
            MsgBox("Isi Data Dengan Benar", MsgBoxStyle.Exclamation, "Kesalahan")
        End If

        If RBSingle.Checked Then
            kawin = "Belum Kawin"
        Else
            kawin = "Kawin"
        End If

        CNN = New OleDbConnection(Koneksi)
        If CNN.State <> ConnectionState.Closed Then CNN.Close()
        CNN.Open()
        olecmd = New OleDbCommand("Update pegawai Set nama='" & TxtNama.Text & "', bagian='" & CmbBgn.Text & "', tgl_lahir='" & DTtgl.Value & "', alamat='" & TxtAlamat.Text & "', pendidikan='" & CmbPend.Text & "', status='" & kawin & "' where NIP='" & TxtNIP.Text & "'", CNN)

        x = olecmd.ExecuteNonQuery

        If x = 1 Then
            MsgBox("Data Berhasil Diedit", MsgBoxStyle.Information, "Informasi")
            Call Bersih()
            TxtNIP.Focus()
        Else
            MsgBox("Salah Menyimpan Data", MsgBoxStyle.Information, "Informasi")
        End If
    End Sub

    Private Sub BtnCari_Click(sender As Object, e As EventArgs) Handles BtnCari.Click
        Dim poppegawai As New poppeg
        poppegawai.ShowDialog()
        If poppegawai.colNama <> "" Then
            TxtNIP.Text = poppegawai.colNIP
            TxtNama.Text = poppegawai.colNama
            CmbBgn.Text = poppegawai.colBgn
            DTtgl.Text = poppegawai.colTglLhr
            TxtAlamat.Text = poppegawai.colAlamat
            CmbPend.Text = poppegawai.colPend
            If poppegawai.colStatus = "Kawin" Then
                RBKawin.Checked = True
            Else
                RBKawin.Checked = False
            End If

            TxtNIP.Enabled = False
            TxtNama.Focus()

        End If
        BtnEdit.Enabled = True
        BtnDelete.Enabled = True
    End Sub

    Private Sub BtnReport_Click(sender As Object, e As EventArgs) Handles BtnReport.Click
        Report_Pegawai.Show()

    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As EventArgs) Handles BtnDelete.Click
        If MsgBox("Ingin menghapus data?", MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.Yes Then
            CNN = New OleDbConnection(Koneksi)
            If CNN.State <> ConnectionState.Closed Then CNN.Close()
            CNN.Open()
            olecmd = New OleDbCommand("delete from pegawai where nip='" & TxtNIP.Text & "'", CNN)
            x = olecmd.ExecuteNonQuery

            If x = 1 Then
                Call Bersih()
                BtnSave.Enabled = True
                BtnEdit.Enabled = False
                BtnDelete.Enabled = False
            Else
                MsgBox("Menghapus Data Gagal", MsgBoxStyle.Exclamation, "Kesalahan")
            End If
        End If
    End Sub
End Class
