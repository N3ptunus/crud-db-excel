Imports System.Data.OleDb

Public Class Form1
    Sub Kosongkan()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox1.Focus()
    End Sub

    Sub TampilGrid()
        DA = New OleDbDataAdapter("select * from [sheet1$]", CONN)
        DS = New DataSet
        DA.Fill(DS)
        DataGridView1.DataSource = DS.Tables(0)
        DataGridView1.ReadOnly = True
    End Sub
     Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call TampilGrid()
        Call Kosongkan()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 5
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            CMD = New OleDbCommand("Select * from [sheet1$] where val(Kode)='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                TextBox2.Text = ""
                TextBox2.Focus()
            Else
                TextBox2.Text = RD.Item("NAMA")
                TextBox2.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        TextBox2.MaxLength = 30
        If e.KeyChar = Chr(13) Then Button1.Focus()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data belum lengkap")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New OleDbCommand("Select * from [sheet1$] where val(Kode)='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                Dim simpan As String = "insert into [sheet1$] values ('" & TextBox1.Text & "','" & TextBox2.Text & "')"
                CMD = New OleDbCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
            Else
                Dim edit As String = "update [sheet1$] set NAMA='" & TextBox2.Text & "' where val(Kode)='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(edit, CONN)
                CMD.ExecuteNonQuery()
            End If

            Call TampilGrid()
            Call Kosongkan()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("val(Kode) Barang masih kosong, silakan diisi dulu")
            TextBox1.Focus()
            Exit Sub
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Dim hapus As String = "delete * from [sheet1$] where val(Kode)='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(hapus, CONN)
                CMD.ExecuteNonQuery()
                Call TampilGrid()
                Call Kosongkan()
            Else
                Call Kosongkan()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call Kosongkan()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        CMD = New OleDbCommand("select * from [sheet1$] where NAMA like '%" & TextBox3.Text & "%'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            DA = New OleDbDataAdapter("select * from [sheet1$] where NAMA like '%" & TextBox3.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS, "ketemu")
            DataGridView1.DataSource = DS.Tables("ketemu")
            DataGridView1.ReadOnly = True
        Else
            MsgBox("data tidak ditemukan")
        End If
    End Sub

End Class
