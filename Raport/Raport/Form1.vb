Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Runtime.InteropServices

Public Class Form1
    '~~> Define your Excel Objectsn
    Dim workbook As Excel.Workbook
    Dim reportsFolder As String
    Dim xlTmp As Excel.Application
    Dim kondisi As Boolean = False
    Dim curFile As String

    Sub getPath()
        reportsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Excel")
        xlTmp = New Excel.Application
    End Sub

    Sub getPathFile()
        reportsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Excel")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call getPath()
        curFile = reportsFolder & "\Siswa.xlsm"
        If File.Exists(curFile) Then
            xlTmp.Workbooks.Open(reportsFolder & "\Siswa.xlsm")
            xlTmp.Visible = True
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
            kondisi = True
        Else
            MsgBox("File Siswa.xlsm tidak bisa ditemukan!", MsgBoxStyle.Exclamation, "ANROnline | Peringatan")
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call getPath()
        curFile = reportsFolder & "\Kelas.xlsm"
        If File.Exists(curFile) Then
            xlTmp.Workbooks.Open(reportsFolder & "\Kelas.xlsm")
            xlTmp.Visible = True
            Button2.Enabled = False
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
            kondisi = True
        Else
            MsgBox("File Kelas.xlsm tidak bisa ditemukan!", MsgBoxStyle.Exclamation, "ANROnline | Peringatan")
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call getPath()
        curFile = reportsFolder & "\Mapel.xlsm"
        If File.Exists(curFile) Then
            xlTmp.Workbooks.Open(reportsFolder & "\Mapel.xlsm")
            xlTmp.Visible = True
            Button2.Enabled = False
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
            kondisi = True
        Else
            MsgBox("File Mapel.xlsm tidak bisa ditemukan!", MsgBoxStyle.Exclamation, "ANROnline | Peringatan")
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If kondisi = True Then
            Button2.Enabled = True
            Button1.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = True
            Button5.Enabled = True
            Button6.Enabled = True
            xlTmp.Workbooks.Close()
            kondisi = False
        Else
            MsgBox("Tidak ada file yang terbuka.", MsgBoxStyle.Exclamation, "ANROnline | Peringatan")
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        curFile = reportsFolder & "\Program_Keahlian.xlsm"
        Dim curFile1 As String = reportsFolder & "\Paket_Keahlian.xlsm"
        If File.Exists(curFile) Or File.Exists(curFile1) Then
            Dim jurusan = New Data_Jurusan
            jurusan.Show()
        Else
            MsgBox("File Program_Keahlian.xlsm atau Paket_Keahlian.xlsm tidak bisa ditemukan!", MsgBoxStyle.Exclamation, "ANROnline | Peringatan")
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call getPath()
        curFile = reportsFolder & "\Guru.xlsm"
        If File.Exists(curFile) Then
            xlTmp.Workbooks.Open(reportsFolder & "\Guru.xlsm")
            xlTmp.Visible = True
            Button2.Enabled = False
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
            kondisi = True
        Else
            MsgBox("File Guru.xlsm tidak bisa ditemukan!", MsgBoxStyle.Exclamation, "ANROnline | Peringatan")
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Call getPath()
        curFile = reportsFolder & "\Nilai.xlsm"
        If File.Exists(curFile) Then
            xlTmp.Workbooks.Open(reportsFolder & "\Nilai.xlsm")
            xlTmp.Visible = True
            Button2.Enabled = False
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
            kondisi = True
        Else
            MsgBox("File Nilai.xlsm tidak bisa ditemukan!", MsgBoxStyle.Exclamation, "ANROnline | Peringatan")
        End If
    End Sub

    Private Sub Form1_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Call getPathFile()
        curFile = reportsFolder & "\Siswa.xlsm"
        If File.Exists(curFile) Then
            Label11.Text = "Ditemukan"
            Label11.ForeColor = Color.Green
        Else
            Label11.Text = "Tidak Ditemukan"
            Label11.ForeColor = Color.Red
        End If

        curFile = reportsFolder & "\Kelas.xlsm"
        If File.Exists(curFile) Then
            Label12.Text = "Ditemukan"
            Label12.ForeColor = Color.Green
        Else
            Label12.Text = "Tidak Ditemukan"
            Label12.ForeColor = Color.Red
        End If

        curFile = reportsFolder & "\Guru.xlsm"
        If File.Exists(curFile) Then
            Label13.Text = "Ditemukan"
            Label13.ForeColor = Color.Green
        Else
            Label13.Text = "Tidak Ditemukan"
            Label13.ForeColor = Color.Red
        End If

        curFile = reportsFolder & "\Mapel.xlsm"
        If File.Exists(curFile) Then
            Label15.Text = "Ditemukan"
            Label15.ForeColor = Color.Green
        Else
            Label15.Text = "Tidak Ditemukan"
            Label15.ForeColor = Color.Red
        End If

        curFile = reportsFolder & "\Nilai.xlsm"
        If File.Exists(curFile) Then
            Label16.Text = "Ditemukan"
            Label16.ForeColor = Color.Green
        Else
            Label16.Text = "Tidak Ditemukan"
            Label16.ForeColor = Color.Red
        End If

        curFile = reportsFolder & "\Program_Keahlian.xlsm"
        Dim curFile1 As String = reportsFolder & "\Paket_Keahlian.xlsm"
        If File.Exists(curFile) Or File.Exists(curFile1) Then
            Label14.Text = "Ditemukan"
            Label14.ForeColor = Color.Green
        Else
            Label14.Text = "Tidak Ditemukan"
            Label14.ForeColor = Color.Red
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
