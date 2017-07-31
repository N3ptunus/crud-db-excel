Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Runtime.InteropServices

Public Class Data_Jurusan
    Dim workbook As Excel.Workbook
    Dim reportsFolder As String
    Dim xlTmp As Excel.Application
    Dim kondisi As Boolean = False

    Sub getPath()
        reportsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Excel")
        xlTmp = New Excel.Application
    End Sub

    Sub getPathFile()
        reportsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Excel")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call getPath()
        If File.Exists(reportsFolder & "\Program_Keahlian.xlsm") Then
            xlTmp.Workbooks.Open(reportsFolder & "\Program_Keahlian.xlsm")
            xlTmp.Visible = True
            Button1.Enabled = False
            Button2.Enabled = False
            kondisi = True
        Else
            MsgBox("File Program_Keahlian.xlsm tidak bisa ditemukan!", MsgBoxStyle.Exclamation, "ANROnline | Peringatan")
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call getPath()
        If File.Exists(reportsFolder & "\Paket_Keahlian.xlsm") Then
            xlTmp.Workbooks.Open(reportsFolder & "\Paket_Keahlian.xlsm")
            xlTmp.Visible = True
            Button1.Enabled = False
            Button2.Enabled = False
            kondisi = True
        Else
            MsgBox("File Paket_Keahlian.xlsm tidak bisa ditemukan!", MsgBoxStyle.Exclamation, "ANROnline | Peringatan")
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If kondisi = True Then
            Button1.Enabled = True
            Button2.Enabled = True
            xlTmp.Workbooks.Close()
            kondisi = False
        Else
            MsgBox("Tidak ada file yang terbuka!", MsgBoxStyle.Exclamation, "ANROnline | Peringatan")
        End If
    End Sub

    Private Sub Data_Jurusan_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Call getPathFile()
        Dim curFile As String
        curFile = reportsFolder & "\Program_Keahlian.xlsm"
        If File.Exists(curFile) Then
            Label11.Text = "Ditemukan"
            Label11.ForeColor = Color.Green
        Else
            Label11.Text = "Tidak Ditemukan"
            Label11.ForeColor = Color.Red
        End If

        curFile = reportsFolder & "\Paket_Keahlian.xlsm"
        If File.Exists(curFile) Then
            Label12.Text = "Ditemukan"
            Label12.ForeColor = Color.Green
        Else
            Label12.Text = "Tidak Ditemukan"
            Label12.ForeColor = Color.Red
        End If
    End Sub
End Class