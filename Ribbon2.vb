Imports Microsoft.Office.Tools.Ribbon
Imports System.IO

Public Class Ribbon2
    Private file_ext As String = ".txt"

    Private Sub Ribbon2_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Function Get_ActiveBook_FullName() As String

        Get_ActiveBook_FullName = Globals.ThisAddIn.Application.ActiveWorkbook.FullName

    End Function

    Private Sub Format_Convert(ByRef filestr As String, ByRef fileToWrite As String, ByVal encoding As Encoding)
        Dim fws As FileStream
        Dim fw As StreamWriter

        fws = New FileStream(fileToWrite, FileMode.Create)
        fw = New StreamWriter(fws, encoding)
        fw.Write(filestr)
        fw.Dispose()
        If Not (fws Is Nothing) Then
            fws.Dispose()
            MsgBox(fileToWrite + " was created successfully.", MsgBoxStyle.Information, "Encoding Convertor")
        End If
    End Sub

    Private Sub Encoding_Convertor(ByVal encoding As Encoding, ByRef file_full_ext As String)
        Dim frs As FileStream
        Dim fr As StreamReader
        Dim filestr As String = ""

        Dim fileFullName As String = Get_ActiveBook_FullName()
        Dim slashPos As Int16 = fileFullName.LastIndexOf("\")

        If fileFullName Is Nothing Then
            OpenFileDialog1.FileName = "*.*"
        Else
            If slashPos > 0 Then
                OpenFileDialog1.InitialDirectory = fileFullName.Substring(0, slashPos)
                OpenFileDialog1.FileName = fileFullName.Substring(slashPos + 1)
            Else
                OpenFileDialog1.FileName = fileFullName
            End If
        End If

        If OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            frs = New FileStream(OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            fr = New StreamReader(frs)
            filestr = fr.ReadToEnd()
            fr.Dispose()
            Format_Convert(filestr, OpenFileDialog1.FileName + file_full_ext, encoding)
        End If
    End Sub


    Private Sub UTF16LEBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles UTF16LEBtn.Click
        Encoding_Convertor(System.Text.Encoding.Unicode, ".UTF16LE" + file_ext)
    End Sub

    Private Sub UTF16BEBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles UTF16BEBtn.Click
        Encoding_Convertor(System.Text.Encoding.BigEndianUnicode, ".UTF16BE" + file_ext)
    End Sub

    Private Sub UTF8Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles UTF8Btn.Click
        Encoding_Convertor(System.Text.Encoding.UTF8, ".UTF8" + file_ext)
    End Sub

    Private Sub SettingBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles SettingBtn.Click
        Dim str As String
        str = InputBox("Please give output file extension:", "Encoding Convertor", file_ext)

        If (str <> "" And str.IndexOf(".") = 0) Then
            file_ext = str
        Else
            MsgBox("Invalid input or extension is not changed.")
        End If

    End Sub
End Class
