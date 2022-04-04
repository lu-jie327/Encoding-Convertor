Imports Microsoft.Office.Tools.Ribbon
Imports System.IO

Public Class Ribbon2

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

    Private Sub UTF16LEBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles UTF16LEBtn.Click
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
            'Set FileStream as FileShare.ReadWrite in case the file has been opened by another process. 
            frs = New FileStream(OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            fr = New StreamReader(frs)
            filestr = fr.ReadToEnd()
            fr.Dispose()
            Format_Convert(filestr, OpenFileDialog1.FileName + ".Unicode.txt", System.Text.Encoding.Unicode)

            'Dim fr1 As FileStream
            'fr1 = File.OpenRead(OpenFileDialog1.FileName)


            'Dim r As BinaryReader = New BinaryReader(fr1, System.Text.Encoding.Default)
            'Dim ss() As Byte = r.ReadBytes(3)
            'fr1.Close()

            'Dim encoding_str As String = ""

            'If (ss(0) >= 239) Then
            '    encoding_str = "utf-8"
            '    If (ss(0) = 254) And (ss(1) = 255) Then
            '        encoding_str = "utf-16 BE"
            '    ElseIf (ss(0) = 255) And (ss(1) = 254) Then
            '        encoding_str = "utf-16 LE"
            '    Else
            '        encoding_str = "default"
            '    End If
            'Else
            '    encoding_str = "default"
            'End If

            'MsgBox(encoding_str)

        End If
    End Sub

    Private Sub UTF16BEBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles UTF16BEBtn.Click
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
            Format_Convert(filestr, OpenFileDialog1.FileName + ".BEUnicode.txt", System.Text.Encoding.BigEndianUnicode)
        End If
    End Sub

    Private Sub UTF8Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles UTF8Btn.Click
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
            Format_Convert(filestr, OpenFileDialog1.FileName + ".UTF8.txt", System.Text.Encoding.UTF8)
        End If
    End Sub
End Class
