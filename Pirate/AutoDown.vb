Imports Piratify.FreeMusic
Imports Piratify.frmMain
Imports System.IO
Imports System.Net
Imports System.Threading
Imports System.Web
Imports System.Text.RegularExpressions
Imports System.Text
Imports Excel = Microsoft.Office.Interop.Excel

Public Class AutoDown

    Private Sub AutoDown_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Down_Excel_Click2(sender As Object, e As EventArgs) Handles Down_Excel2.Click
        Dim songs As New List(Of AutoDownSong)
        Dim newPath As String
        Using dialog As New OpenFileDialog
            If dialog.ShowDialog() <> DialogResult.OK Then Return
            newPath = dialog.FileName
            MsgBox("Program will now try to download all the music! Don't stop if you think it is hanging, It's not oke:)")
            Dim index As Integer = 1
            Dim oApp As New Excel.Application()
            Dim oldCI As System.Globalization.CultureInfo = _
                System.Threading.Thread.CurrentThread.CurrentCulture
            System.Threading.Thread.CurrentThread.CurrentCulture = _
                New System.Globalization.CultureInfo("en-US")
            Dim bufString As String = "Init"
            Dim sheet As Excel.Sheets = oApp.Workbooks.Open(newPath).Sheets
            While (True)
                If (bufString.Equals("")) Then
                    Exit While
                End If

                Try
                    bufString = sheet.Item(1).Cells(index, 1).Value().ToString
                Catch
                    Exit While
                End Try
                Dim song As New AutoDownSong
                song.FileName = bufString
                songs.Add(song)
            End While
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Using


        Dim SearchIdMax As Integer = 0
        For Each song As FreeMusic.AutoDownSong In songs
            song.RowId = SearchIdMax

            Dim r As New XPTable.Models.Row()
            r.Cells.Add(New XPTable.Models.Cell(song.RowId))
            r.Cells.Add(New XPTable.Models.Cell("Download"))
            tmDown.Rows.Add(r)
            SearchIdMax += 1
        Next
        Table1.ScrollToTop()



    End Sub

End Class