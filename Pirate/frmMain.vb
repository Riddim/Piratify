
Imports System.IO
Imports System.Net
Imports System.Threading
Imports System.Web
Imports Pirate.FreeMusic
Imports System.Text.RegularExpressions
Imports System.Text
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmMain

#Region "Main variables"

    Public WithEvents music As New FreeMusic
    Private Settings As New frmSettings
    Public songs As New List(Of FreeMusic.Song)
    Delegate Sub UpdateSearchDelegate(ByVal result As List(Of FreeMusic.Song))
    Delegate Sub UpdateSearchDelegate2()
    Delegate Sub UpdateProgressDelegate(ByVal song As FreeMusic.Song)
    Delegate Sub UpdateDownloadDelegate(ByVal info() As String)
    Private progress As Integer = 0
    Private isMouseDown As Boolean = False
    Private progressY As Integer = 0
    Private CurrentDownloads As New ArrayList
    Private DownloadIdMax As Integer = 0
    Private didCancel As Boolean
    Private searchString As String = ""
    Private SearchIdMax As Integer = 0
    Private SongsToFetch As Integer = 0

    Public counter As Integer = 0
    Public counter2 As Integer = 0
    Public counter3 As Integer = 0
    Public stringArray(1000) As String
    Public failedSongs As New List(Of String)
    Public downloadList As Boolean = False
#End Region

#Region "Data handling"

#Region "Search"

    Public Sub Search()
        If searchString <> txtSearch.Text Then
            searchString = txtSearch.Text
            Me.songs.Clear()
            tmSearch.Rows.Clear()
            SearchIdMax = 0
        End If
        pbProgress.Value = 0
        btnSearch.Text = "Searching.."
        btnSearch.Enabled = False
        didCancel = False
        Dim thread As New Thread(New ThreadStart(AddressOf SearchThread))
        thread.Start()
    End Sub

    Public Sub SearchThread()
        Try
            If Not music.IsLoggedIn And My.Settings.AuthCustom Then
                music.Login(My.Settings.AuthUser, My.Settings.AuthPass)
            ElseIf Not music.IsLoggedIn Then
                music.Login()
            End If
            Dim result As List(Of FreeMusic.Song) = music.Search(txtSearch.Text, SearchIdMax)
            songs.AddRange(result)

            Invoke(New UpdateSearchDelegate(AddressOf SearchCompleted), result)
        Catch ex As Exception
            'MsgBox("Could not search and parse data: " & vbCrLf & vbCrLf & ex.ToString, MsgBoxStyle.Exclamation, "An error occured")
            If (downloadList) Then
                Invoke(New UpdateSearchDelegate2(AddressOf SearchCompleted2))
            End If
        End Try
    End Sub

    Public Sub SearchCompleted(ByVal result As List(Of FreeMusic.Song))
        For Each song As FreeMusic.Song In result
            song.RowId = SearchIdMax
            Dim length As New TimeSpan(0, 0, song.Duration)
            Dim r As New XPTable.Models.Row()
            r.Cells.Add(New XPTable.Models.Cell(song.RowId))
            r.Cells.Add(New XPTable.Models.Cell(song.Quantity))
            r.Cells.Add(New XPTable.Models.Cell(song.Artist))
            r.Cells.Add(New XPTable.Models.Cell(song.Title))
            r.Cells.Add(New XPTable.Models.Cell(length.Minutes & ":" & length.Seconds.ToString.PadLeft(2, "0")))
            r.Cells.Add(New XPTable.Models.Cell(song.Bitrate))
            r.Cells.Add(New XPTable.Models.Cell(song.Size))
            r.Cells.Add(New XPTable.Models.Cell("Download"))
            tmSearch.Rows.Add(r)
            SearchIdMax += 1
        Next
        tblSearch.ScrollToTop()
        FetchDetails(result)

    End Sub

    Public Sub SearchCompleted2()
        counter3 += 1
        Amountfail.Text = counter3
        failedSongs.Add(txtSearch.Text)

        nextDownload()

    End Sub

#End Region

#Region "Details"

    Public Sub FetchDetails(ByVal songs As List(Of FreeMusic.Song))
        If songs.Count > 0 Then
            btnSearch.Text = "Fetching.."
            SongsToFetch = songs.Count
            progress = 0
            For i As Integer = 1 To songs.Count
                ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf FetchDetail), songs(i - 1))
                pbProgress.Value = Math.Round(i / songs.Count * 100)
            Next
        Else
            FinishSearch()
        End If
    End Sub

    Public Sub FetchDetail(ByVal song As FreeMusic.Song)
        Try
            If didCancel Then Exit Sub
            song = music.FetchDetail(song)
            Invoke(New UpdateProgressDelegate(AddressOf UpdateProgress), New Object() {song})
        Catch ex As Exception
        End Try
    End Sub

    Public Sub UpdateProgress(ByVal song As FreeMusic.Song)
        progress += 1

        For Each row As XPTable.Models.Row In tmSearch.Rows
            If row.Cells(0).Data = song.RowId Then
                row.Cells(5).Data = song.Bitrate
                row.Cells(6).Data = Math.Round(song.Size / 1024 / 1024, 2)
                Exit For
            End If
        Next

        If progress = SongsToFetch Then

            FinishSearch()
        End If
    End Sub

    Private Sub FinishSearch()

        btnSearch.Enabled = True
        tblSearch.Sort(tblSearch.SortingColumn, cmSearch.Columns(tblSearch.SortingColumn).SortOrder)
        pbProgress.Value = 100


        If downloadList Then


            Dim faultyletters As New List(Of String)
            faultyletters.Add("к")
            faultyletters.Add("ч")
            faultyletters.Add("т")
            faultyletters.Add("ь")
            faultyletters.Add("Ш")
            faultyletters.Add("Ф")
            faultyletters.Add("Ц")
            faultyletters.Add("И")
            faultyletters.Add("ђ")
            faultyletters.Add("preview")
            faultyletters.Add("vk.com")
            faultyletters.Add("ringtone")
            faultyletters.Add("realtones")
            faultyletters.Add("instrumental")
            faultyletters.Add("live ")
            faultyletters.Add("acapella")
            faultyletters.Add("mash ")
            faultyletters.Add("cover")
            faultyletters.Add("acoustic")
            faultyletters.Add("boosted ")
            faultyletters.Add("rip ")
            faultyletters.Add("processed")
            faultyletters.Add("continuous")

            If Not txtSearch.Text.ToLower().IndexOf("remix ") <> -1 Then
                faultyletters.Add("remix ")
            End If
            If Not txtSearch.Text.ToLower().IndexOf("edit ") <> -1 Then
                faultyletters.Add("edit ")
            End If
            If Not txtSearch.Text.ToLower().IndexOf("mix ") <> -1 Then
                faultyletters.Add("mix ")
            End If
            If Not txtSearch.Text.ToLower().IndexOf("bootleg") <> -1 Then
                faultyletters.Add("bootleg")
            End If

            Dim row As Object = 0
            Dim amount As Integer = tmSearch.Rows.Count
            Dim index As Integer = 0
            Dim indexsize = faultyletters.Count
            Do Until index = indexsize - 1
                If amount > row Then
                    Dim filename As String = tmSearch.Rows(row).Cells(2).Text & " - " & tmSearch.Rows(row).Cells(3).Text & ".mp3"
                    Dim size As String = tmSearch.Rows(row).Cells(4).Text
                    Dim size2 As String = Replace(size, ":", "")
                    Dim sizeint As Integer = Convert.ToDecimal(size2)
                    If (filename.ToLower.Contains(faultyletters.ElementAt(index).ToLower) OrElse sizeint < 120) Then
                        row += 1
                        index = 0
                    End If
                    index += 1
                Else
                    Exit Do
                End If

            Loop
            If amount > row Then
                StartDownload(row)
            Else
                failedSongs.Add(txtSearch.Text)
                Amountfail.Text = failedSongs.Count
            End If
            nextDownload()
        Else
            RenderSearchButton()
        End If

    End Sub

    Private Sub RenderSearchButton()
        If txtSearch.Text = searchString And searchString.Length > 0 Then
            btnSearch.Text = "More.."
        Else
            btnSearch.Text = "Search"
        End If
    End Sub

#End Region

#Region "Download"

    Private Sub StartDownload(ByVal row)
        Dim song As FreeMusic.Song = GetSongFromRow(row)
        If Not IsNothing(song) Then
            Dim filename As String = tmSearch.Rows(row).Cells(2).Text & " - " & tmSearch.Rows(row).Cells(3).Text & ".mp3"
            If My.Settings.JustDownload Then
                filename = My.Settings.DownloadDir & filename
            Else
                sfdDialog.FileName = filename
                sfdDialog.InitialDirectory = My.Settings.DownloadDir
                If sfdDialog.ShowDialog() = Windows.Forms.DialogResult.Cancel Then Exit Sub
                filename = sfdDialog.FileName
            End If

            If Not My.Settings.OverwriteFile Then
                Dim fileexists As Boolean = File.Exists(filename)
                Dim counter As Integer = 1
                Do While fileexists
                    filename = filename.Replace(If(counter = 1, ".mp3", " (" & (counter - 1) & ").mp3"), " (" & counter & ").mp3")
                    counter += 1
                    fileexists = File.Exists(filename)
                Loop
            End If
            Dim illegalChars As String = New String(Path.GetInvalidFileNameChars()) + New String(Path.GetInvalidPathChars())
            For Each pathChar As Char In filename
                filename.Replace(pathChar.ToString(), "")
            Next
            Download(song.Url, filename)
        End If
    End Sub

    Public Sub Download(ByVal url As String, ByVal file As String)
        Dim id As Integer = DownloadIdMax
        DownloadIdMax += 1
        Dim r As New XPTable.Models.Row
        r.Cells.Add(New XPTable.Models.Cell(id))
        r.Cells.Add(New XPTable.Models.Cell(file))
        r.Cells.Add(New XPTable.Models.Cell(0))
        r.Cells.Add(New XPTable.Models.Cell("0 KB/s"))
        tmDownload.Rows.Insert(0, r)
        CurrentDownloads.Add(New Integer() {id, 0})

        Dim info() As String = {url, file, id}
        ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf DownloadFile), info)
    End Sub

    Public Sub DownloadFile(ByVal info() As String)
        Try
            ' Make request
            Dim request As HttpWebRequest
            request = WebRequest.Create(info(0))
            request.Method = "GET"

            ' Get response
            Dim response As HttpWebResponse = request.GetResponse
            Dim length As Integer = response.ContentLength
            Dim responseStream As Stream = response.GetResponseStream
            Dim writeStream As New FileStream(info(1), FileMode.Create)

            ' Download file
            Dim speedTimer As New Stopwatch
            Dim totalRead As Integer = 0
            Dim readings As Integer = 0
            Dim speed As Double = 0
            Do
                speedTimer.Start()

                Dim readBytes(8191) As Byte
                Dim bytesread As Integer = responseStream.Read(readBytes, 0, 8192)
                totalRead += bytesread

                Dim tmp() As String = {info(2), totalRead, length, speed}
                Invoke(New UpdateDownloadDelegate(AddressOf UpdateDownload), New Object() {tmp})

                If bytesread = 0 Then Exit Do
                writeStream.Write(readBytes, 0, bytesread)

                speedTimer.Stop()

                For Each i() As Integer In CurrentDownloads
                    If i(0) = info(2) And i(1) = 1 Then
                        responseStream.Close()
                        writeStream.Close()
                        response.Close()
                        File.Delete(info(1))
                        Exit Sub
                    End If
                Next

                readings += 1
                If readings >= 10 Then
                    speed = 81920 / (speedTimer.ElapsedMilliseconds / 1000)
                    speedTimer.Reset()
                    readings = 0
                End If
            Loop
            responseStream.Close()
            writeStream.Close()
            response.Close()
        Catch ex As Exception
        End Try
    End Sub

    Public Sub UpdateDownload(ByVal info() As String)
        Dim p As Integer = CInt(Math.Round((info(1) / info(2)) * 100))
        For Each row As XPTable.Models.Row In tmDownload.Rows
            If row.Cells(0).Data = info(0) Then
                row.Cells(2).Data = p
                row.Cells(3).Text = If(progress >= 100, "Completed", Math.Round(CType(info(3), Double) / 1024, 2) & " KB/s")
                Exit For
            End If
        Next
    End Sub

#End Region

#End Region

#Region "Control methods"

    Private Sub pbProgress_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pbProgress.MouseDown
        progressY = pbProgress.PointToClient(Cursor.Position).Y
        isMouseDown = True
    End Sub

    Private Sub pbProgress_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pbProgress.MouseMove
        If isMouseDown Then
            Dim frmPoint As Integer = Me.PointToClient(Cursor.Position).Y
            SplitContainer1.SplitterDistance = frmPoint - progressY - 30
            pbProgress.Location = New Point(2, frmPoint - progressY)
        End If
    End Sub

    Private Sub pbProgress_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pbProgress.MouseUp
        isMouseDown = False
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Search()
    End Sub

    Private Sub txtSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyDown
        If e.KeyCode = Keys.Enter And btnSearch.Enabled Then
            Search()
        ElseIf e.KeyCode = Keys.Escape Then
            didCancel = True
            FinishSearch()
        End If
    End Sub

    Private Sub tblSearch_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tblSearch.DoubleClick
        If tblSearch.SelectedIndicies.Count = 1 Then
            StartDownload(tblSearch.SelectedIndicies(0))
        End If
    End Sub

    Private Sub tblSearch_CellButtonClicked(ByVal sender As Object, ByVal e As XPTable.Events.CellButtonEventArgs) Handles tblSearch.CellButtonClicked
        StartDownload(e.Row)
    End Sub

    Private Sub tblDownload_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tblDownload.DoubleClick
        If tblDownload.SelectedItems.Count = 1 Then
            Try
                System.Diagnostics.Process.Start(tblDownload.SelectedItems(0).Cells(1).Text)
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub btnSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSettings.Click
        Settings.Close()
        Settings = New frmSettings
        Settings.Show()
        Settings.Focus()
    End Sub

    Private Sub tblDownload_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tblDownload.KeyDown
        If e.KeyCode = Keys.Delete Then
            If tblDownload.SelectedItems.Count = 1 Then
                For Each i() As Integer In CurrentDownloads
                    If i(0) = tblDownload.SelectedItems(0).Cells(0).Data Then
                        i(1) = 1
                        tmDownload.Rows.RemoveAt(tblDownload.SelectedIndicies(0))
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        RenderSearchButton()
    End Sub


    Private Sub Down_Excel_Click(sender As Object, e As EventArgs) Handles Down_Excel.Click

        Dim newPath As String
        Using dialog As New OpenFileDialog
            If dialog.ShowDialog() <> DialogResult.OK Then Return
            newPath = dialog.FileName
            MsgBox("Program will now try to download all the music! Don't stop if you think it is hanging, It's not oke:)")
            txtSearch.Enabled = False
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
                stringArray(index - 1) = bufString
                index = index + 1
            End While
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Using


        fetchList()

    End Sub

#End Region

#Region "Helpers"

    Private Function GetSongFromRow(ByVal row As Integer) As FreeMusic.Song
        Return songs(tmSearch.Rows(row).Cells(0).Data)
    End Function

#End Region

#Region "Form events"

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        AutoUpdate.AutoUpdate()
        ThreadPool.SetMinThreads(12, 24)
        tblSearch.Sort(5, SortOrder.Descending)
    End Sub

    Private Sub frmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        cmSearch.Columns(1).Width = 40
        cmSearch.Columns(2).Width = Math.Round((tblSearch.Width - 314) / 2)
        cmSearch.Columns(3).Width = Math.Round((tblSearch.Width - 314) / 2)
        cmSearch.Columns(4).Width = 55
        cmSearch.Columns(5).Width = 55
        cmSearch.Columns(6).Width = 71
        cmSearch.Columns(7).Width = 70

        cmDownload.Columns(1).Width = Math.Round(tblDownload.Width - 350)
        cmDownload.Columns(2).Width = 252
        cmDownload.Columns(3).Width = 75

        pbProgress.Location = New Point(2, SplitContainer1.SplitterDistance + 30)
    End Sub

#End Region

#Region "LuukCode"

    Public Sub fetchList()
        counter3 = 0
        downloadList = True
        Down_Excel.Text = "Busy.."
        Down_Excel.Enabled = False
        nextDownload()
    End Sub

    Public Sub nextDownload()
        If (counter <> stringArray.Length) Then
            txtSearch.Text = stringArray(counter)
            counter += 1
            Search()
        Else
            downloadList = False
            Down_Excel.Text = "Down Excel"
            Down_Excel.Enabled = True
            txtSearch.Text = ""
            txtSearch.Enabled = True

            For Each row As XPTable.Models.Row In tmDownload.Rows
                If row.Cells(3).Text = "0 KB/s" Then
                    failedSongs.Add(row.Cells(1).Text)
                End If
            Next

            MsgBox("Choose a place and name for the file with failed song names! You can change the names in the file, and try again! Or just search manually!")


            Dim saveFileDialog1 As New SaveFileDialog
            saveFileDialog1.InitialDirectory = "Desktop"
            saveFileDialog1.Title = "Failed Songs"
            saveFileDialog1.CheckFileExists = False
            saveFileDialog1.CheckPathExists = False
            saveFileDialog1.DefaultExt = "txt"
            saveFileDialog1.Filter = "Text files (*.txt)|*.txt"
            saveFileDialog1.FilterIndex = 2
            saveFileDialog1.RestoreDirectory = True

            If (saveFileDialog1.ShowDialog() = DialogResult.OK) Then

                Dim fail As String
                ' Create or overwrite the file. 
                Dim fs As FileStream = File.Create(saveFileDialog1.FileName)
                For Each fail In failedSongs
                    ' Add text to the file. 
                    Dim info As Byte() = New UTF8Encoding(True).GetBytes(fail + System.Environment.NewLine)
                    fs.Write(info, 0, info.Length)
                Next

                fs.Close()
            End If


            counter = 0
            failedSongs.Clear()

        End If

    End Sub


#End Region


End Class

#Region "rubbish"
'    Dim excel_app As Excel.Application
'    Dim workbook As Excel.Workbook
'    Dim sheet As Excel.Worksheet
'    Dim vValue As Object

'    excel_app = New Excel.ApplicationClass
'    excel_app.Visible = True
'    excel_app.UserControl = True
'    Dim oldCI As System.Globalization.CultureInfo = _
'System.Threading.Thread.CurrentThread.CurrentCulture
'    System.Threading.Thread.CurrentThread.CurrentCulture = _
'        New System.Globalization.CultureInfo("en-US")
'    excel_app.Workbooks.Add()
'    System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

'    workbook = excel_app.Workbooks.Open("c:\Users\Luuk\Desktop\songs.xlsx")
'    sheet = workbook.Worksheets("Sheet1")

'    vValue = sheet.Cells(1, 1).Value       'Get the value from cell A1
'    MsgBox(vValue)
#End Region




