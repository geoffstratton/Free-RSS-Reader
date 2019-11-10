Imports System
Imports System.IO
Imports System.Net
Imports System.Xml
Imports System.Text.RegularExpressions
 
Public Class frmRSS
    Dim rssDirPath As String = "C:\Users\" & Environment.UserName & "\Documents\"
    Dim rssFilePath As String = rssDirPath & "rssFeed.txt"
 
    Private Sub rssForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ' Check for default text file on program start, quietly ignore if nonexistent
        Try
            Dim textIn As New StreamReader(
                New FileStream(rssFilePath, FileMode.Open, FileAccess.Read))
            buildLinksListFromTxtFile(textIn)
            textIn.Close()
        Catch ex As FileNotFoundException
        Catch ex As DirectoryNotFoundException
        Catch ex As IOException
        End Try
    End Sub
 
    Private Sub buildLinksListFromTxtFile(ByVal file As StreamReader)
        ' Read contents of RSS links text file, check for valid format
        Dim rssLinks As New List(Of String)
        Dim lineNumber As Integer = 0
 
        Do While file.Peek <> -1
            lineNumber = lineNumber + 1
            Dim row As String = file.ReadLine
            ' Hack to check if rssFeeds file is binary
            If row.Contains("\0\0") Or row.Contains("ÿ") Then
                MsgBox("This feeds list appears not to be a text file!", MsgBoxStyle.OkOnly)
                Exit Do
            ElseIf lineNumber = 1 And Not row.Contains("http") Then
                ' Hack to check that first line contains http; otherwise call it invalid text file
                MsgBox("This feeds list is invalid!", MsgBoxStyle.OkOnly)
                Exit Do
            Else
                rssLinks.Add(row)
            End If
        Loop
 
        ' Zero out cmbFeedList
        cmbFeedList.Items.Clear()
 
        For Each rssLink In rssLinks
            cmbFeedList.Items.Add(rssLink)
        Next
        file.Close()
    End Sub
 
    Private Sub btnFetch_Click(sender As System.Object, e As System.EventArgs) Handles btnFetch.Click
        ' Check for strings that are zero-length or without http protocol
        ' Set up some regular expressions to check for valid URLs
        Dim pattern As String = "^https?://[a-z0-9-]+(\.[a-z0-9-]+)+([/?].+)?$"
        Dim pattern2 As String = "^https?://www.[a-z0-9-]+(\.[a-z0-9-]+)+([/?].+)?$"
        Dim validURL As New Regex(pattern)
        Dim validURL2 As New Regex(pattern2)
        ' If URL contains www, it has to contain four parts: the protocol, www,
        ' the domain suffix, and some trailing characters, i.e., http://www.mysite.com/feed
        ' If URL doesn't contain www, it has to contain three parts: the protocol, domain
        ' suffix, and trailing characters, i.e., http://mysite.com/feed
        If ((cmbFeedList.Text.Contains("www") And validURL2.IsMatch(cmbFeedList.Text)) Or
            (Not cmbFeedList.Text.Contains("www") And validURL.IsMatch(cmbFeedList.Text))) Then
            ' Don't duplicate entries
            If Not cmbFeedList.Items.Contains(cmbFeedList.Text) Then
                cmbFeedList.Items.Add(cmbFeedList.Text)
            End If
            fetchRSS()
        Else
            MessageBox.Show("Please enter a valid RSS feed!", "Invalid RSS Feed", _
                MessageBoxButtons.OK)
        End If
    End Sub
 
    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) _
        Handles btnDelete.Click
        Dim deleteConfirm = MessageBox.Show("Are you sure you want to delete this feed?", _
               "Confirm feed deletion", MessageBoxButtons.YesNo)
        If deleteConfirm = Windows.Forms.DialogResult.Yes Then
            cmbFeedList.Items.Remove(cmbFeedList.Text)
        End If
    End Sub
 
    Private Sub rssForm_FormClosing(sender As System.Object, e As _
         System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        ' Write rss list to text file
        Dim saveConfirm = MessageBox.Show("Do you want to save your current feed _
            list? (If no, your existing rssFeeds.txt will be left untouched.)", "Save feed list", _
            MessageBoxButtons.YesNoCancel)
        If saveConfirm = Windows.Forms.DialogResult.Yes Then
            writeRSSFile()
        ElseIf saveConfirm = Windows.Forms.DialogResult.Cancel Then
            e.Cancel = True
        End If
    End Sub
 
    Private Sub fetchRSS()
        Dim rssURL = cmbFeedList.Text
        Dim rssFeed As Stream = Nothing
        Dim errorMsg As String = Nothing
 
        ' Set up an HTTP request
        Dim request As HttpWebRequest = CType(WebRequest.Create(rssURL), HttpWebRequest)
 
        ' Try the download, check for HTTP OK status, grab feed if successful
        Try
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            If response.StatusCode = HttpStatusCode.OK Then
                rssFeed = response.GetResponseStream()
                showRSS(rssFeed)
            End If
        Catch e As WebException
            errorMsg = "Download failed. The response from the server was: " +
                CType(e.Response, HttpWebResponse).StatusDescription
            MessageBox.Show(errorMsg, "Error", MessageBoxButtons.OK)
        Catch e As Exception
            errorMsg = "This doesn't look like an RSS feed. The specific error is: " + e.Message
            'errorMsg = "Hmm, there was a problem: " + e.Message
            MessageBox.Show(errorMsg, "Error", MessageBoxButtons.OK)
        End Try
    End Sub
 
    Private Overloads Sub writeRSSFile()
        ' Write rss list to text file - method with fixed path called on program close
        Dim textOut As New StreamWriter(
            New FileStream(rssFilePath, FileMode.Create, FileAccess.Write))
        For Each rssLink In cmbFeedList.Items
            textOut.WriteLine(rssLink.ToString)
        Next
        textOut.Close()
    End Sub
 
    Private Overloads Sub writeRSSFile(ByVal file As Integer)
        ' Write rss list to text file - method called from File menu with user-determined path
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.InitialDirectory = rssFilePath
        saveFileDialog1.Filter = "txt files (*.txt)|*.txt"
 
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim outputFile As StreamWriter = New StreamWriter(saveFileDialog1.OpenFile())
            If (outputFile IsNot Nothing) Then
                For Each rssLink In cmbFeedList.Items
                    outputFile.WriteLine(rssLink.ToString)
                Next
                outputFile.Close()
            End If
        End If
    End Sub
 
    Private Sub showRSS(ByVal rssStream As Stream)
        ' Process contents of RSS feed (XML)
 
        Dim rssFeed = XDocument.Load(rssStream)
        Dim output As String = Nothing
 
        ' Use XML literals to pull out the tags I want
        For Each post In From element In rssFeed...<item>
            output += "<h3>" + post.<title>.Value + "</h3>"
            ' Fix the date
            Dim correctDate = DateTime.Parse(post.<pubDate>.Value)
            output += "<strong>Posted on " + correctDate + "</strong>"
            '"<a href=""" + post.<link>.Value + " target=""_blank"">"
            output += post.<description>.Value
        Next
 
        ' Rewrite articles to open links in new window; replaced by navigating event override below
        ' Dim fixedOutput = output.Replace("<a rel=""nofollow""", "<a rel=""nofollow"" target=""_blank""")
        wbFeedList.DocumentText = "<html><body><font face=""sans-serif"">" + _
           output.ToString() + "</font></body></html>"
    End Sub
 
    Private Sub HelpToolStripMenuItemHelp_Click(sender As System.Object, e As _
        System.EventArgs) Handles HelpToolStripMenuItemHelp.Click
        rssHelpBox.Show()
    End Sub
 
    Private Sub AboutToolStripMenuItem_Click(sender As System.Object, e As _
        System.EventArgs) Handles AboutToolStripMenuItem.Click
        rssAboutBox.Show()
    End Sub
 
    Private Sub QuitToolStripMenuItem_Click(sender As System.Object, e As _
        System.EventArgs) Handles QuitToolStripMenuItem.Click
        writeRSSFile()
        Me.Close()
    End Sub
 
    Private Sub OpenFeedlistToolStripMenuItem_Click(sender As System.Object, e As _
        System.EventArgs) Handles OpenFeedlistToolStripMenuItem.Click
        ' Open feeds file from file menu
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.InitialDirectory = rssFilePath
        openFileDialog1.Filter = "txt files (*.txt)|*.txt"
        openFileDialog1.RestoreDirectory = True
 
        ' Offer to save existing feeds list first
        Dim saveOnClose = MsgBox("Would you like to save your existing feeds list ?", _
              MsgBoxStyle.YesNoCancel)
        If saveOnClose = MsgBoxResult.Yes Then
            writeRSSFile(1)
        ElseIf saveOnClose = MsgBoxResult.Cancel Then
            Return
        End If
 
        ' Open open file dialog
        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                Dim fileStream = New StreamReader(openFileDialog1.OpenFile)
                If (fileStream IsNot Nothing) Then
                    ' Open dialog
                    buildLinksListFromTxtFile(fileStream)
                End If
            Catch Ex As IOException
                MessageBox.Show("Cannot read file from disk. The error is: " & Ex.Message)
            End Try
        End If
    End Sub
 
    Private Sub SaveFeedlistToolStripMenuItem_Click(sender As System.Object, e As _
        System.EventArgs) Handles SaveFeedlistToolStripMenuItem.Click
        ' Save current file as txt
        If cmbFeedList.Items.Count > 0 Then
            writeRSSFile(1)
        Else
            MessageBox.Show("Add some RSS feeds before you try to save them!", "Error", _
            MessageBoxButtons.OK)
        End If
    End Sub
 
    Private Sub wbFeedList_Navigating(sender As Object, e As _
        System.Windows.Forms.WebBrowserNavigatingEventArgs) Handles wbFeedList.Navigating
        ' Open links in default browser instead of webbrowser control
        If Not (e.Url.ToString().Equals("about:blank", StringComparison.InvariantCultureIgnoreCase)) Then
            e.Cancel = True
            Process.Start(e.Url.ToString())
        End If
    End Sub
End Class