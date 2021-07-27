Imports System.Configuration
Imports System.Diagnostics
Imports System.IO

''' <summary>About box: used to provide information about version/buildtime and links for local help and project homepage</summary>
Public NotInheritable Class AboutBox1
    ''' <summary>flag for quitting excel after closing (set on CheckForUpdates_Click)</summary>
    Public quitExcelAfterwards As Boolean = False

    ''' <summary>set up Aboutbox</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Set the title of the form.
        Dim sModuleInfo As String = vbNullString
        For Each tModule As ProcessModule In Process.GetCurrentProcess().Modules
            Dim sModule As String = tModule.FileName
            If sModule.ToUpper.Contains("SCRIPTADDIN.XLL") Then
                sModuleInfo = FileDateTime(sModule).ToString()
            End If
        Next
        Me.LabelProductName.Text = "ScriptAddin Help"
        Me.Text = String.Format("About {0}", My.Application.Info.Title)
        Me.LabelVersion.Text = String.Format("Version {0}:{1}", My.Application.Info.Version.ToString, sModuleInfo)
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelCompanyName.Text = "Help and Sources on: " + My.Application.Info.CompanyName
        Me.TextBoxDescription.Text = My.Application.Info.Description
        checkForUpdate(False)
    End Sub

    ''' <summary>Close Aboutbox</summary>
    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As EventArgs) Handles OKButton.Click
        Me.Close()
    End Sub

    ''' <summary>Click on Project homepage: activate hyperlink in browser</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub LabelCompanyName_Click(sender As Object, e As EventArgs) Handles LabelCompanyName.Click
        Try
            Process.Start(My.Application.Info.CompanyName)
        Catch ex As Exception
            ScriptAddin.LogWarn(ex.Message)
        End Try
    End Sub

    ''' <summary>Click on Local help: activate hyperlink in browser</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub LabelProductName_Click(sender As Object, e As EventArgs) Handles LabelProductName.Click
        Try
            Process.Start(fetchSetting("LocalHelp", ""))
        Catch ex As Exception
            ScriptAddin.LogWarn(ex.Message)
        End Try
    End Sub

    ''' <summary>refresh ScriptDefinitions clicked: refresh all ScriptDefinitions in current workbook</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub refreshScriptDef_Click(sender As Object, e As EventArgs) Handles refreshScriptDef.Click
        ScriptAddin.initScriptExecutables()
        Dim errStr As String = ScriptAddin.startScriptNamesRefresh()
        If Len(errStr) > 0 Then
            ScriptAddin.UserMsg("refresh Error: " & errStr, True, True)
        Else
            If UBound(Scriptcalldefnames) = -1 Then
                ScriptAddin.UserMsg("no ScriptDefinitions found for ScriptAddin in current Workbook (3 column named range (type/value/path), minimum types: scriptexec and script)!", True)
            Else
                ScriptAddin.UserMsg("refreshed ScriptDefnames from current Workbook !", True)
            End If
        End If
    End Sub
    Private Sub CheckForUpdates_Click(sender As Object, e As EventArgs) Handles CheckForUpdates.Click
        checkForUpdate(True)
    End Sub


    ''' <summary>checks for updates of DB-Addin, asks for download and downloads them</summary>
    ''' <param name="doUpdate">only display result of check (false) or actually perform the update and download new version (true)</param>
    Public Sub checkForUpdate(doUpdate As Boolean)
        Const AddinName = "ScriptAddin-"
        Const updateFilenameZip = "downloadedVersion.zip"
        Dim localUpdateFolder As String = ScriptAddin.fetchSetting("localUpdateFolder", "")
        Dim localUpdateMessage As String = ScriptAddin.fetchSetting("localUpdateMessage", "A new version is available in the local update folder, after quitting Excel (is done next) start deployAddin.cmd to install it.")
        Dim updatesMajorVersion As String = ScriptAddin.fetchSetting("updatesMajorVersion", "1.0.0.")
        Dim updatesDownloadFolder As String = ScriptAddin.fetchSetting("updatesDownloadFolder", "C:\temp\")
        Dim updatesUrlBase As String = ScriptAddin.fetchSetting("updatesUrlBase", "https://github.com/rkapl123/ScriptAddin/archive/refs/tags/")
        Dim response As Net.HttpWebResponse = Nothing
        Dim urlFile As String = ""

        ' check for zip file of next higher revision
        Dim curRevision As Integer = My.Application.Info.Version.Revision
        ' try with highest possible Security protocol
        Try
            Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12 Or Net.SecurityProtocolType.SystemDefault
        Catch ex As Exception
            ScriptAddin.UserMsg("Error setting the SecurityProtocol: " + ex.Message())
            Exit Sub
        End Try

        ' always accept url certificate as valid
        Net.ServicePointManager.ServerCertificateValidationCallback = AddressOf ValidationCallbackHandler

        Do
            urlFile = updatesUrlBase + updatesMajorVersion + (curRevision + 1).ToString() + ".zip"
            Dim request As Net.HttpWebRequest
            Try
                request = Net.WebRequest.Create(urlFile)
                response = Nothing
                request.Method = "HEAD"
                response = request.GetResponse()
            Catch ex As Exception
            End Try
            If response IsNot Nothing Then
                curRevision += 1
                response.Close()
            End If
        Loop Until response Is Nothing
        ' get out if no newer version found
        If curRevision = My.Application.Info.Version.Revision Then
            Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "You have the latest version (" + updatesMajorVersion + curRevision.ToString() + ")."
            Me.TextBoxDescription.BackColor = Drawing.Color.FromKnownColor(Drawing.KnownColor.Control)
            Me.CheckForUpdates.Text = "no Update ..."
            Me.CheckForUpdates.Enabled = False
            Me.Refresh()
            Exit Sub
        Else
            Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "A new version (" + updatesMajorVersion + curRevision.ToString() + ") is available " + IIf(localUpdateFolder <> "", "in " + localUpdateFolder, "on github")
            Me.TextBoxDescription.BackColor = Drawing.Color.DarkOrange
            Me.CheckForUpdates.Text = "get Update ..."
            Me.CheckForUpdates.Enabled = True
            Me.Refresh()
            If Not doUpdate Then Exit Sub
        End If
        ' if there is a maintained local update folder, open it and let user update from there...
        If localUpdateFolder <> "" Then
            Try
                If ScriptAddin.QuestionMsg(localUpdateMessage, MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                    System.Diagnostics.Process.Start("explorer.exe", localUpdateFolder)
                    Me.quitExcelAfterwards = True
                    Me.Close()
                End If
            Catch ex As Exception
                ScriptAddin.UserMsg("Error when opening local update folder: " + ex.Message())
            End Try
            Exit Sub
        End If

        ' continue with download
        urlFile = updatesUrlBase + updatesMajorVersion + curRevision.ToString() + ".zip"

        ' create the download folder
        Try
            IO.Directory.CreateDirectory(updatesDownloadFolder)
        Catch ex As Exception
            ScriptAddin.UserMsg("Couldn't create file download folder (" + updatesDownloadFolder + "): " + ex.Message())
            Exit Sub
        End Try

        Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "Downloading new version from " + urlFile
        Me.Refresh()
        ' get the new version zip-file
        Dim requestGet As Net.HttpWebRequest = Net.WebRequest.Create(urlFile)
        requestGet.Method = "GET"
        Try
            response = requestGet.GetResponse()
        Catch ex As Exception
            ScriptAddin.UserMsg("Error when downloading new version: " + ex.Message())
            Exit Sub
        End Try
        ' save the version as zip file
        If response IsNot Nothing Then
            Dim receiveStream As Stream = response.GetResponseStream()
            Using downloadFile As IO.FileStream = File.Create(updatesDownloadFolder + updateFilenameZip)
                receiveStream.CopyTo(downloadFile)
            End Using
        End If
        response.Close()
        Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "Extracting " + urlFile + " to " + updatesDownloadFolder
        Me.Refresh()
        ' now extract the downloaded file and open the Distribution folder, first remove any existing folder...
        Try
            Directory.Delete(updatesDownloadFolder + AddinName + updatesMajorVersion + curRevision.ToString(), True)
        Catch ex As Exception : End Try
        Try
            Compression.ZipFile.ExtractToDirectory(updatesDownloadFolder + updateFilenameZip, updatesDownloadFolder)
        Catch ex As Exception
            ScriptAddin.UserMsg("Error when extracting new version: " + ex.Message())
        End Try
        Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "New version in " + updatesDownloadFolder + AddinName + updatesMajorVersion + curRevision.ToString() + "\Distribution, start deployAddin.cmd to install the new Version."
        Me.Refresh()
        Try
            System.Diagnostics.Process.Start("explorer.exe", updatesDownloadFolder + AddinName + updatesMajorVersion + curRevision.ToString() + "\Distribution")
        Catch ex As Exception
            ScriptAddin.UserMsg("Error when opening Distribution folder of new version: " + ex.Message())
        End Try
    End Sub

    Private Function ValidationCallbackHandler() As Boolean
        Return True
    End Function

End Class
