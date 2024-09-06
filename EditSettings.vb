Imports System.Xml
Imports System.Windows.Forms
Imports System.IO
Imports System.Diagnostics

''' <summary>Dialog used to display and edit the three parts of ScriptAddin settings (Addin level, user and central)</summary>
Public Class EditSettings
    ''' <summary>the settings path for user or central setting (for re-saving after modification)</summary>
    Private settingsPath As String = ""

    ''' <summary>put the custom xml definition in the edit box for display/editing</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EditSettings_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' show ScriptAddin settings (user/central/addin level): set the appropriate config xml into EditBox (depending on Me.Tag)
        ' find path of xll:
        For Each tModule As ProcessModule In Process.GetCurrentProcess().Modules
            Dim sModule As String = tModule.FileName
            If sModule.ToUpper.Contains("SCRIPTADDIN") Then
                settingsPath = tModule.FileName
                Exit For
            End If
        Next
        ' read setting from xll path + ".config": addin level settings
        Me.Text = "ScriptAddin.xll.config settings in " + settingsPath
        Dim settingsStr As String
        Try
            settingsPath += ".config"
            settingsStr = File.ReadAllText(settingsPath, System.Text.Encoding.Default)
        Catch ex As Exception
            ScriptAddin.UserMsg("Couldn't read ScriptAddin.xll.config settings from " + settingsPath + ":" + ex.Message, True, True)
            Exit Sub
        End Try
        ' if central or user settings were chosen, overwrite settingsStr
        If Me.Tag = "central" Or Me.Tag = "user" Then
            ' get central settings filename from ScriptAddin.xll.config appSettings file attribute or
            ' get user settings filename from ScriptAddin.xll.config UserSettings configSource attribute 
            Dim doc As New XmlDocument()
            Dim xpathStr As String = If(Me.Tag = "central", "/configuration/appSettings/@file", "/configuration/UserSettings/@configSource")
            doc.LoadXml(settingsStr)
            If Not IsNothing(doc.SelectSingleNode(xpathStr)) Then
                Dim settingfilename As String = doc.SelectSingleNode(xpathStr).Value
                ' no path given in central filename: assume it is in same directory
                If InStr(settingfilename, "\") = 0 Then settingfilename = Replace(settingsPath, "ScriptAddin.xll.config", "") + settingfilename
                ' and read central settings
                settingsPath = settingfilename
                Try
                    settingsStr = File.ReadAllText(settingsPath, System.Text.Encoding.Default)
                Catch ex As Exception
                    ScriptAddin.UserMsg("Couldn't read Script Add-in " + Me.Tag + " settings from " + settingsPath + ":" + ex.Message, True, True)
                    Exit Sub
                End Try
                Me.Text = Me.Tag + " settings in " + settingsPath
            Else
                ScriptAddin.UserMsg("No attribute available as filename reference to " + Me.Tag + " settings (searched xpath: " + xpathStr + ") !", True, True)
                Exit Sub
            End If
        End If
        Me.OKBtn.Text = "Save"
        Me.ToolTip1.SetToolTip(OKBtn, "save " + Me.Text)
        Me.EditBox.Text = settingsStr
    End Sub

    ''' <summary>store the displayed/edited text box content back into the custom xml definition, including validation feedback</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OKBtn_Click(sender As Object, e As EventArgs) Handles OKBtn.Click
        ' save Add-in, users or central settings...
        Dim doc As New XmlDocument()
        Try
            ' validate settings
            Dim schemaString As String = My.Resources.SchemaFiles.DotNetConfig20
            If Me.Tag = "central" Then schemaString = My.Resources.SchemaFiles.ScriptAddinCentral
            If Me.Tag = "user" Then schemaString = My.Resources.SchemaFiles.ScriptAddinUser
            Dim schemadoc As XmlReader = XmlReader.Create(New StringReader(schemaString))
            doc.Schemas.Add("", schemadoc)
            Dim eventHandler As New Schema.ValidationEventHandler(AddressOf myValidationEventHandler)
            doc.LoadXml(Me.EditBox.Text)
            doc.Validate(eventHandler)
        Catch ex As Exception
            ScriptAddin.UserMsg("Problems with parsing changed " + Me.Tag + " settings: " + ex.Message, True, True)
            Exit Sub
        End Try
        Try
            File.WriteAllText(settingsPath, Me.EditBox.Text, System.Text.Encoding.UTF8)
        Catch ex As Exception
            ScriptAddin.UserMsg("Couldn't write " + Me.Tag + " settings into " + settingsPath + ": " + ex.Message, True, True)
            Exit Sub
        End Try
        If Me.Tag = "addin" Then ScriptAddin.UserMsg("Restart Addin (or Excel) to reflect changes in Addin settings.", True, False)
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    ''' <summary>validation handler for XML schema (user/app settings) checking</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Sub myValidationEventHandler(sender As Object, e As Schema.ValidationEventArgs)
        ' simply pass back Errors and Warnings as an exception
        If e.Severity = Schema.XmlSeverityType.Error Or e.Severity = Schema.XmlSeverityType.Warning Then Throw New Exception(e.Message)
    End Sub

    ''' <summary>no change was made to definition</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CancelBtn_Click(sender As Object, e As EventArgs) Handles CancelBtn.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>show the current line and column for easier detection of problems in xml document</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EditBox_SelectionChanged(sender As Object, e As EventArgs) Handles EditBox.SelectionChanged
        Me.PosIndex.Text = "Line: " + (Me.EditBox.GetLineFromCharIndex(Me.EditBox.SelectionStart) + 1).ToString() + ", Column: " + (Me.EditBox.SelectionStart - Me.EditBox.GetFirstCharIndexOfCurrentLine + 1).ToString()
    End Sub
End Class