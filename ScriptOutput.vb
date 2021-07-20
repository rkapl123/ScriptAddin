Imports System.Windows.Forms
Imports System.IO
Imports System.ComponentModel
Imports System.Diagnostics

''' <summary>Dialog for script invocation and output of standard out and standard err, allowing standard input as well</summary>
Public Class ScriptOutput
    Public cmd As Process
    Public errMsg As String = ""

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Try
            Dim args As String = ScriptAddin.ScriptExecArgs + " """ + ScriptAddin.fullScriptPath + "\" + ScriptAddin.script + """"
            Directory.SetCurrentDirectory(ScriptAddin.fullScriptPath)
            Dim pstartInfo = New ProcessStartInfo With {
                .FileName = ScriptAddin.ScriptExec,
                .Arguments = args,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .RedirectStandardInput = True,
                .CreateNoWindow = True,
                .UseShellExecute = False,
                .WorkingDirectory = ScriptAddin.fullScriptPath,
                .WindowStyle = ProcessWindowStyle.Hidden
            }

            pstartInfo.EnvironmentVariables.Item("PATH") = pstartInfo.EnvironmentVariables.Item("PATH") + ";" + ScriptAddin.ScriptExecAddPath

            cmd = New Process With {
                .StartInfo = pstartInfo,
                .EnableRaisingEvents = True
            }
            AddHandler cmd.OutputDataReceived, AddressOf myOutHandler
            AddHandler cmd.ErrorDataReceived, AddressOf myErrHandler
            AddHandler cmd.Exited, AddressOf myExitHandler

            Dim result As Boolean = cmd.Start()
            cmd.BeginOutputReadLine()
            cmd.BeginErrorReadLine()
        Catch ex As Exception
            ScriptAddin.myMsgBox("Error occured when invoking script '" + ScriptAddin.fullScriptPath + "\" + ScriptAddin.script + "', using '" + ScriptAddin.ScriptExec + "'" + ex.Message + vbCrLf, True, True)
            Me.errMsg = ex.Message
            Me.Hide()
        End Try

    End Sub

    Private Sub myOutHandler(sender As Object, e As DataReceivedEventArgs)
        Dim appendAction As Action(Of String, Boolean) = AddressOf appendTxt
        appendAction.Invoke(e.Data + vbCrLf, False)
    End Sub

    Private Sub myErrHandler(sender As Object, e As DataReceivedEventArgs)
        Dim appendAction As Action(Of String, Boolean) = AddressOf appendTxt
        LogWarn("scripterror: " + e.Data)
        Me.errMsg += e.Data
        appendAction.Invoke(e.Data + vbCrLf, True)
    End Sub

    Private Sub myExitHandler(sender As Object, e As System.EventArgs)
        LogInfo("executed " + ScriptAddin.fullScriptPath)
        Me.Text = "Script Output ....... Finished script execution, exit code: " + cmd.ExitCode.ToString()
        If Not ScriptAddin.debugScript Then Me.Hide()
    End Sub

    Private Sub appendTxt(theText As String, errCol As Boolean)
        Dim pos As Integer = ScriptOutputTextbox.TextLength
        ScriptOutputTextbox.AppendText(theText)
        If errCol Then
            ScriptOutputTextbox.Select(pos, theText.Length)
            ScriptOutputTextbox.SelectionColor = System.Drawing.Color.Red
            ScriptOutputTextbox.Select()
        End If
    End Sub

    Private Sub ScriptOutput_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If Not IsNothing(cmd) Then
            If e.KeyCode = Keys.Escape Then
                Me.Hide()
            ElseIf e.KeyCode = Keys.Enter Then
                cmd.StandardInput.Write(e.KeyCode)
                cmd.StandardInput.WriteLine()
            Else
                cmd.StandardInput.Write(e.KeyCode)
            End If
        End If
    End Sub

    Private Sub ScriptOutput_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Not IsNothing(cmd) AndAlso Not cmd.HasExited Then
            cmd.CancelErrorRead()
            cmd.CancelOutputRead()
            cmd.Close()
        End If
    End Sub

End Class