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
            Dim args As String = ScriptAddin.ScriptExecArgs + " """ + ScriptAddin.fullScriptPath + "\" + ScriptAddin.script + """ " + ScriptAddin.scriptarguments
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
            ScriptAddin.UserMsg("Error occurred when invoking script '" + ScriptAddin.fullScriptPath + "\" + ScriptAddin.script + "', using '" + ScriptAddin.ScriptExec + "'" + ex.Message + vbCrLf, True, True)
            Me.errMsg = ex.Message
        End Try

    End Sub

    Private Sub myOutHandler(sender As Object, e As DataReceivedEventArgs)
        If IsNothing(e.Data) Then Exit Sub
        LogInfo("script out: " + e.Data)
        Dim appendAction As Action(Of String, System.Drawing.Color) = AddressOf appendTxt
        appendAction.Invoke(e.Data + vbCrLf, System.Drawing.Color.White)
    End Sub

    Private Sub myErrHandler(sender As Object, e As DataReceivedEventArgs)
        If IsNothing(e.Data) Then Exit Sub
        Dim appendAction As Action(Of String, System.Drawing.Color) = AddressOf appendTxt
        If ScriptAddin.StdErrMeansError Then Me.errMsg += e.Data
        LogWarn("script error: " + e.Data)
        appendAction.Invoke(e.Data + vbCrLf, System.Drawing.Color.Red)
    End Sub

    Private Sub myExitHandler(sender As Object, e As System.EventArgs)
        LogInfo("executed " + ScriptAddin.fullScriptPath)
        ' need this line to wait for stdout/stderr to finish writing...
        cmd.WaitForExit()
        If ScriptAddin.debugScript Then
            Dim appendAction As Action(Of String, System.Drawing.Color) = AddressOf appendTxt
            appendAction.Invoke("Finished script execution, exit code: " + cmd.ExitCode.ToString(), System.Drawing.Color.Yellow)
        Else
            Me.Hide()
        End If
    End Sub

    Private Sub appendTxt(theText As String, textCol As System.Drawing.Color)
        Dim pos As Integer = ScriptOutputTextbox.TextLength
        ScriptOutputTextbox.AppendText(theText)
        If textCol <> System.Drawing.Color.White Then
            ' select text to be colored
            ScriptOutputTextbox.Select(pos, theText.Length)
            ScriptOutputTextbox.SelectionColor = textCol
            ' deselect the text
            ScriptOutputTextbox.Select(pos + theText.Length, 0)
        End If
    End Sub

    Private Sub ScriptOutput_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If Not IsNothing(cmd) Then
            If e.KeyCode = Keys.Escape Then
                cmd.Close()
                Me.Hide()
            ElseIf e.KeyCode = Keys.Enter Then
                cmd.StandardInput.WriteLine()
            End If
        End If
    End Sub

    Private Sub ScriptOutput_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        cmd.StandardInput.Write(e.KeyChar)
    End Sub

    Private Sub ScriptOutput_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Not IsNothing(cmd) Then
            Dim procFinished As Boolean = False
            Try : procFinished = cmd.HasExited : Catch ex As Exception : End Try
            If procFinished Then
                cmd.CancelErrorRead()
                cmd.CancelOutputRead()
                cmd.Close()
            End If
        End If
    End Sub

End Class