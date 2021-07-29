Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel
Imports System.Diagnostics
Imports System.Collections.Generic

''' <summary>Events from Addin (AutoOpen/Close) and Excel (Workbook_Save ...)</summary>
Public Class AddInEvents
    Implements IExcelAddIn

    ''' <summary>the Application object for event registration</summary>
    WithEvents Application As Application

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        Application = ExcelDnaUtil.Application
        theMenuHandler = New MenuHandler
        initScriptExecutables()
        Dim Wb As Workbook = Application.ActiveWorkbook
        If Not Wb Is Nothing Then
            Dim errStr As String = doDefinitions(Wb)
            ScriptAddin.dropDownSelected = False
            If errStr = "no ScriptAddinNames" Then
                ScriptAddin.resetScriptDefinitions()
            ElseIf errStr <> vbNullString Then
                ScriptAddin.UserMsg("Error when getting definitions in Workbook_Activate: " + errStr, True, True)
            End If
        End If
        Trace.Listeners.Add(New ExcelDna.Logging.LogDisplayTraceListener())
        ' get module info for path of xll (to get config there):
        For Each tModule As Diagnostics.ProcessModule In Diagnostics.Process.GetCurrentProcess().Modules
            ScriptAddin.UserSettingsPath = tModule.FileName
            If ScriptAddin.UserSettingsPath.ToUpper.Contains("SCRIPTADDIN") Then
                ScriptAddin.UserSettingsPath = Replace(UserSettingsPath, ".xll", "User.config")
                Exit For
            End If
        Next
    End Sub

    ''' <summary>clean up when Scriptaddin is deactivated</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        theMenuHandler = Nothing
        Application = Nothing
    End Sub

    ''' <summary>save arg ranges to text files as well </summary>
    Private Sub Workbook_Save(Wb As Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
        Dim errStr As String
        ' avoid resetting ScriptDefinition when dropdown selected for a specific ScriptDefinition !
        If ScriptAddin.dropDownSelected Then
            errStr = ScriptAddin.getScriptDefinitions()
            If errStr <> vbNullString Then ScriptAddin.UserMsg("Error while getScriptDefinitions (dropdown selected !) in Workbook_Save: " + errStr, True, True)
        Else
            errStr = doDefinitions(Wb) ' includes getScriptDefinitions - for top sorted ScriptDefinition
            If errStr = "no ScriptAddinNames" Then Exit Sub
            If errStr <> vbNullString Then
                ScriptAddin.UserMsg("Error when getting definitions in Workbook_Save: " + errStr, True, True)
                Exit Sub
            End If
        End If
        ScriptAddin.avoidFurtherMsgBoxes = False
        ScriptAddin.storeArgs()
        ScriptAddin.removeResultsDiags() ' remove results specified by rres
    End Sub

    ''' <summary>refresh ribbon is being treated in Workbook_Activate...</summary>
    Private Sub Workbook_Open(Wb As Workbook) Handles Application.WorkbookOpen
    End Sub

    ''' <summary>refresh ribbon with current workbook's ScriptAddin Names</summary>
    Private Sub Workbook_Activate(Wb As Workbook) Handles Application.WorkbookActivate
        Dim errStr As String = doDefinitions(Wb)
        ScriptAddin.dropDownSelected = False
        If errStr = "no ScriptAddinNames" Then
            ScriptAddin.resetScriptDefinitions()
        ElseIf errStr <> vbNullString Then
            ScriptAddin.UserMsg("Error when getting definitions in Workbook_Activate: " + errStr, True, True)
        End If
        ScriptAddin.theRibbon.Invalidate()
    End Sub

    ''' <summary>get ScriptAddin Names of current workbook and load ScriptDefinitions of first name in ScriptAddin Names</summary>
    Private Function doDefinitions(Wb As Workbook) As String
        Dim errStr As String
        ScriptAddin.currWb = Wb
        ' always reset ScriptDefinitions when changing Workbooks (may not be the current one, if saved programmatically!), otherwise this is not being refilled in getScriptNames
        ScriptDefinitionRange = Nothing
        ' get the defined ScriptAddin Names
        errStr = ScriptAddin.getScriptNames()
        If errStr = "no ScriptAddinNames" Then Return errStr
        If errStr <> vbNullString Then
            Return "Error while getScriptNames in doDefinitions: " + errStr
        End If
        ' get the definitions from the current defined range (first name in ScriptAddin Names)
        errStr = ScriptAddin.getScriptDefinitions()
        If errStr <> vbNullString Then Return "Error while getScriptDefinitions in doDefinitions: " + errStr
        LogInfo("done ScriptDefinitions for workbook " + Wb.Name)
        Return vbNullString
    End Function

    ''' <summary>Close Workbook: remove references to current Workbook and Script Definitions</summary>
    Private Sub Application_WorkbookDeactivate(Wb As Workbook) Handles Application.WorkbookDeactivate
        currWb = Nothing
        ScriptAddin.dropDownSelected = False
        ReDim Preserve Scriptcalldefnames(-1)
        ReDim Preserve Scriptcalldefs(-1)
        ScriptDefsheetColl = New Dictionary(Of String, Dictionary(Of String, Range))
        ScriptDefsheetMap = New Dictionary(Of String, String)
        ScriptAddin.resetScriptDefinitions()
        ScriptAddin.theRibbon.Invalidate()
    End Sub
End Class
