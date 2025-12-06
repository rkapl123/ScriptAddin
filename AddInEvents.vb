Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Vbe.Interop ' also need to add reference to Microsoft.Vbe.Interop.Forms, otherwise commandbuttons won't work
Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core


''' <summary>Events from Addin (AutoOpen/Close) and Excel (Workbook_Save ...)</summary>
<ComVisible(True)>
Public Class AddInEvents
    Implements IExcelAddIn

    ''' <summary>the Application object for event registration</summary>
    WithEvents Application As Excel.Application
    ''' <summary>collection of command button handlers for assigned script actions</summary>
    Public Shared colCommandButtons As New Collection

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        Application = ExcelDnaUtil.Application
        theMenuHandler = New MenuHandler
        ' for finding out what happened attach internal trace to ExcelDNA LogDisplay
        ScriptAddin.theLogDisplaySource = New TraceSource("ExcelDna.Integration")

        initScriptExecutables()
        Dim Wb As Workbook = Application.ActiveWorkbook
        If Wb IsNot Nothing Then
            Dim errStr As String = doDefinitions(Wb)
            ScriptAddin.dropDownSelected = False
            If errStr = "no ScriptAddinNames" Then
                ScriptAddin.resetScriptDefinitions()
            ElseIf errStr <> vbNullString Then
                ScriptAddin.UserMsg("Error when getting definitions in AutoOpen: " + errStr, True, True)
            End If
        End If

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
    Private Sub Application_WorkbookBeforeSave(Wb As Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave
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
    Private Sub Application_WorkbookActivate(Wb As Workbook) Handles Application.WorkbookActivate
        Dim errStr As String = doDefinitions(Wb)
        ScriptAddin.dropDownSelected = False
        If errStr = "no ScriptAddinNames" Then
            ScriptAddin.resetScriptDefinitions()
        ElseIf errStr <> vbNullString Then
            ScriptAddin.UserMsg("Error when getting definitions in Workbook_Activate: " + errStr, True, True)
        End If
        InitializeCBHandlers(Wb)
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
        If errStr <> vbNullString Then Return "Error during getScriptDefinitions in doDefinitions: " + errStr
        LogInfo("done ScriptDefinitions for workbook " + Wb.Name)
        Return vbNullString
    End Function

    ''' <summary>Close Workbook: remove references to current Workbook and Script Definitions</summary>
    Private Sub Application_WorkbookDeactivate(Wb As Workbook) Handles Application.WorkbookDeactivate
        currWb = Nothing
        ScriptAddin.dropDownSelected = False
        Scriptcalldefnames = {}
        Scriptcalldefs = {}
        ScriptDefsheetColl = New Dictionary(Of String, Dictionary(Of String, Range))
        ScriptDefsheetMap = New Dictionary(Of String, String)
        ScriptAddin.resetScriptDefinitions()
        ScriptAddin.theRibbon.Invalidate()
    End Sub

    ''' <summary>assign click handlers to command buttons in passed workbook Wb</summary>
    ''' <param name="wb">Workbook where command buttons are located</param>
    Public Sub InitializeCBHandlers(wb As Object)
        Try
            For Each ws As Worksheet In wb.Worksheets
                For Each shp As Excel.Shape In ws.Shapes
                    ' only for OLE Control buttons...
                    If shp.Type = MsoShapeType.msoOLEControlObject Then
                        ' Associate click-event handler of a CommandButton if its name matches the DB modifiers name.
                        Dim ctrlName As String
                        Try : ctrlName = ws.OLEObjects(shp.Name).Object.Name : Catch ex As Exception : ctrlName = "" : End Try
                        If Left(ctrlName, 7) = "Script_" And Not colCommandButtons.Contains(wb.Name + ws.Name + ctrlName) Then
                            Dim cbCH As New CommandbuttonClickHandler With {.cb = ws.OLEObjects(shp.Name).Object}
                            colCommandButtons.Add(cbCH, wb.Name + ws.Name + ctrlName)
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            LogWarn("InitializeCBHandlers exception occurred: " + ex.Message)
        End Try
    End Sub

    ''' <summary>used for releasing com objects</summary>
    Protected Overrides Sub Finalize()
        LogInfo("Addin finalizing: Base finalize")
        MyBase.Finalize()
        LogInfo("Addin finalizing: releasing com objects of control buttons")
        For Each cbCH As CommandbuttonClickHandler In colCommandButtons
            Try : Marshal.ReleaseComObject(cbCH.cb) : Catch ex As Exception : End Try
        Next
        colCommandButtons.Clear()
        Try : Marshal.ReleaseComObject(ScriptAddin.currWb) : Catch ex As Exception : End Try
        Try : Marshal.ReleaseComObject(ScriptAddin.ScriptDefinitionRange) : Catch ex As Exception : End Try
        For Each scdrange As Range In ScriptAddin.Scriptcalldefs
            Try : Marshal.ReleaseComObject(scdrange) : Catch ex As Exception : End Try
        Next
    End Sub
End Class

''' <summary>Event handler class for click events on control buttons that are associated to script addin actions</summary>
Class CommandbuttonClickHandler
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range)</summary>
    Public WithEvents cb As Forms.CommandButton

    '''' <summary>common click handler for all command buttons</summary>
    Private Sub cb_Click() Handles cb.Click
        Dim errStr As String
        ' name of command button, defines whether a Script definition is invoked (starts with Script_)
        Dim cbName As String = cb.Name
        ' set ScriptDefinition to callers range
        Try
            ScriptAddin.ScriptDefinitionRange = ExcelDna.Integration.ExcelDnaUtil.Application.Range(cbName)
        Catch ex As Exception
            ScriptAddin.UserMsg("No range " + cbName + " (Script definitions) found !", True, True)
            Exit Sub
        End Try
        ScriptAddin.SkipScriptAndPreparation = My.Computer.Keyboard.CtrlKeyDown
        Dim origSelection As Range = ExcelDna.Integration.ExcelDnaUtil.Application.Selection
        Try
            ScriptAddin.ScriptDefinitionRange.Parent.Select()
        Catch ex As Exception
            ScriptAddin.UserMsg("Selection of worksheet of Script Definition Range not possible (probably because you're editing a cell)!", True, True)
        End Try
        ScriptAddin.ScriptDefinitionRange.Select()
        errStr = ScriptAddin.startScriptprocess()
        origSelection.Parent.Select()
        origSelection.Select()
        If errStr <> "" Then ScriptAddin.UserMsg(errStr, True, True)
    End Sub

End Class