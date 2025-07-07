Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Vbe.Interop ' also need to add reference to Microsoft.Vbe.Interop.Forms, otherwise commandbuttons cb1 to cb0 won't work
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
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Public Shared WithEvents cb1 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb2 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb3 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb4 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb5 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb6 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb7 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb8 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb9 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb0 As Forms.CommandButton

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
                ScriptAddin.UserMsg("Error when getting definitions in Workbook_Activate: " + errStr, True, True)
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

    ''' <summary>specific click handler for the 1st definable command button</summary>
    Private Shared Sub cb1_Click() Handles cb1.Click
        cbClick(cb1.Name)
    End Sub
    ''' <summary>specific click handler for the 2nd definable command button</summary>
    Private Shared Sub cb2_Click() Handles cb2.Click
        cbClick(cb2.Name)
    End Sub
    ''' <summary>specific click handler for the 3rd definable command button</summary>
    Private Shared Sub cb3_Click() Handles cb3.Click
        cbClick(cb3.Name)
    End Sub
    ''' <summary>specific click handler for the 4th definable command button</summary>
    Private Shared Sub cb4_Click() Handles cb4.Click
        cbClick(cb4.Name)
    End Sub
    ''' <summary>specific click handler for the 5th definable command button</summary>
    Private Shared Sub cb5_Click() Handles cb5.Click
        cbClick(cb5.Name)
    End Sub
    ''' <summary>specific click handler for the 6th definable command button</summary>
    Private Shared Sub cb6_Click() Handles cb6.Click
        cbClick(cb6.Name)
    End Sub
    ''' <summary>specific click handler for the 7th definable command button</summary>
    Private Shared Sub cb7_Click() Handles cb7.Click
        cbClick(cb7.Name)
    End Sub
    ''' <summary>specific click handler for the 8th definable command button</summary>
    Private Shared Sub cb8_Click() Handles cb8.Click
        cbClick(cb8.Name)
    End Sub
    ''' <summary>specific click handler for the 9th definable command button</summary>
    Private Shared Sub cb9_Click() Handles cb9.Click
        cbClick(cb9.Name)
    End Sub
    ''' <summary>specific click handler for the 10th definable command button</summary>
    Private Shared Sub cb0_Click() Handles cb0.Click
        cbClick(cb0.Name)
    End Sub

    ''' <summary>common click handler for all command buttons</summary>
    ''' <param name="cbName">name of command button, defines whether a script is invoked (starts with Script_)</param>
    Private Shared Sub cbClick(cbName As String)
        Dim errStr As String
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

    ''' <summary>assign click handlers to command buttons in passed sheet Sh, maximum 10 buttons are supported</summary>
    ''' <param name="Sh">sheet where command buttons are located</param>
    Public Shared Function assignHandler(Sh As Object) As Boolean
        cb1 = Nothing : cb2 = Nothing : cb3 = Nothing : cb4 = Nothing : cb5 = Nothing : cb6 = Nothing : cb7 = Nothing : cb8 = Nothing : cb9 = Nothing : cb0 = Nothing
        assignHandler = True
        Dim collShpNames As String = ""
        Try
            For Each shp As Excel.Shape In Sh.Shapes
                ' only for OLE Control buttons...
                If shp.Type = MsoShapeType.msoOLEControlObject Then
                    ' Associate click-handler with all click events of the CommandButtons.
                    Dim ctrlName As String
                    Try : ctrlName = Sh.OLEObjects(shp.Name).Object.Name : Catch ex As Exception : ctrlName = "" : End Try
                    If Left(ctrlName, 7) = "Script_" Then
                        collShpNames += IIf(collShpNames <> "", ",", "") + shp.Name
                        If cb1 Is Nothing Then
                            cb1 = Sh.OLEObjects(shp.Name).Object
                        ElseIf cb2 Is Nothing Then
                            cb2 = Sh.OLEObjects(shp.Name).Object
                        ElseIf cb3 Is Nothing Then
                            cb3 = Sh.OLEObjects(shp.Name).Object
                        ElseIf cb4 Is Nothing Then
                            cb4 = Sh.OLEObjects(shp.Name).Object
                        ElseIf cb5 Is Nothing Then
                            cb5 = Sh.OLEObjects(shp.Name).Object
                        ElseIf cb6 Is Nothing Then
                            cb6 = Sh.OLEObjects(shp.Name).Object
                        ElseIf cb7 Is Nothing Then
                            cb7 = Sh.OLEObjects(shp.Name).Object
                        ElseIf cb8 Is Nothing Then
                            cb8 = Sh.OLEObjects(shp.Name).Object
                        ElseIf cb9 Is Nothing Then
                            cb9 = Sh.OLEObjects(shp.Name).Object
                        ElseIf cb0 Is Nothing Then
                            cb0 = Sh.OLEObjects(shp.Name).Object
                        Else
                            UserMsg("Only max. of 10 Script-Addin Buttons are allowed on a Worksheet, currently in use: " + collShpNames + " !")
                            assignHandler = False
                            Exit For
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            LogWarn("assignHandler exception occurred: " + ex.Message)
        End Try
    End Function

    ''' <summary>assign command buttons anew with each change of sheets</summary>
    ''' <param name="Sh"></param>
    Private Sub Application_SheetActivate(Sh As Object) Handles Application.SheetActivate
        ' only when needed assign button handler for this sheet ...
        If Not IsNothing(ScriptDefsheetColl) AndAlso Not IsNothing(ScriptDefsheetMap) AndAlso ScriptDefsheetColl.Count > 0 AndAlso ScriptDefsheetMap.Count > 0 Then assignHandler(Sh)
    End Sub

    ''' <summary>used for releasing com objects</summary>
    Protected Overrides Sub Finalize()
        LogInfo("Addin finalizing: Base finalize")
        MyBase.Finalize()
        LogInfo("Addin finalizing: releasing com objects of control buttons")
        If Not IsNothing(cb1) Then Marshal.ReleaseComObject(cb1)
        If Not IsNothing(cb2) Then Marshal.ReleaseComObject(cb2)
        If Not IsNothing(cb3) Then Marshal.ReleaseComObject(cb3)
        If Not IsNothing(cb4) Then Marshal.ReleaseComObject(cb4)
        If Not IsNothing(cb5) Then Marshal.ReleaseComObject(cb5)
        If Not IsNothing(cb6) Then Marshal.ReleaseComObject(cb6)
        If Not IsNothing(cb7) Then Marshal.ReleaseComObject(cb7)
        If Not IsNothing(cb8) Then Marshal.ReleaseComObject(cb8)
        If Not IsNothing(cb9) Then Marshal.ReleaseComObject(cb9)
        If Not IsNothing(cb0) Then Marshal.ReleaseComObject(cb0)
    End Sub
End Class
