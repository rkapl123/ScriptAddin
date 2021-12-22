Imports Microsoft.Office.Interop.Excel
Imports System.Collections.Generic
Imports System.Configuration
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports ExcelDna.Integration
Imports System.Diagnostics
Imports System.Collections.Specialized


''' <summary>The main functions for working with ScriptDefinitions (named ranges in Excel) and starting the Script processes (writing input, invoking scripts and retrieving results)</summary>
Public Module ScriptAddin
    ''' <summary>script type for calling scripts (could be R, perl, etc)</summary>
    Public ScriptType As String
    ''' <summary>executable name for calling scripts</summary>
    Public ScriptExec As String
    ''' <summary>optional arguments to executable for calling scripts</summary>
    Public ScriptExecArgs As String
    ''' <summary>optional additional path settings for ScriptExec</summary>
    Public ScriptExecAddPath As String
    ''' <summary>If Scriptengine writes to StdError, regard this as an error for further processing (some write to StdError in case of no error)</summary>
    Public StdErrMeansError As Boolean
    ''' <summary>for ScriptAddin invocations by executeScript, this is set to true, avoiding a MsgBox</summary>
    Public nonInteractive As Boolean = False
    ''' <summary>collect non interactive error messages here</summary>
    Public nonInteractiveErrMsgs As String
    ''' <summary>Debug the Addin: write trace messages</summary>
    Public DebugAddin As Boolean
    ''' <summary>The path where the User specific settings (overrides) can be found</summary>
    Public UserSettingsPath As String
    ''' <summary>skip script preparation and execution</summary>
    Public SkipScriptAndPreparation As Boolean
    ''' <summary>indicates an error in execution of DBModifiers, used for commit/rollback and for noninteractive message return</summary>
    Public hadError As Boolean

    ''' <summary>the current workbook, used for reference of all script related actions (only one workbook is supported to hold script definitions)</summary>
    Public currWb As Workbook
    ''' <summary>the current script definition range (three columns)</summary>
    Public ScriptDefinitionRange As Range
    ''' <summary></summary>
    Public Scriptcalldefnames As String() = {}
    ''' <summary></summary>
    Public Scriptcalldefs As Range() = {}
    ''' <summary></summary>
    Public ScriptDefsheetColl As Dictionary(Of String, Dictionary(Of String, Range))
    ''' <summary></summary>
    Public ScriptDefsheetMap As Dictionary(Of String, String)
    ''' <summary>reference object for the Addins ribbon</summary>
    Public theRibbon As CustomUI.IRibbonUI
    ''' <summary>ribbon menu handler</summary>
    Public theMenuHandler As MenuHandler
    ''' <summary></summary>
    Public avoidFurtherMsgBoxes As Boolean
    ''' <summary></summary>
    Public dirglobal As String
    ''' <summary>show the script output for debugging purposes (invisible otherwise)</summary>
    Public debugScript As Boolean
    ''' <summary>needed for workbook save (saves selected ScriptDefinition)</summary>
    Public dropDownSelected As Boolean
    ''' <summary>set to true if warning was issued, this flag indicates that the log button should get an exclamation sign</summary>
    Public WarningIssued As Boolean

    ''' <summary>definitions of current script invocations (scripts, args, results, diags...)</summary>
    Public ScriptDefDic As Dictionary(Of String, String()) = New Dictionary(Of String, String())
    ''' <summary>file suffix for currently selected ScriptType</summary>
    Private ScriptFileSuffix As String

    ''' <summary>startRprocess: started from GUI (button) and accessible from VBA (via Application.Run)</summary>
    ''' <returns>Error message or null string in case of success</returns>
    Public Function startScriptprocess() As String
        Dim errStr As String
        avoidFurtherMsgBoxes = False
        ' get the definition range
        errStr = getScriptDefinitions()
        If errStr <> vbNullString Then Return "Failed getting ScriptDefinitions: " + errStr
        If SkipScriptAndPreparation Then
            finishScriptprocess()
        Else
            Try
                If Not removeFiles() Then Return vbNullString
                If Not storeArgs() Then Return vbNullString
                If Not storeScriptRng() Then Return vbNullString
                If Not invokeScripts() Then Return vbNullString
            Catch ex As Exception
                Return "Exception in ScriptDefinitions preparation and execution: " + ex.Message + ex.StackTrace
            End Try
        End If
        ' all is OK = return nullstring
        Return vbNullString
    End Function

    ''' <summary>execute given ScriptDefName, used for VBA call by Application.Run</summary>
    ''' <param name="ScriptDefName">Name of Script Definition</param>
    ''' <param name="headLess">if set to true, ScriptAddin will avoid to issue messages and return messages in exceptions which are returned (headless)</param>
    ''' <returns>empty string on success, error message otherwise</returns>
    <ExcelCommand(Name:="executeScript")>
    Public Function executeScript(ScriptDefName As String, Optional headLess As Boolean = False) As String
        hadError = False : nonInteractive = headLess
        nonInteractiveErrMsgs = "" ' reset noninteractive messages
        Try
            ScriptAddin.ScriptDefinitionRange = ExcelDnaUtil.Application.ActiveWorkbook.Names.Item(ScriptDefName).RefersToRange
        Catch ex As Exception
            nonInteractive = False
            Return "No Script Definition Range (" + ScriptDefName + ") found in Active Workbook: " + ex.Message
        End Try
        LogInfo("Doing Script '" + ScriptDefName + "'.")
        Try
            currWb = ExcelDnaUtil.Application.ActiveWorkbook
            Dim errStr As String = ScriptAddin.getScriptNames()
            If errStr <> "" Then Throw New Exception("Error in ScriptAddin.getScriptNames: " + errStr)
            errStr = ScriptAddin.startScriptprocess()
            If errStr <> "" Then Throw New Exception("Error in ScriptAddin.startScriptprocess: " + errStr)
        Catch ex As Exception
            nonInteractive = False
            hadError = True
            Return "Script Definition '" + ScriptDefName + "' execution had following error(s): " + ex.Message
        End Try
        nonInteractive = False
        If hadError Then Return nonInteractiveErrMsgs
        Return "" ' no error, no message
    End Function

    ''' <summary>After all script invocations have finished, this is called to get the results and diagrams into the current excel workbook</summary>
    ''' <returns>Error message or null string in case of success</returns>
    Public Function finishScriptprocess() As String
        Try
            If Not getResults() Then Return vbNullString
            If Not getDiags() Then Return vbNullString
        Catch ex As Exception
            Return "Exception in ScriptDefinitions finalization: " + ex.Message + ex.StackTrace
        End Try
        ' all is OK = return nullstring
        Return vbNullString
    End Function

    ''' <summary>refresh ScriptNames from Workbook on demand (currently when invoking about box)</summary>
    ''' <returns>Error message or null string in case of success</returns>
    Public Function startScriptNamesRefresh() As String
        Dim errStr As String
        If currWb Is Nothing Then Return "No Workbook active to refresh ScriptNames from..."
        ' always reset ScriptDefinitions when refreshing, otherwise this is not being refilled in getRNames
        ScriptDefinitionRange = Nothing
        ' get the defined Script_/R_Addin Names
        errStr = getScriptNames()
        If errStr = "no ScriptAddinNames" Then
            Return vbNullString
        ElseIf errStr <> vbNullString Then
            Return "Error while getting Script in startScriptNamesRefresh: " + errStr
        End If
        theRibbon.Invalidate()
        Return vbNullString
    End Function

    ''' <summary>gets defined named ranges for script invocation in the current workbook</summary>
    ''' <returns>Error message or null string in case of success</returns>
    Public Function getScriptNames() As String
        ReDim Preserve Scriptcalldefnames(-1)
        ReDim Preserve Scriptcalldefs(-1)
        ScriptDefsheetColl = New Dictionary(Of String, Dictionary(Of String, Range))
        ScriptDefsheetMap = New Dictionary(Of String, String)
        Dim i As Integer = 0
        For Each namedrange As Name In currWb.Names
            Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
            If Left(cleanname, 7) = "Script_" Or Left(cleanname, 7) = "R_Addin" Then
                Dim prefix As String = Left(cleanname, 7)
                If InStr(namedrange.RefersTo, "#REF!") > 0 Then Return "ScriptDefinitions range " + namedrange.Parent.name + "!" + namedrange.Name + " contains #REF!"
                If namedrange.RefersToRange.Columns.Count <> 3 Then Return "ScriptDefinitions range " + namedrange.Parent.name + "!" + namedrange.Name + " doesn't have 3 columns !"
                ' final name of entry is without Script_/R_Addin and !
                Dim finalname As String = Replace(Replace(namedrange.Name, prefix, ""), "!", "")
                Dim nodeName As String = Replace(Replace(namedrange.Name, prefix, ""), namedrange.Parent.Name & "!", "")
                If nodeName = "" Then nodeName = "MainScript"
                ' first definition as standard definition (works without selecting a ScriptDefinition)
                If ScriptDefinitionRange Is Nothing Then ScriptDefinitionRange = namedrange.RefersToRange
                If Not InStr(namedrange.Name, "!") > 0 Then
                    finalname = currWb.Name + finalname
                End If
                ReDim Preserve Scriptcalldefnames(Scriptcalldefnames.Length)
                ReDim Preserve Scriptcalldefs(Scriptcalldefs.Length)
                Scriptcalldefnames(Scriptcalldefnames.Length - 1) = finalname
                Scriptcalldefs(Scriptcalldefs.Length - 1) = namedrange.RefersToRange

                Dim scriptColl As Dictionary(Of String, Range)
                If Not ScriptDefsheetColl.ContainsKey(namedrange.Parent.Name) Then
                    ' add to new sheet "menu"
                    scriptColl = New Dictionary(Of String, Range)
                    scriptColl.Add(nodeName, namedrange.RefersToRange)
                    ScriptDefsheetColl.Add(namedrange.Parent.Name, scriptColl)
                    ScriptDefsheetMap.Add("ID" + i.ToString(), namedrange.Parent.Name)
                    i += 1
                Else
                    ' add ScriptDefinition to existing sheet "menu"
                    scriptColl = ScriptDefsheetColl(namedrange.Parent.Name)
                    scriptColl.Add(nodeName, namedrange.RefersToRange)
                End If
            End If
        Next
        If UBound(Scriptcalldefnames) = -1 Then Return "no ScriptAddinNames"
        Return vbNullString
    End Function

    ''' <summary>reset all ScriptDefinition representations</summary>
    Public Sub resetScriptDefinitions()
        ScriptDefDic("args") = {}
        ScriptDefDic("argspaths") = {}
        ScriptDefDic("results") = {}
        ScriptDefDic("rresults") = {}
        ScriptDefDic("resultspaths") = {}
        ScriptDefDic("diags") = {}
        ScriptDefDic("diagspaths") = {}
        ScriptDefDic("scripts") = {}
        ScriptDefDic("skipscripts") = {}
        ScriptDefDic("scriptspaths") = {}
        ScriptDefDic("scriptrng") = {}
        ScriptDefDic("scriptrngpaths") = {}
        ScriptExec = Nothing : dirglobal = vbNullString
    End Sub

    ''' <summary>gets definitions from current selected script invocation range (ScriptDefinitions)</summary>
    ''' <returns>Error message or null string in case of success</returns>
    Public Function getScriptDefinitions() As String
        resetScriptDefinitions()
        Try
            ScriptExecArgs = "" ' reset ScriptExec arguments as they might have been set elsewhere...
            ScriptExec = Nothing ' same for ScriptExec and other settings
            ScriptExecAddPath = ""
            ScriptFileSuffix = ""
            StdErrMeansError = True
            For Each defRow As Range In ScriptDefinitionRange.Rows
                Dim deftype As String, defval As String, deffilepath As String
                deftype = LCase(defRow.Cells(1, 1).Value2)
                defval = defRow.Cells(1, 2).Value2
                defval = If(defval = vbNullString, "", defval)
                deffilepath = defRow.Cells(1, 3).Value2
                deffilepath = If(deffilepath = vbNullString, "", deffilepath)
                If (deftype = "exec" Or deftype = "rexec") Then
                    If defval <> "" Then
                        ScriptExec = defval
                        ScriptExecArgs = deffilepath
                    End If
                ElseIf deftype = "skipscript" Or deftype = "script" Then
                    If defval <> "" Then
                        ReDim Preserve ScriptDefDic("scripts")(ScriptDefDic("scripts").Length)
                        ScriptDefDic("scripts")(ScriptDefDic("scripts").Length - 1) = defval
                        ReDim Preserve ScriptDefDic("scriptspaths")(ScriptDefDic("scriptspaths").Length)
                        ScriptDefDic("scriptspaths")(ScriptDefDic("scriptspaths").Length - 1) = deffilepath
                        ReDim Preserve ScriptDefDic("skipscripts")(ScriptDefDic("skipscripts").Length)
                        ScriptDefDic("skipscripts")(ScriptDefDic("skipscripts").Length - 1) = (deftype = "skipscript")
                    End If
                ElseIf deftype = "path" And defval <> "" Then
                    If defval <> "" Then
                        ScriptExecAddPath = defval
                        ScriptFileSuffix = deffilepath
                    End If
                ElseIf deftype = "type" Then
                    If ScriptExecutables.Contains(defval) Then
                        ScriptType = defval
                        theMenuHandler.selectedScriptExecutable = ScriptExecutables.IndexOf(ScriptType)
                        theRibbon.InvalidateControl("execDropDown")
                        StdErrMeansError = Not (deffilepath.ToLower() = "n" Or deffilepath.ToLower() = "no")
                    Else
                        Return "Error in setting type, not contained in available types/executables (check AppSettings for available ExePath<> entries)!"
                    End If
                ElseIf deftype = "arg" Then
                    ReDim Preserve ScriptDefDic("args")(ScriptDefDic("args").Length)
                    ScriptDefDic("args")(ScriptDefDic("args").Length - 1) = defval
                    ReDim Preserve ScriptDefDic("argspaths")(ScriptDefDic("argspaths").Length)
                    ScriptDefDic("argspaths")(ScriptDefDic("argspaths").Length - 1) = deffilepath
                ElseIf deftype = "scriptrng" Or deftype = "scriptcell" Then
                    ReDim Preserve ScriptDefDic("scriptrng")(ScriptDefDic("scriptrng").Length)
                    ScriptDefDic("scriptrng")(ScriptDefDic("scriptrng").Length - 1) = IIf(Right(deftype, 4) = "cell", "=", "") + defval
                    ReDim Preserve ScriptDefDic("scriptrngpaths")(ScriptDefDic("scriptrngpaths").Length)
                    ScriptDefDic("scriptrngpaths")(ScriptDefDic("scriptrngpaths").Length - 1) = deffilepath
                    ' don't set skipscripts here to False as this is done in method storeScriptRng
                ElseIf deftype = "res" Or deftype = "rres" Then
                    ReDim Preserve ScriptDefDic("rresults")(ScriptDefDic("rresults").Length)
                    ScriptDefDic("rresults")(ScriptDefDic("rresults").Length - 1) = (deftype = "rres")
                    ReDim Preserve ScriptDefDic("results")(ScriptDefDic("results").Length)
                    ScriptDefDic("results")(ScriptDefDic("results").Length - 1) = defval
                    ReDim Preserve ScriptDefDic("resultspaths")(ScriptDefDic("resultspaths").Length)
                    ScriptDefDic("resultspaths")(ScriptDefDic("resultspaths").Length - 1) = deffilepath
                ElseIf deftype = "diag" Then
                    ReDim Preserve ScriptDefDic("diags")(ScriptDefDic("diags").Length)
                    ScriptDefDic("diags")(ScriptDefDic("diags").Length - 1) = defval
                    ReDim Preserve ScriptDefDic("diagspaths")(ScriptDefDic("diagspaths").Length)
                    ScriptDefDic("diagspaths")(ScriptDefDic("diagspaths").Length - 1) = deffilepath
                ElseIf deftype = "dir" Then
                    dirglobal = defval
                ElseIf deftype <> "" Then
                    Return "Error in getScriptDefinitions: invalid type '" + deftype + "' found in script definition!"
                End If
            Next
            ' get default ScriptExec path from user (or overriden in appSettings tag as redirect to global) settings. This can be overruled by individual script exec settings in ScriptDefinitions
            If ScriptExec Is Nothing Then ScriptExec = fetchSetting("ExePath" + ScriptType, "")
            If ScriptExecAddPath = "" Then ScriptExecAddPath = fetchSetting("PathAdd" + ScriptType, "")
            If ScriptFileSuffix = "" Then ScriptFileSuffix = fetchSetting("FSuffix" + ScriptType, "")
            If ScriptExecArgs = "" Then ScriptExecArgs = fetchSetting("ExeArgs" + ScriptType, "")
            If ScriptExec = "" Then Return "Error in getScriptDefinitions: ScriptExec not defined (check AppSettings for available ExePath<> entries)"
            If ScriptDefDic("scripts").Length = 0 And ScriptDefDic("scriptrng").Length = 0 Then Return "Error in getScriptDefinitions: no script(s) or scriptRng(s) defined in " + ScriptDefinitionRange.Name.Name
            If StdErrMeansError Then StdErrMeansError = CBool(fetchSetting("StdErrX" + ScriptType, "True"))
        Catch ex As Exception
            Return "Error in getScriptDefinitions: " + ex.Message
        End Try
        Return vbNullString
    End Function

    ''' <summary>remove results in all result Ranges (before saving)</summary>
    Public Sub removeResultsDiags()
        For Each namedrange As Name In currWb.Names
            If Left(namedrange.Name, 15) = "___ScriptResult" Or Left(namedrange.Name, 15) = "___RaddinResult" Then
                namedrange.RefersToRange.ClearContents()
                namedrange.Delete()
            End If
        Next
    End Sub

    ''' <summary>prepare parameter (script, args, results, diags) for usage in invokeScripts, storeArgs, getResults and getDiags</summary>
    ''' <param name="index">index of parameter to be prepared in ScriptDefDic</param>
    ''' <param name="name">name (type) of parameter: script, scriptrng, args, results, diags</param>
    ''' <param name="ScriptDataRange">returned Range of data area for scriptrng, args, results and diags</param>
    ''' <param name="returnName">returned name of data file for the parameter: same as range name</param>
    ''' <param name="returnPath">returned path of data file for the parameter</param>
    ''' <param name="ext">extension of filename that should be used for file containing data for that type (e.g. txt for args/results or png for diags)</param>
    ''' <returns>True if success, False otherwise</returns>
    Private Function prepareParam(index As Integer, name As String, ByRef ScriptDataRange As Range, ByRef returnName As String, ByRef returnPath As String, ext As String) As String
        Dim value As String = ScriptDefDic(name)(index)
        If value = "" Then Return "Empty definition value for parameter " + name + ", index: " + index.ToString()
        ' allow for other extensions than txt if defined in ScriptDefDic(name)(index)
        If InStr(value, ".") > 0 Then ext = ""
        ' only for args, results and diags (scripts dont have a target range)
        Dim ScriptDataRangeAddress As String = ""
        If name = "args" Or name = "results" Or name = "diags" Or name = "scriptrng" Then
            Try
                ScriptDataRange = currWb.Names.Item(value).RefersToRange
                ScriptDataRangeAddress = ScriptDataRange.Parent.Name + "!" + ScriptDataRange.Address
            Catch ex As Exception
                Return "Error occured when looking up " + name + " range '" + value + "' in Workbook " + currWb.Name + " (defined correctly ?), " + ex.Message
            End Try
        End If
        ' if argvalue refers to a WS Name, cut off WS name prefix for Script file name...
        Dim posWSseparator = InStr(value, "!")
        If posWSseparator > 0 Then
            value = value.Substring(posWSseparator)
        End If
        ' get path of data file, if it is defined
        If ScriptDefDic.ContainsKey(name + "paths") Then
            If Len(ScriptDefDic(name + "paths")(index)) > 0 Then
                returnPath = ScriptDefDic(name + "paths")(index)
            End If
        End If
        returnName = value + ext
        LogInfo("prepared param in index:" + index.ToString() + ",type:" + name + ",returnName:" + returnName + ",returnPath:" + returnPath + IIf(ScriptDataRangeAddress <> "", ",ScriptDataRange: " + ScriptDataRangeAddress, ""))
        Return vbNullString
    End Function

    ''' <summary>creates Inputfiles for defined arg ranges, tab separated, decimalpoint always ".", dates are stored as "yyyy-MM-dd"
    ''' otherwise: "what you see is what you get"
    '''</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function storeArgs() As Boolean
        Dim argFilename As String = vbNullString, argdir As String
        Dim ScriptDataRange As Range = Nothing
        Dim outputFile As StreamWriter = Nothing

        argdir = dirglobal
        For c As Integer = 0 To ScriptDefDic("args").Length - 1
            Try
                Dim errMsg As String
                errMsg = prepareParam(c, "args", ScriptDataRange, argFilename, argdir, ".txt")
                If Len(errMsg) > 0 Then
                    If Not ScriptAddin.UserMsg(errMsg) Then Return False
                End If

                ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
                Dim curWbPrefix As String = IIf(Left(argdir, 2) = "\\" Or Mid(argdir, 2, 2) = ":\", "", currWb.Path + "\")
                outputFile = New StreamWriter(curWbPrefix + argdir + "\" + argFilename)
                ' make sure we're writing a dot decimal separator
                Dim customCulture As System.Globalization.CultureInfo
                customCulture = System.Threading.Thread.CurrentThread.CurrentCulture.Clone()
                customCulture.NumberFormat.NumberDecimalSeparator = "."
                System.Threading.Thread.CurrentThread.CurrentCulture = customCulture

                ' write the ScriptDataRange to file
                Dim i As Integer = 1
                Do
                    Dim j As Integer = 1
                    Dim writtenLine As String = ""
                    If ScriptDataRange(i, 1).Value2 IsNot Nothing Then
                        Do
                            Dim printedValue As String
                            If ScriptDataRange(i, j).NumberFormat.ToString().Contains("yy") Then
                                printedValue = DateTime.FromOADate(ScriptDataRange(i, j).Value2).ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)
                            ElseIf IsNumeric(ScriptDataRange(i, j).Value2) Then
                                printedValue = String.Format("{0:###################0.################}", ScriptDataRange(i, j).Value2)
                            Else
                                printedValue = ScriptDataRange(i, j).Value2
                            End If
                            writtenLine += printedValue + vbTab
                            j += +1
                        Loop Until j > ScriptDataRange.Columns.Count
                        outputFile.WriteLine(Left(writtenLine, Len(writtenLine) - 1))
                    End If
                    i += 1
                Loop Until i > ScriptDataRange.Rows.Count
                LogInfo("stored args to " + curWbPrefix + argdir + "\" + argFilename)
            Catch ex As Exception
                If outputFile IsNot Nothing Then outputFile.Close()
                If Not ScriptAddin.UserMsg("Error occured when creating inputfile '" + argFilename + "', " + ex.Message + " (maybe defined the wrong cell format for values?)",, True) Then Return False
            End Try
            If outputFile IsNot Nothing Then outputFile.Close()
        Next
        Return True
    End Function

    ''' <summary>creates script files for defined scriptRng ranges</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function storeScriptRng() As Boolean
        Dim scriptRngFilename As String = vbNullString, scriptText = vbNullString
        Dim ScriptDataRange As Range = Nothing
        Dim outputFile As StreamWriter = Nothing

        Dim scriptRngdir As String = dirglobal
        For c As Integer = 0 To ScriptDefDic("scriptrng").Length - 1
            Try
                Dim ErrMsg As String
                ' scriptrng beginning with a "=" is a scriptcell (as defined in getScriptDefinitions) ...
                If Left(ScriptDefDic("scriptrng")(c), 1) = "=" Then
                    scriptText = ScriptDefDic("scriptrng")(c).Substring(1)
                    scriptRngFilename = "ScriptDataRangeRow" + c.ToString() + ScriptFileSuffix
                Else
                    ErrMsg = prepareParam(c, "scriptrng", ScriptDataRange, scriptRngFilename, scriptRngdir, ScriptFileSuffix)
                    If Len(ErrMsg) > 0 Then
                        If Not ScriptAddin.UserMsg(ErrMsg) Then Return False
                    End If
                End If

                ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptRngdir
                Dim curWbPrefix As String = IIf(Left(scriptRngdir, 2) = "\\" Or Mid(scriptRngdir, 2, 2) = ":\", "", currWb.Path + "\")
                outputFile = New StreamWriter(curWbPrefix + scriptRngdir + "\" + scriptRngFilename, False, Encoding.Default)

                ' reuse the script invocation methods by setting the respective parameters
                ReDim Preserve ScriptDefDic("scripts")(ScriptDefDic("scripts").Length)
                ScriptDefDic("scripts")(ScriptDefDic("scripts").Length - 1) = scriptRngFilename
                ReDim Preserve ScriptDefDic("scriptspaths")(ScriptDefDic("scriptspaths").Length)
                ScriptDefDic("scriptspaths")(ScriptDefDic("scriptspaths").Length - 1) = scriptRngdir
                ReDim Preserve ScriptDefDic("skipscripts")(ScriptDefDic("skipscripts").Length)
                ScriptDefDic("skipscripts")(ScriptDefDic("skipscripts").Length - 1) = False

                ' write the ScriptDataRange or scriptText (if script directly in cell/formula right next to scriptrng) to file
                If Not IsNothing(scriptText) Then
                    outputFile.WriteLine(scriptText)
                Else
                    Dim i As Integer = 1
                    Do
                        Dim j As Integer = 1
                        Dim writtenLine As String = ""
                        If ScriptDataRange(i, 1).Value2 IsNot Nothing Then
                            Do
                                writtenLine += ScriptDataRange(i, j).Value2
                                j += 1
                            Loop Until j > ScriptDataRange.Columns.Count
                            outputFile.WriteLine(writtenLine)
                        End If
                        i += 1
                    Loop Until i > ScriptDataRange.Rows.Count
                End If
                LogInfo("stored Script to " + curWbPrefix + scriptRngdir + "\" + scriptRngFilename)
            Catch ex As Exception
                If outputFile IsNot Nothing Then outputFile.Close()
                If Not ScriptAddin.UserMsg("Error occured when creating script file '" + scriptRngFilename + "', " + ex.Message,, True) Then Return False
            End Try
            If outputFile IsNot Nothing Then outputFile.Close()
        Next
        Return True
    End Function

    Public fullScriptPath As String
    Public script As String
    Public scriptarguments As String
    Public previousDir As String
    Public theScriptOutput As ScriptOutput

    ''' <summary>invokes current scripts/args/results definition</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function invokeScripts() As Boolean
        Dim scriptpath As String
        previousDir = Directory.GetCurrentDirectory()
        scriptpath = dirglobal
        ' start script invocation loop as asynchronous thread to allow blocking ShowDialog while not blocking main Excel GUI thread (allows switching the dialog on/off)
        Task.Run(Async Function()
                     Dim ErrMsg As String = ""
                     For c As Integer = 0 To ScriptDefDic("scripts").Length - 1
                         ' skip script if defined...
                         If ScriptDefDic("skipscripts")(c) Then Continue For
                         ' in case you are wondering about scriptrng, this reuses the scripts dictionary by adding the saved file to the parameters...
                         ErrMsg = prepareParam(c, "scripts", Nothing, script, scriptpath, "")
                         If Len(ErrMsg) > 0 Then
                             ' allow to ignore preparation errors...
                             If Not ScriptAddin.UserMsg(ErrMsg) Then Exit For
                             ErrMsg = ""
                         End If

                         ' a blank separator indicates additional arguments, separate argument passing because of possible blanks in path -> need quotes around path + scriptname
                         ' assumption: scriptname itself may not have blanks in it.
                         If InStr(script, " ") > 0 Then
                             scriptarguments = script.Substring(InStr(script, " "))
                             script = script.Substring(0, InStr(script, " ") - 1)
                         End If

                         ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptpath
                         Dim curWbPrefix As String = IIf(Left(scriptpath, 2) = "\\" Or Mid(scriptpath, 2, 2) = ":\", "", currWb.Path + "\")
                         fullScriptPath = curWbPrefix + scriptpath
                         ' blocking wait for finish of script dialog
                         Await Task.Run(Sub()
                                            theScriptOutput = New ScriptOutput()
                                            If theScriptOutput.errMsg <> "" Then Exit Sub
                                            ' hide script output if not in debug mode
                                            If Not ScriptAddin.debugScript Then theScriptOutput.Opacity = 0
                                            theScriptOutput.ShowInTaskbar = True
                                            theScriptOutput.BringToFront()
                                            theScriptOutput.ShowDialog()
                                            ErrMsg = theScriptOutput.errMsg
                                        End Sub)
                     Next
                     ' after all scripts were finished and no ErrMsg from prepareParam or script, continue with result collection
                     If ErrMsg = "" Then
                         ScriptAddin.finishScriptprocess()
                     Else
                         ScriptAddin.UserMsg("Errors occurred in script, no returned results/diagrams will be fetched !", True, True)
                     End If
                     ' reset current dir
                     Directory.SetCurrentDirectory(previousDir)
                 End Function)
        Return True
    End Function

    ''' <summary>get Outputfiles for defined results ranges, tab separated
    ''' otherwise:  "what you see is what you get"
    ''' </summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function getResults() As Boolean
        Dim resFilename As String = vbNullString, readdir As String
        Dim ScriptDataRange As Range = Nothing
        Dim previousResultRange As Range
        Dim errMsg As String

        readdir = dirglobal
        For c As Integer = 0 To ScriptDefDic("results").Length - 1
            errMsg = prepareParam(c, "results", ScriptDataRange, resFilename, readdir, ".txt")
            If Len(errMsg) > 0 Then
                If Not ScriptAddin.UserMsg(errMsg,, True) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\readdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            If Not File.Exists(curWbPrefix + readdir + "\" + resFilename) Then
                If Not ScriptAddin.UserMsg("Results file '" + curWbPrefix + readdir + "\" + resFilename + "' not found!",, True) Then Return False
            End If
            ' remove previous content, might not exist, so catch any exception
            If ScriptDefDic("rresults")(c) Then
                Try
                    previousResultRange = currWb.Names.Item("___ScriptResult" + ScriptDefDic("results")(c)).RefersToRange
                    previousResultRange.ClearContents()
                Catch ex As Exception : End Try
                ' legacy R addin results
                Try
                    previousResultRange = currWb.Names.Item("___RaddinResult" + ScriptDefDic("results")(c)).RefersToRange
                    previousResultRange.ClearContents()
                Catch ex As Exception : End Try
            Else ' if we changed from rresults to results, need to remove hiddent ___ScriptResult name, otherwise results would still be removed when saving
                Try
                    currWb.Names.Item("___ScriptResult" + ScriptDefDic("results")(c)).Delete()
                Catch ex As Exception : End Try
                ' legacy R addin results
                Try
                    currWb.Names.Item("___RaddinResult" + ScriptDefDic("results")(c)).Delete()
                Catch ex As Exception : End Try
            End If

            Try
                Dim newQueryTable As QueryTable
                newQueryTable = ScriptDataRange.Worksheet.QueryTables.Add(Connection:="TEXT;" & curWbPrefix + readdir + "\" + resFilename, Destination:=ScriptDataRange)
                '                    .TextFilePlatform = 850
                With newQueryTable
                    .Name = "ScriptAddinResultData"
                    .FieldNames = True
                    .RowNumbers = False
                    .FillAdjacentFormulas = False
                    .PreserveFormatting = True
                    .RefreshOnFileOpen = False
                    .RefreshStyle = XlCellInsertionMode.xlOverwriteCells
                    .SavePassword = False
                    .SaveData = True
                    .AdjustColumnWidth = False
                    .RefreshPeriod = 0
                    .TextFileStartRow = 1
                    .TextFileParseType = XlTextParsingType.xlDelimited
                    .TextFileTabDelimiter = True
                    .TextFileSpaceDelimiter = False
                    .TextFileSemicolonDelimiter = False
                    .TextFileCommaDelimiter = False
                    .TextFileDecimalSeparator = "."
                    .TextFileThousandsSeparator = ","
                    .TextFileTrailingMinusNumbers = True
                    .Refresh(BackgroundQuery:=False)
                End With
                If ScriptDefDic("rresults")(c) Then
                    currWb.Names.Add(Name:="___ScriptResult" + ScriptDefDic("results")(c), RefersTo:=newQueryTable.ResultRange, Visible:=False)
                End If
                ' to avoid "hanging" names (Data) which add up quickly, try to remove the actually given name (could also be Data_1 if Data already exists) both from workbook and from current sheet
                Try : currWb.Names.Item(newQueryTable.Name).Delete() : Catch ex As Exception : End Try
                Try : ScriptDataRange.Worksheet.Names.Item(newQueryTable.Name).Delete() : Catch ex As Exception : End Try
                newQueryTable.Delete()
                LogInfo("inserted results from " + curWbPrefix + readdir + "\" + resFilename)
            Catch ex As Exception
                If Not ScriptAddin.UserMsg("Error in placing results in to Excel: " + ex.Message,, True) Then Return False
            End Try
        Next
        Return True
    End Function

    ''' <summary>get Output diagrams (png) for defined diags ranges</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function getDiags() As Boolean
        Dim diagFilename As String = vbNullString, readdir As String
        Dim ScriptDataRange As Range = Nothing
        Dim errMsg As String

        readdir = dirglobal
        For c As Integer = 0 To ScriptDefDic("diags").Length - 1
            errMsg = prepareParam(c, "diags", ScriptDataRange, diagFilename, readdir, ".png")
            If Len(errMsg) > 0 Then
                If Not ScriptAddin.UserMsg(errMsg,, True) Then Return False
            End If
            ' clean previously set shape...
            For Each oldShape As Shape In ScriptDataRange.Worksheet.Shapes
                If oldShape.Name = diagFilename Then
                    oldShape.Delete()
                    Exit For
                End If
            Next
            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\readdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            If Not File.Exists(curWbPrefix + readdir + "\" + diagFilename) Then
                If Not ScriptAddin.UserMsg("Diagram file '" + curWbPrefix + readdir + "\" + diagFilename + "' not found!",, True) Then Return False
            End If

            ' add new shape from picture
            Try
                With ScriptDataRange.Worksheet.Shapes.AddPicture(Filename:=curWbPrefix + readdir + "\" + diagFilename,
                    LinkToFile:=Microsoft.Office.Core.MsoTriState.msoFalse, SaveWithDocument:=Microsoft.Office.Core.MsoTriState.msoTrue, Left:=ScriptDataRange.Left, Top:=ScriptDataRange.Top, Width:=-1, Height:=-1)
                    .Name = diagFilename
                End With
                LogInfo("added shape for diagram " + curWbPrefix + readdir + "\" + diagFilename)
            Catch ex As Exception
                If Not ScriptAddin.UserMsg("Error occured when placing the diagram into target range '" + ScriptDefDic("diags")(c) + "', " + ex.Message,, True) Then Return False
            End Try
        Next
        Return True
    End Function

    ''' <summary>remove result, diagram and temporary script files</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function removeFiles() As Boolean
        Dim filename As String = vbNullString
        Dim readdir As String = dirglobal
        Dim ScriptDataRange As Range = Nothing
        Dim errMsg As String

        ' check for script existence before creating any potential missing folders below...
        For c As Integer = 0 To ScriptDefDic("scripts").Length - 1
            ' skip script if defined...
            If ScriptDefDic("skipscripts")(c) Then Continue For
            Dim script As String = vbNullString
            ' returns script and readdir !
            errMsg = prepareParam(c, "scripts", Nothing, script, readdir, "")
            If Len(errMsg) > 0 Then
                If Not ScriptAddin.UserMsg(errMsg,, True) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\scriptpath
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            Dim fullScriptPath = curWbPrefix + readdir

            ' a blank separator indicates additional arguments, separate argument passing because of possible blanks in path -> need quotes around path + scriptname
            ' assumption: scriptname itself may not have blanks in it.
            If InStr(script, " ") > 0 Then script = script.Substring(0, InStr(script, " "))
            If Not File.Exists(fullScriptPath + "\" + script) Then
                ScriptAddin.UserMsg("Script '" + fullScriptPath + "\" + script + "' not found!" + vbCrLf, True, True)
                Return False
            End If
            ' check if executable exists or exists somewhere in the path....
            Dim foundExe As Boolean = False
            Dim exe As String = Environment.ExpandEnvironmentVariables(ScriptExec)
            If Not File.Exists(exe) Then
                If Path.GetDirectoryName(exe) = String.Empty Then
                    For Each test In (Environment.GetEnvironmentVariable("PATH")).Split(";")
                        Dim thePath As String = test.Trim()
                        If Len(thePath) > 0 And File.Exists(Path.Combine(thePath, exe)) Then
                            foundExe = True
                            Exit For
                        End If
                    Next
                    If Not foundExe Then
                        ScriptAddin.UserMsg("Executable '" + ScriptExec + "' not found!" + vbCrLf, True, True)
                        Return False
                    Else
                        LogInfo("found exec " + ScriptExec)
                    End If
                End If
            End If
        Next

        ' remove input argument files
        For c As Integer = 0 To ScriptDefDic("args").Length - 1
            ' returns filename and readdir !
            errMsg = prepareParam(c, "args", ScriptDataRange, filename, readdir, ".txt")
            If Len(errMsg) > 0 Then
                If Not ScriptAddin.UserMsg(errMsg,, True) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' special comfort: if containing folder doesn't exist, create it now:
            If Not Directory.Exists(curWbPrefix + readdir) Then
                Try
                    Directory.CreateDirectory(curWbPrefix + readdir)
                Catch ex As Exception
                    If Not ScriptAddin.UserMsg("Error occured when trying to create input arguments containing folder '" + curWbPrefix + readdir + "', " + ex.Message,, True) Then Return False
                End Try
            End If
            ' remove any existing input files...
            If File.Exists(curWbPrefix + readdir + "\" + filename) Then
                File.Delete(curWbPrefix + readdir + "\" + filename)
                LogInfo("deleted input " + curWbPrefix + readdir + "\" + filename)
            End If
        Next

        ' remove result files
        For c As Integer = 0 To ScriptDefDic("results").Length - 1
            ' returns filename and readdir !
            errMsg = prepareParam(c, "results", ScriptDataRange, filename, readdir, ".txt")
            If Len(errMsg) > 0 Then
                If Not ScriptAddin.UserMsg(errMsg,, True) Then Return False
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' special comfort: if containing folder doesn't exist, create it now:
            If Not Directory.Exists(curWbPrefix + readdir) Then
                Try
                    Directory.CreateDirectory(curWbPrefix + readdir)
                Catch ex As Exception
                    If Not ScriptAddin.UserMsg("Error occured when trying to create result containing folder '" + curWbPrefix + readdir + "', " + ex.Message,, True) Then Return False
                End Try
            End If
            ' remove any existing result files...
            If File.Exists(curWbPrefix + readdir + "\" + filename) Then
                Try
                    File.Delete(curWbPrefix + readdir + "\" + filename)
                    LogInfo("deleted result " + curWbPrefix + readdir + "\" + filename)
                Catch ex As Exception
                    If Not ScriptAddin.UserMsg("Error occured when trying to remove '" + curWbPrefix + readdir + "\" + filename + "', " + ex.Message,, True) Then Return False
                End Try
            End If
        Next

        ' remove diagram files
        For c As Integer = 0 To ScriptDefDic("diags").Length - 1
            ' returns filename and readdir !
            errMsg = prepareParam(c, "diags", ScriptDataRange, filename, readdir, ".png")
            If Len(errMsg) > 0 Then
                If Not ScriptAddin.UserMsg(errMsg,, True) Then Return False
            End If
            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' special comfort: if containing folder doesn't exist, create it now:
            If Not Directory.Exists(curWbPrefix + readdir) Then
                Try
                    Directory.CreateDirectory(curWbPrefix + readdir)
                Catch ex As Exception
                    If Not ScriptAddin.UserMsg("Error occured when trying to create diagram container folder '" + curWbPrefix + readdir + "', " + ex.Message,, True) Then Return False
                End Try
            End If
            ' remove any existing diagram files...
            If File.Exists(curWbPrefix + readdir + "\" + filename) Then
                Try
                    File.Delete(curWbPrefix + readdir + "\" + filename)
                    LogInfo("deleted diagram " + curWbPrefix + readdir + "\" + filename)
                Catch ex As Exception
                    If Not ScriptAddin.UserMsg("Error occured when trying to remove '" + curWbPrefix + readdir + "\" + filename + "', " + ex.Message,, True) Then Return False
                End Try
            End If
        Next

        ' remove temporary script files
        For c As Integer = 0 To ScriptDefDic("scriptrng").Length - 1
            ' returns filename and readdir !
            errMsg = prepareParam(c, "scriptrng", ScriptDataRange, filename, readdir, ScriptFileSuffix)
            If Len(errMsg) > 0 Then
                filename = "ScriptDataRangeRow" + c.ToString() + ScriptFileSuffix
            End If

            ' absolute paths begin with \\ or X:\ -> dont prefix with currWB path, else currWBpath\argdir
            Dim curWbPrefix As String = IIf(Left(readdir, 2) = "\\" Or Mid(readdir, 2, 2) = ":\", "", currWb.Path + "\")
            ' special comfort: if containing folder doesn't exist, create it now:
            If Not Directory.Exists(curWbPrefix + readdir) Then
                Try
                    Directory.CreateDirectory(curWbPrefix + readdir)
                Catch ex As Exception
                    If Not ScriptAddin.UserMsg("Error occured when trying to create script containing folder '" + curWbPrefix + readdir + "', " + ex.Message,, True) Then Return False
                End Try
            End If
            ' remove any existing diagram files...
            If File.Exists(curWbPrefix + readdir + "\" + filename) Then
                Try
                    File.Delete(curWbPrefix + readdir + "\" + filename)
                    LogInfo("deleted temporary rscript " + curWbPrefix + readdir + "\" + filename)
                Catch ex As Exception
                    If Not ScriptAddin.UserMsg("Error occured when trying to remove '" + curWbPrefix + readdir + "\" + filename + "', " + ex.Message,, True) Then Return False
                End Try
            End If
        Next
        Return True
    End Function

    ''' <summary>encapsulates setting fetching (currently ConfigurationManager from DBAddin.xll.config)</summary>
    ''' <param name="Key">registry key to take value from</param>
    ''' <param name="defaultValue">Value that is taken if Key was not found</param>
    ''' <returns>the setting value</returns>
    Public Function fetchSetting(Key As String, defaultValue As String) As String
        Dim UserSettings As NameValueCollection = Nothing
        Dim AddinAppSettings As NameValueCollection = Nothing
        Try : UserSettings = ConfigurationManager.GetSection("UserSettings") : Catch ex As Exception : LogWarn("Error reading UserSettings: " + ex.Message) : End Try
        Try : AddinAppSettings = ConfigurationManager.AppSettings : Catch ex As Exception : LogWarn("Error reading AppSettings: " + ex.Message) : End Try
        ' user specific settings are in UserSettings section in separate file
        If IsNothing(UserSettings) OrElse IsNothing(UserSettings(Key)) Then
            If Not IsNothing(AddinAppSettings) Then
                fetchSetting = AddinAppSettings(Key)
            Else
                fetchSetting = Nothing
            End If
        ElseIf Not (IsNothing(UserSettings) OrElse IsNothing(UserSettings(Key))) Then
            fetchSetting = UserSettings(Key)
        Else
            fetchSetting = Nothing
        End If
        If fetchSetting Is Nothing Then fetchSetting = defaultValue
    End Function

    ''' <summary>change or add a key/value pair in the user settings</summary>
    ''' <param name="theKey">key to change (or add)</param>
    ''' <param name="theValue">value for key</param>
    Public Sub setUserSetting(theKey As String, theValue As String)
        ' check if key exists
        Dim doc As New Xml.XmlDocument()
        doc.Load(UserSettingsPath)
        Dim keyNode As Xml.XmlNode = doc.SelectSingleNode("/UserSettings/add[@key='" + System.Security.SecurityElement.Escape(theKey) + "']")
        If IsNothing(keyNode) Then
            ' if not, add to settings
            Dim nodeRegion As Xml.XmlElement = doc.CreateElement("add")
            nodeRegion.SetAttribute("key", theKey)
            nodeRegion.SetAttribute("value", theValue)
            doc.SelectSingleNode("//UserSettings").AppendChild(nodeRegion)
        Else
            keyNode.Attributes().GetNamedItem("value").InnerText = theValue
        End If
        doc.Save(UserSettingsPath)
        ConfigurationManager.RefreshSection("UserSettings")
    End Sub

    ''' <summary>Msgbox that avoids further Msgboxes (click Yes) or cancels run altogether (click Cancel)</summary>
    ''' <param name="message"></param>
    ''' <returns>True if further Msgboxes should be avoided, False otherwise</returns>
    Public Function UserMsg(message As String, Optional noAvoidChoice As Boolean = False, Optional IsWarning As Boolean = False) As Boolean
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName & "." & theMethod.Name
        WriteToLog(message, If(IsWarning, EventLogEntryType.Warning, EventLogEntryType.Information), caller)
        If nonInteractive Then Return False
        theRibbon.InvalidateControl("showLog")
        If noAvoidChoice Then
            MsgBox(message, MsgBoxStyle.OkOnly + IIf(IsWarning, MsgBoxStyle.Critical, MsgBoxStyle.Information), "ScriptAddin Message")
            Return False
        Else
            If avoidFurtherMsgBoxes Then Return True
            Dim retval As MsgBoxResult = MsgBox(message + vbCrLf + "Avoid further Messages (Yes/No) or abort ScriptDefinition (Cancel)", MsgBoxStyle.YesNoCancel, "ScriptAddin Message")
            If retval = MsgBoxResult.Yes Then avoidFurtherMsgBoxes = True
            Return (retval = MsgBoxResult.Yes Or retval = MsgBoxResult.No)
        End If
    End Function

    ''' <summary>ask User (default OKCancel) and log as warning if Critical Or Exclamation (logged errors would pop up the trace information window)</summary> 
    ''' <param name="theMessage">the question to be shown/logged</param>
    ''' <param name="questionType">optionally pass question box type, default MsgBoxStyle.OKCancel</param>
    ''' <param name="questionTitle">optionally pass a title for the msgbox instead of default DBAddin Question</param>
    ''' <param name="msgboxIcon">optionally pass a different Msgbox icon (style) instead of default MsgBoxStyle.Question</param>
    ''' <returns>choice as MsgBoxResult (Yes, No, OK, Cancel...)</returns>
    Public Function QuestionMsg(theMessage As String, Optional questionType As MsgBoxStyle = MsgBoxStyle.OkCancel, Optional questionTitle As String = "ScriptAddin Question", Optional msgboxIcon As MsgBoxStyle = MsgBoxStyle.Question) As MsgBoxResult
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
        WriteToLog(theMessage, If(msgboxIcon = MsgBoxStyle.Critical Or msgboxIcon = MsgBoxStyle.Exclamation, EventLogEntryType.Warning, EventLogEntryType.Information), caller) ' to avoid popup of trace log
        If nonInteractive Then
            If questionType = MsgBoxStyle.OkCancel Then Return MsgBoxResult.Cancel
            If questionType = MsgBoxStyle.YesNo Then Return MsgBoxResult.No
            If questionType = MsgBoxStyle.YesNoCancel Then Return MsgBoxResult.No
            If questionType = MsgBoxStyle.RetryCancel Then Return MsgBoxResult.Cancel
        End If
        ' tab is not activated BEFORE Msgbox as Excel first has to get into the interaction thread outside this one..
        If theRibbon IsNot Nothing Then theRibbon.ActivateTab("ScriptaddinTab")
        Return MsgBox(theMessage, msgboxIcon + questionType, questionTitle)
    End Function

    ''' <summary>Logs Message of eEventType to System.Diagnostics.Trace</summary>
    ''' <param name="Message">Message to be logged</param>
    ''' <param name="eEventType">event type: info, warning, error</param>
    ''' <param name="caller">reflection based caller information: module.method</param>
    Private Sub WriteToLog(Message As String, eEventType As EventLogEntryType, caller As String)
        If nonInteractive Then
            ' collect errors and warnings for returning messages in executeScript
            If eEventType = EventLogEntryType.Error Or eEventType = EventLogEntryType.Warning Then nonInteractiveErrMsgs += caller + ":" + Message + vbCrLf
            Trace.TraceInformation("Noninteractive: {0}: {1}", caller, Message)
        Else
            Select Case eEventType
                Case EventLogEntryType.Information
                    Trace.TraceInformation("{0}: {1}", caller, Message)
                Case EventLogEntryType.Warning
                    Trace.TraceWarning("{0}: {1}", caller, Message)
                    WarningIssued = True
                    ' at Addin Start ribbon has not been loaded so avoid call to it here..
                    If theRibbon IsNot Nothing Then theRibbon.InvalidateControl("showLog")
                Case EventLogEntryType.Error
                    Trace.TraceError("{0}: {1}", caller, Message)
            End Select
        End If
    End Sub

    ''' <summary>Logs error messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogError(LogMessage As String)
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
        WriteToLog(LogMessage, EventLogEntryType.Error, caller)
    End Sub

    ''' <summary>Logs warning messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogWarn(LogMessage As String)
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
        WriteToLog(LogMessage, EventLogEntryType.Warning, caller)
    End Sub

    ''' <summary>Logs informational messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogInfo(LogMessage As String)
        If DebugAddin Then
            Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
            Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
            WriteToLog(LogMessage, EventLogEntryType.Information, caller)
        End If
    End Sub

    Public ScriptExecutables As List(Of String)

    ''' <summary>initialise the ScriptExecutables list</summary>
    Public Sub initScriptExecutables()
        Dim AddinAppSettings As NameValueCollection = Nothing
        Try : AddinAppSettings = ConfigurationManager.AppSettings : Catch ex As Exception : LogWarn("Error reading AppSettings for ScriptExecutables (ExePath) entries: " + ex.Message) : End Try
        ScriptExecutables = New List(Of String)
        ' getting Usersettings might fail (formatting, etc)...
        If Not IsNothing(AddinAppSettings) Then
            For Each key As String In AddinAppSettings.AllKeys
                If Left(key, 7) = "ExePath" Then ScriptExecutables.Add(key.Substring(7))
            Next
        End If
    End Sub

    Public Sub insertScriptExample()
        If QuestionMsg("Inserting Example Script definition starting in current cell, overwriting 8 rows and 3 columns with example definitions!") = MsgBoxResult.Cancel Then Exit Sub
        Dim curCell As Range = ExcelDnaUtil.Application.ActiveCell
        curCell.Value = "Dir"
        curCell.Offset(0, 1).Value = "."
        curCell.Offset(1, 0).Value = "Type"
        curCell.Offset(1, 1).Value = "R"
        curCell.Offset(2, 0).Value = "script"
        curCell.Offset(2, 1).Value = "yourScript.R"
        curCell.Offset(2, 2).Value = "."
        curCell.Offset(3, 0).Value = "scriptCell"
        curCell.Offset(3, 1).Value = "# your script code in this cell"
        curCell.Offset(3, 2).Value = "."
        curCell.Offset(4, 0).Value = "scriptRange"
        curCell.Offset(4, 1).Value = "yourScriptCodeInThisRange"
        curCell.Offset(4, 2).Value = "."
        curCell.Offset(5, 0).Value = "arg"
        curCell.Offset(5, 1).Value = "yourArgInputRange"
        curCell.Offset(5, 2).Value = "."
        curCell.Offset(6, 0).Value = "res"
        curCell.Offset(6, 1).Value = "yourResultOutRange"
        curCell.Offset(6, 2).Value = "."
        curCell.Offset(7, 0).Value = "diag"
        curCell.Offset(7, 1).Value = "yourDiagramPlaceRange"
        curCell.Offset(7, 2).Value = "."
        Try
            ExcelDnaUtil.Application.ActiveSheet.Range(curCell, curCell.Offset(7, 2)).Name = "Script_Example"
        Catch ex As Exception
            UserMsg("Couldn't name example definitions as 'Script_Example': " + ex.Message)
        End Try
    End Sub

End Module
