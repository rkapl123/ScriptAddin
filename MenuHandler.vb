Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI
Imports ExcelDna.Logging
Imports System.Configuration

''' <summary>Events from Ribbon</summary>
<ComVisible(True)>
Public Class MenuHandler
    Inherits ExcelRibbon
    ''' <summary>the selected index of the script executable (R, Python,...)</summary>
    Public selectedScriptExecutable As Integer

    ''' <summary></summary>
    Public Sub ribbonLoaded(myribbon As IRibbonUI)
        ScriptAddin.theRibbon = myribbon
        ScriptAddin.debugScript = CBool(ScriptAddin.fetchSetting("debugScript", "False"))
        selectedScriptExecutable = CInt(ScriptAddin.fetchSetting("selectedScriptExecutable", "0"))
        ScriptAddin.WarningIssued = False
        If ScriptAddin.ScriptExecutables.Count > 0 Then ScriptAddin.ScriptType = ScriptAddin.ScriptExecutables(selectedScriptExecutable)
    End Sub

    ''' <summary>creates the Ribbon</summary>
    ''' <returns></returns>
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='ribbonLoaded'>" +
        "<ribbon><tabs><tab id='ScriptaddinTab' label='ScriptAddin'>" +
            "<group id='ScriptaddinGroup' label='General settings'>" +
              "<dropDown id='scriptDropDown' label='ScriptDefinition:' sizeString='12345678901234567890' getItemCount='GetItemCount' getItemID='GetItemID' getItemLabel='GetItemLabel' getSelectedItemIndex='GetSelectedScript' onAction='selectItem'/>" +
              "<dropDown id='execDropDown' label='ScriptExecutable:' sizeString='12345678901234' getItemCount='GetItemCountExec' getItemID='GetItemIDExec' getItemLabel='GetItemLabelExec' getSelectedItemIndex='GetSelectedExec' onAction='selectItemExec'/>" +
              "<buttonGroup id='butGrp'>" +
                "<menu id='configMenu' label='Settings'>" +
                  "<button id='insExample' label='insert Example' tag='5' screentip='insert an Example Script Range' imageMso='SignatureLineInsert' onAction='insertExample'/>" +
                  "<button id='user' label='User settings' onAction='showAddinConfig' imageMso='ControlProperties' screentip='Show/edit user settings for Script Addin' />" +
                  "<button id='central' label='Central settings' onAction='showAddinConfig' imageMso='TablePropertiesDialog' screentip='Show/edit central settings for Script Addin' />" +
                  "<button id='addin' label='ScriptAddin settings' onAction='showAddinConfig' imageMso='ServerProperties' screentip='Show/edit standard Addin settings for Script Addin' />" +
                "</menu>" +
                "<toggleButton id='debug' label='script output win' onAction='toggleButton' getImage='getImage' getPressed='getPressed' tag='3' screentip='toggles script output window visibility' supertip='for debugging you can display the script output' />" +
                "<button id='showLog' label='Log' tag='4' screentip='shows Scriptaddins Diagnostic Display' getImage='getLogsImage' onAction='clickShowLog'/>" +
              "</buttonGroup>" +
            "<dialogBoxLauncher><button id='dialog' label='About Scriptaddin' onAction='refreshScriptDefs' tag='5' screentip='Show Aboutbox (and refresh ScriptDefinitions from current Workbook from there)'/></dialogBoxLauncher></group>" +
            "<group id='ScriptsGroup' label='Run Scripts defined in WB/sheet names'>"
        Dim presetSheetButtonsCount As Integer = Int16.Parse(ScriptAddin.fetchSetting("presetSheetButtonsCount", "15"))
        Dim thesize As String = IIf(presetSheetButtonsCount < 15, "normal", "large")
        For i As Integer = 0 To presetSheetButtonsCount
            customUIXml = customUIXml + "<dynamicMenu id='ID" + i.ToString() + "' " +
                                            "size='" + thesize + "' getLabel='getSheetLabel' imageMso='SignatureLineInsert' " +
                                            "screentip='Select script to run' " +
                                            "getContent='getDynMenContent' getVisible='getDynMenVisible'/>"
        Next
        customUIXml += "</group></tab></tabs></ribbon></customUI>"
        Return customUIXml
    End Function

    ''' <summary>show xll standard config (AppSetting), central config (referenced by App Settings file attr) or user config (referenced by CustomSettings configSource attr)</summary>
    ''' <param name="control"></param>
    Public Sub showAddinConfig(control As IRibbonControl)
        ' if settings (addin, user, central) should not be displayed according to setting then exit...
        If InStr(ScriptAddin.fetchSetting("disableSettingsDisplay", ""), control.Id) > 0 Then
            ScriptAddin.UserMsg("Display of " + control.Id + " settings disabled !", True, True)
            Exit Sub
        End If
        Dim theSettingsDlg As EditSettings = New EditSettings()
        theSettingsDlg.Tag = control.Id
        theSettingsDlg.ShowDialog()
        If control.Id = "addin" Or control.Id = "central" Then
            ConfigurationManager.RefreshSection("appSettings")
        Else
            ConfigurationManager.RefreshSection("UserSettings")
        End If
        ' reflect changes in settings
        initScriptExecutables()
        ' also display in ribbon
        theRibbon.Invalidate()
    End Sub

    ''' <summary>after clicking on the script drop down button, the defined script definition is started</summary>
    Public Sub startScript(control As IRibbonControl)
        Dim errStr As String
        ' set ScriptDefinition to invocaters range... invocating sheet is put into Tag
        ScriptAddin.ScriptDefinitionRange = ScriptAddin.ScriptDefsheetColl(control.Tag).Item(control.Id)
        ScriptAddin.SkipScriptAndPreparation = My.Computer.Keyboard.CtrlKeyDown
        Try
            ScriptAddin.ScriptDefinitionRange.Parent.Select()
        Catch ex As Exception
            ScriptAddin.UserMsg("Selection of worksheet of Script Definition Range not possible (probably because you're editing a cell)!", True, True)
        End Try
        ScriptAddin.ScriptDefinitionRange.Select()
        errStr = ScriptAddin.startScriptprocess()
        If errStr <> "" Then ScriptAddin.UserMsg(errStr, True, True)
    End Sub

    ''' <summary>reflect the change in the toggle buttons title</summary>
    ''' <returns></returns>
    Public Function getImage(control As IRibbonControl) As String
        If ScriptAddin.debugScript And control.Id = "debug" Then
            Return "AcceptTask"
        Else
            Return "DeclineTask"
        End If
    End Function

    ''' <summary>reflect the change in the toggle buttons title</summary>
    ''' <returns>True for the respective control if activated</returns>
    Public Function getPressed(control As IRibbonControl) As Boolean
        If control.Id = "debug" Then
            Return ScriptAddin.debugScript
        Else
            Return False
        End If
    End Function

    ''' <summary>toggle debug button</summary>
    ''' <param name="pressed"></param>
    Public Sub toggleButton(control As IRibbonControl, pressed As Boolean)
        If control.Id = "debug" Then
            ScriptAddin.debugScript = pressed
            ScriptAddin.setUserSetting("debugScript", pressed.ToString())
            If Not IsNothing(ScriptAddin.theScriptOutput) Then
                If pressed Then
                    ScriptAddin.theScriptOutput.Opacity = 1.0
                    ScriptAddin.theScriptOutput.BringToFront()
                    ScriptAddin.theScriptOutput.Refresh()
                Else
                    ScriptAddin.theScriptOutput.Opacity = 0.0
                End If
            End If
            ' invalidate to reflect the change in the toggle buttons image
            ScriptAddin.theRibbon.InvalidateControl(control.Id)
        End If
    End Sub

    ''' <summary></summary>
    Public Sub refreshScriptDefs(control As IRibbonControl)
        Dim myAbout As AboutBox1 = New AboutBox1
        myAbout.ShowDialog()
    End Sub

    ''' <summary></summary>
    ''' <returns></returns>
    Public Function GetItemCount(control As IRibbonControl) As Integer
        Return (ScriptAddin.Scriptcalldefnames.Length)
    End Function

    ''' <summary></summary>
    ''' <returns></returns>
    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String
        Return ScriptAddin.Scriptcalldefnames(index)
    End Function

    ''' <summary></summary>
    ''' <returns></returns>
    Public Function GetItemID(control As IRibbonControl, index As Integer) As String
        Return ScriptAddin.Scriptcalldefnames(index)
    End Function

    Private selectedScript As Integer

    ''' <summary>after selection of script used to return the selected script</summary>
    ''' <returns></returns>
    Public Function GetSelectedScript(control As IRibbonControl) As Integer
        Return selectedScript
    End Function

    ''' <summary></summary>
    Public Sub selectItem(control As IRibbonControl, id As String, index As Integer)
        ' needed for workbook save (saves selected ScriptDefinition)
        selectedScript = index
        ScriptAddin.dropDownSelected = True
        ScriptAddin.ScriptDefinitionRange = Scriptcalldefs(index)
        ScriptAddin.ScriptDefinitionRange.Parent.Select()
        ScriptAddin.ScriptDefinitionRange.Select()
    End Sub

    ''' <summary></summary>
    ''' <returns></returns>
    Public Function GetItemCountExec(control As IRibbonControl) As Integer
        Return ScriptExecutables.Count
    End Function

    ''' <summary></summary>
    ''' <returns></returns>
    Public Function GetItemLabelExec(control As IRibbonControl, index As Integer) As String
        If ScriptExecutables.Count > 0 Then
            Return ScriptExecutables(index)
        Else
            Return ""
        End If
    End Function

    ''' <summary></summary>
    ''' <returns></returns>
    Public Function GetItemIDExec(control As IRibbonControl, index As Integer) As String
        If ScriptExecutables.Count > 0 Then
            Return ScriptExecutables(index)
        Else
            Return ""
        End If
    End Function

    ''' <summary>after selection of executable used to return the selected executable for display</summary>
    ''' <returns></returns>
    Public Function GetSelectedExec(control As IRibbonControl) As Integer
        Return selectedScriptExecutable
    End Function

    ''' <summary>select a script executable from the ScriptExecutable dropdown</summary>
    Public Sub selectItemExec(control As IRibbonControl, id As String, index As Integer)
        selectedScriptExecutable = index
        ScriptAddin.ScriptType = ScriptAddin.ScriptExecutables(selectedScriptExecutable)
        ScriptAddin.setUserSetting("selectedScriptExecutable", index.ToString())
    End Sub

    ''' <summary>display warning icon on log button if warning has been logged...</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function getLogsImage(control As IRibbonControl) As String
        If ScriptAddin.WarningIssued Then
            Return "IndexUpdate"
        Else
            Return "MailMergeStartLetters"
        End If
    End Function

    ''' <summary>insert an Script_Example</summary>
    ''' <param name="control"></param>
    Public Sub insertExample(control As IRibbonControl)
        ScriptAddin.insertScriptExample()
    End Sub


    ''' <summary>show the trace log</summary>
    ''' <param name="control"></param>
    Public Sub clickShowLog(control As IRibbonControl)
        LogDisplay.Show()
        ' reset warning flag
        ScriptAddin.WarningIssued = False
        theRibbon.InvalidateControl("showLog")
    End Sub

    ''' <summary>set the name of the WB/sheet dropdown to the sheet name (for the WB dropdown this is the WB name)</summary>
    ''' <returns></returns>
    Public Function getSheetLabel(control As IRibbonControl) As String
        getSheetLabel = vbNullString
        If ScriptAddin.ScriptDefsheetMap.ContainsKey(control.Id) Then getSheetLabel = ScriptAddin.ScriptDefsheetMap(control.Id)
    End Function

    ''' <summary>create the buttons in the WB/sheet dropdown</summary>
    ''' <returns></returns>
    Public Function getDynMenContent(control As IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"
        Dim currentSheet As String = ScriptAddin.ScriptDefsheetMap(control.Id)
        For Each nodeName As String In ScriptAddin.ScriptDefsheetColl(currentSheet).Keys
            xmlString = xmlString + "<button id='" + nodeName + "' label='run " + nodeName + "' imageMso='SignatureLineInsert' onAction='startScript' tag ='" + currentSheet + "' screentip='run " + nodeName + " ScriptDefinition' supertip='runs script defined in " + nodeName + " ScriptAddin range on sheet " + currentSheet + "' />"
        Next
        xmlString += "</menu>"
        Return xmlString
    End Function

    ''' <summary>shows the sheet button only if it was collected...</summary>
    ''' <returns>visible or not</returns>
    Public Function getDynMenVisible(control As IRibbonControl) As Boolean
        Return ScriptAddin.ScriptDefsheetMap.ContainsKey(control.Id)
    End Function

End Class