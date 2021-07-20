Script Addin provides an easy way to define and run script interactions started from Excel.

# Using ScriptAddin

Running an script is simply done by selecting the desired executable (R, Python, Perl, whatever was defined) on the Script Addin Ribbon Tab and clicking "run <ScriptDefinition>"
beneath the Sheet-button in the Ribbon group "Run Scripts defined in WB/sheets names". Activating the "debug script" toggle button opens the script output.
Selecting the Script definition in the ScriptDefinition dropdown highlights the specified definition range.

When running scripts, following is executed:

The input arguments (arg, see below) are written to files, the scripts defined inside Excel are written and called using the executable located in ExePath/exec (see settings), the defined results/diagrams that were written to file are read and placed in Excel.


![Image of screenshot1](https://raw.githubusercontent.com/rkapl123/ScriptAddin/master/docs/screenshot1.png)

# Defining ScriptAddin script interactions (ScriptDefinitions)

script interactions (ScriptDefinitions) are defined using a 3 column named range (1st col: definition type, 2nd: definition value, 3rd: (optional) definition path):

The Rdefinition range name must start with "Script_" (or "R_Addin" as a legacy compatibility with the old R Addin) and can have a postfix as an additional definition name.
If there is no postfix after "Script_", the script is called "MainScript" in the Workbook/Worksheet.

A range name can be at Workbook level or worksheet level.
In the ScriptDefinition dropdowns the worksheet name (for worksheet level names) or the workbook name (for workbook level names) is prepended to the additional postfixed definition name.

So for the 10 definitions (range names) currently defined in the test workbook testRAddin.xlsx, there should be 10 entries in the Rdefinition dropdown:

- testScriptAddin.xlsx, (Workbooklevel name, runs as MainScript)
- testScriptAddin.xlsxAnotherDef (Workbooklevel name),
- testScriptAddin.xlsxErrorInDef (Workbooklevel name),
- testScriptAddin.xlsxNewResDiagDir (Workbooklevel name),
- Test_OtherSheet, (name in Test_OtherSheet)
- Test_OtherSheetAnotherDef (name in Test_OtherSheet),
- Test_scriptRngScriptCell (Test_scriptRng) and
- Test_scriptRngScriptRange (Test_scriptRng)

In the 1st column of the Rdefinition range are the definition types, possible types are
- exec: an executable, being able to run the script in line "script". This is only needed for overriding the ExePath<executable> in the AppSettings in the ScriptAddin.xll.config file.
- path: path to folders with dlls/executables (semicolon separated), in case you need to add them. Only needed when overriding the PathAdd<executable> in the AppSettings in the ScriptAddin.xll.config file.
- dir: the path where below files (scripts, args, results and diagrams) are stored.
- script: full path of an executable script.
- arg/arg[rc] (input objects, txt files): variable name and path/filename, where the (input) arguments are stored.
- res/rres (output objects, txt files): variable name and path/filename, where the (output) results are expected. If the definition type is rres, results are removed from excel before saving and rerunning the script
- diag (output diagrams, png format): path/filename, where (output) diagrams are expected.
- scriptrng/scriptcell (Scripts directly within Excel): either ranges, where a script is stored (scriptrng) or directly as a cell value (text content or formula result) in the 2nd column (scriptcell)

Scripts (defined with the script, scriptrng or scriptcell definition types) are executed in sequence of their appearance. Although exec, path and dir definitions can appear more than once, only the last definition is used.

In the 2nd column are the definition values as described above.
- For arg, res, scriptrng and diag these are range names referring to the respective ranges to be taken as arg, res, scriptrng or diag target in the excel workbook.
- The range names that are referred in arg, res, scriptrng and diag types can also be either workbook global (having no ! in front of the name) or worksheet local (having the worksheet name + ! in front of the name)
- for exec this can either be the full path for the executable, or - in case the executable is already in the windows default path - a simple filename (like cmd.exe or perl.exe)
- for path, an additional path to folders with dlls/executables (semicolon separated), in case you need to add them

In the 3rd column are the definition paths of the files referred to in arg, res and diag
- Absolute Paths in dir or the definition path column are defined by starting with \\ or X:\ (X being a mapped drive letter)
- not existing paths for arg, res, scriptrng/scriptcell and diag are created automatically, so dynamical paths can be given here.
- for exec, additional commandline switches can be passed here to the executable (like "/c" for cmd.exe, this is required to start a subsequent script)
- for path, the default file suffix for the script files can be given/overruled here (".R", ".pl", ".py"...)

The definitions are loaded into the ScriptDefinition dropdown either on opening/activating a Workbook with above named areas or by pressing the small dialogBoxLauncher "Show AboutBox" on the Script Addin Ribbon Tab and clicking "refresh ScriptDefinitions":  
![Image of screenshot2](https://raw.githubusercontent.com/rkapl123/ScriptAddin/master/docs/screenshot2.png)

The mentioned hyperlink to the local help file can be configured in the app config file (ScriptAddin.xll.config) with key "LocalHelp".
When saving the Workbook the input arguments (definition with arg) defined in the currently selected Scriptdefinition dropdown are stored as well. If nothing is selected, the first Scriptdefinition of the dropdown is chosen.

The error messages are logged to a diagnostic log provided by ExcelDna, which can be accessed by clicking on "show Log". The log level can be set in the `system.diagnostics` section of the app-config file (Scriptaddin.xll.config):
Either you set the switchValue attribute of the source element to prevent any trace messages being generated at all, or you set the initializeData attribute of the added LogDisplay listener to prevent the generated messages to be shown (below the chosen level)  

Issues/Enhancements:

- [ ] Implement a faster way to save textfiles from excel

# Installation of ScriptAddin

run Distribution/deployAddin.cmd (this puts ScriptAddin32.xll/ScriptAddin64.xll as ScriptAddin.xll and ScriptAddin.xll.config into %appdata%\Microsoft\AddIns and starts installAddinInExcel.vbs (setting AddIns("ScriptAddin.xll").Installed = True in Excel)).

Adapt the settings in ScriptAddin.xll.config:

```XML
  <system.diagnostics>
    <sources>
      <source name="ExcelDna.Integration" switchValue="All">
        <listeners>
          <remove name="System.Diagnostics.DefaultTraceListener" />
          <add name="LogDisplay" type="ExcelDna.Logging.LogDisplayTraceListener,ExcelDna.Integration">
            <!-- EventTypeFilter takes a SourceLevel as the initializeData:
                    Off, Critical, Error, Warning (default), Information, Verbose, All -->
            <filter type="System.Diagnostics.EventTypeFilter" initializeData="Warning" />
          </add>
        </listeners>
      </source>
    </sources>
  </system.diagnostics>
  <appSettings file="your.Central.Configfile.Path"> : This is a redirection to a central config file containing the same information below
    <add key="LocalHelp" value="C:\YourPathToLocalHelp\LocalHelp.htm" /> : If you download this page to your local site, put it there.
    <add key="ExePathR" value="C:\Program Files\R\R-4.0.4\bin\x64\Rscript.exe" /> : The Executable Path used for R
    <add key="FSuffixR" value=".R" />
    <add key="ExePathPython" value="C:\Users\rolan\anaconda3\pythonw.exe" /> : The Executable Path used for Python
    <add key="PathAddPython" value="C:\Users\rolan\anaconda3\Scripts;C:\Users\rolan\anaconda3\Library\bin;C:\Users\rolan\anaconda3\Library\bin;C:\Users\rolan\anaconda3\Library\usr\bin;C:\Users\rolan\anaconda3\Library\mingw-w64\bin" /> : Additional Path Setting for Python
    <add key="FSuffixPython" value=".py" />
    <add key="ExePathPerl" value="C:\Strawberry\perl\bin\perl.exe" /> : The Executable Path used for Perl
    <add key="PathAddPerl" value="C:\Strawberry\c\bin;C:\Strawberry\perl\site\bin;C:\Strawberry\perl\bin" /> : Additional Path Setting for Perl
    <add key="FSuffixPerl" value=".pl" />
    <add key="ExePathWinCmd" value="C:\Windows\System32\cmd.exe" />
    <add key="FSuffixWinCmd" value=".cmd" />
    <add key="ExePathCscript" value="C:\Windows\System32\cscript.exe" />
    <add key="FSuffixCscript" value=".vbs" />
    <add key="presetSheetButtonsCount" value="24"/> : the preset maximum Button Count for Sheets (if you expect more sheets with ScriptDefinitions, you can set it accordingly)
  </appSettings>
```

# Building

All packages necessary for building are contained, simply open ScriptAddin.sln and build the solution. The script deployForTest.cmd can be used to deploy the built xll and config to %appdata%\Microsoft\AddIns
