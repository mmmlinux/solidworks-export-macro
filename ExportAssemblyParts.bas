' =====================================================================
'  ExportAssemblyParts  -  SOLIDWORKS 2026 VBA macro
'
'  Run with an assembly open. For every part used in the assembly
'  (all levels, including parts inside subassemblies) it exports:
'     <PartName>-<Config>.step
'     <PartName>-<Config>.png   (isometric, zoom to fit)
'     <PartName>.pdf            (if a matching drawing file exists)
'  The top-level assembly and every subassembly get the same treatment.
'  An assembly STEP is always written as one file: the "Export assembly
'  components as separate STEP files" option is turned off for the run and
'  restored afterwards.
'
'  Run on a single part instead and it exports every configuration that
'  part has, since there is no assembly to say which one matters.
'
'  The configuration name is only added to a file name when the model has
'  more than one configuration.
'
'  Only the configuration(s) actually referenced by the assembly are
'  exported. Each model/config pair is exported once, however many
'  instances there are. Suppressed components are skipped silently, as is
'  anything inside a suppressed subassembly.
'
'  Output goes next to the model file it came from, in an Export folder
'  with a STEP, PNG and PDF subfolder. Turn that off and the files land in
'  the model's own folder instead.
'  Virtual parts (stored inside the assembly) have no folder of their own,
'  so theirs go next to the assembly file.
'
'  A settings dialog appears at the start of each run so any of this can be
'  changed for that run without editing the macro. The constants under
'  Settings are the starting values; set SHOW_SETTINGS_DIALOG to False to
'  skip the dialog and always use them.
'
'  BOM tables found on the drawings are collected as the run goes, then
'  offered as a list at the end so you can pick which to write out as CSV.
'  The drawings are only opened once for this.
'
'  Two UserForms are needed in the same project, each holding nothing but
'  two CommandButtons named cmdOK and cmdCancel:
'     frmSettings   the settings dialog
'     frmPickBoms   the list of BOM tables
'  Every other control on them is created at run time.
'
'  A CSV log (<AssemblyName>_export_log.csv) is written next to the
'  assembly, one row per output file, for downstream scripts to read.
'
'  The macro switches a model's active configuration temporarily and puts
'  it back afterwards. It never saves the parts, subassemblies or assembly.
'
'  The component tree is read in full before any document is touched,
'  because component pointers can go stale once the assembly rebuilds.
'  Document pointers are kept instead; those stay valid as long as the
'  assembly holds the model loaded. Every export is error-trapped, so one
'  bad model is logged and skipped rather than stopping the run.
' =====================================================================
Option Explicit

' ---- Settings -------------------------------------------------------
' These are the defaults. Unless SHOW_SETTINGS_DIALOG is False, a dialog
' appears at the start of each run so they can be changed for that run
' without editing the macro. Change them here to move the starting point.

' Show the settings dialog on every run? False runs straight from these
' defaults with no prompting.
Const SHOW_SETTINGS_DIALOG As Boolean = True

Const DEF_EXPORT_STEP As Boolean = True              ' STEP for parts
Const DEF_ASSEMBLY_STEP As Boolean = True            ' STEP for assemblies, always as one file
Const DEF_ASSEMBLY_STEP_TOP_ONLY As Boolean = False  ' True = only the assembly you ran this on,
                                                     ' not every subassembly as well
Const DEF_EXPORT_PNG As Boolean = True
Const DEF_EXPORT_DRAWING_PDF As Boolean = True
Const DEF_EXPORT_BOM As Boolean = True               ' BOM tables from drawings, as CSV
Const DEF_INCLUDE_SUBASSEMBLIES As Boolean = True    ' include subassemblies at all
Const DEF_INCLUDE_TOP_ASSEMBLY As Boolean = True     ' include the assembly you ran this on
Const DEF_CONFIG_IN_NAME As Boolean = True           ' only for models with more than one configuration
' True  = put output in <model folder>\Export\STEP, \PNG and \PDF.
' False = put output straight into the model's own folder.
' Folders are created as needed; if one can't be created the file falls
' back to the model's folder rather than being lost.
Const DEF_USE_SUBFOLDERS As Boolean = True
Const DEF_EXPORT_FOLDER_NAME As String = "Export"
' Set True while troubleshooting: the macro stops on the offending line
' instead of logging the error and carrying on.
Const DEF_STOP_ON_ERROR As Boolean = False
' ---------------------------------------------------------------------

' The live settings for this run, seeded from the defaults above and then
' optionally changed in the dialog.
Dim optExportStep As Boolean
Dim optAssemblyStep As Boolean
Dim optAssemblyStepTopOnly As Boolean
Dim optExportPng As Boolean
Dim optExportPdf As Boolean
Dim optExportBom As Boolean
Dim optIncludeSubassemblies As Boolean
Dim optIncludeTopAssembly As Boolean
Dim optConfigInName As Boolean
Dim optUseSubfolders As Boolean
Dim optStopOnError As Boolean
Dim optExportFolderName As String

Dim swApp As SldWorks.SldWorks
Dim rows As Collection         ' CSV rows, each an array of fields
Dim usedNames As Object
Dim drawingsDone As Object
Dim bomList As Collection      ' BOM tables found while drawings were open
Dim okCount As Long
Dim failCount As Long
Dim curStep As String          ' what the macro is doing, for error messages
Dim rootPath As String         ' the assembly being exported - never close this
Dim rootTitle As String

Sub main()
    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "Open an assembly or a part first.", vbExclamation
        Exit Sub
    End If

    Dim docType As Long
    docType = swModel.GetType
    If docType <> swDocumentTypes_e.swDocASSEMBLY And docType <> swDocumentTypes_e.swDocPART Then
        MsgBox "Run this on an assembly or a part.", vbExclamation
        Exit Sub
    End If

    Dim asmPath As String
    asmPath = swModel.GetPathName
    If asmPath = "" Then
        MsgBox "Save the document first.", vbExclamation
        Exit Sub
    End If

    ' Settings for this run: the defaults, then whatever the dialog says.
    LoadDefaultSettings
    If SHOW_SETTINGS_DIALOG Then
        If Not ShowSettingsDialog(docType = swDocumentTypes_e.swDocASSEMBLY) Then Exit Sub
    End If

    ' Used for the log file and for virtual parts.
    Dim asmDir As String
    asmDir = Left(asmPath, InStrRev(asmPath, "\"))

    Dim asmTitle As String
    asmTitle = swModel.GetTitle

    ' Remembered so the macro can never close the document it is exporting.
    rootPath = asmPath
    rootTitle = asmTitle

    Set usedNames = CreateObject("Scripting.Dictionary")
    Set drawingsDone = CreateObject("Scripting.Dictionary")
    Set rows = New Collection
    Set bomList = New Collection
    okCount = 0
    failCount = 0

    Dim startTime As Double
    startTime = Timer

    ' ---------------- Pass 1: work out what to export ---------------
    ' Nothing is opened or activated here. Each entry is an array of the
    ' document, the configuration to export, and a name for the log.
    Dim items As Collection

    If docType = swDocumentTypes_e.swDocASSEMBLY Then
        Dim swAssy As SldWorks.AssemblyDoc
        Set swAssy = swModel
        ' Lightweight components have no model loaded, so resolve them all.
        swAssy.ResolveAllLightWeightComponents False
        Set items = CollectItems(swAssy)
    Else
        ' Run on a part directly: export every configuration it has, since
        ' there's no assembly here to say which one matters.
        Set items = CollectConfigs(swModel)
    End If

    If items.Count = 0 Then
        MsgBox "Nothing found to export.", vbInformation
        Exit Sub
    End If

    ' An assembly STEP must be one file, whatever the user's own setting is.
    ' "Export assembly components as separate STEP files" is turned off for
    ' the run and put back at the end.
    Dim origAtomic As Boolean
    Dim atomicChanged As Boolean
    If optAssemblyStep Then
        On Error Resume Next
        origAtomic = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swStepExportAtomicSave)
        If Err.Number = 0 Then
            If origAtomic Then
                swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swStepExportAtomicSave, False
                atomicChanged = (Err.Number = 0)
            End If
        End If
        Err.Clear
        On Error GoTo 0
    End If

    ' ---------------- Pass 2: export --------------------------------
    Dim i As Long
    Dim entry As Variant
    Dim modelDoc As SldWorks.ModelDoc2
    Dim cfg As String
    Dim label As String

    For i = 1 To items.Count
        entry = items(i)
        Set modelDoc = entry(0)
        cfg = entry(1)
        label = entry(2)

        curStep = "starting"

        ' Make sure the document being exported is still open. If something
        ' closed it, stop rather than failing on every remaining item.
        If Not RootStillOpen() Then
            AddRow "", label, cfg, "", "", "", "error", _
                   "the document being exported is no longer open - run stopped here"
            failCount = failCount + 1
            Exit For
        End If

        If optStopOnError Then
            ExportModel modelDoc, cfg, asmDir
        Else
            On Error Resume Next
            ExportModel modelDoc, cfg, asmDir
            If Err.Number <> 0 Then
                AddRow "", label, cfg, "", "", "", "error", _
                       "while " & curStep & ": " & Err.Number & " - " & Err.Description
                failCount = failCount + 1
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next i

    ' Put the STEP export setting back the way it was.
    If atomicChanged Then
        On Error Resume Next
        swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swStepExportAtomicSave, origAtomic
        Err.Clear
        On Error GoTo 0
    End If

    ' Bring the assembly back to the front.
    Dim errs As Long
    On Error Resume Next
    swApp.ActivateDoc3 asmTitle, False, swRebuildOnActivation_e.swDontRebuildActiveDoc, errs
    On Error GoTo 0

    ' Any BOM tables found along the way are offered as a list now. The
    ' drawings were read once during the run, so nothing is reopened here.
    Dim pickSeconds As Double
    If optExportBom And bomList.Count > 0 Then
        Dim pickStart As Double
        pickStart = Timer

        Dim chosen As Collection
        Set chosen = PickBoms(bomList)

        pickSeconds = Timer - pickStart
        If pickSeconds < 0 Then pickSeconds = pickSeconds + 86400

        Dim bi As Long
        Dim bomEntry As Variant
        For bi = 1 To chosen.Count
            ' Copied into a variable first: passing a Collection item
            ' directly into a parameter is a ByRef mismatch in VBA.
            bomEntry = chosen(bi)
            If optStopOnError Then
                WriteBomCsv bomEntry
            Else
                On Error Resume Next
                WriteBomCsv bomEntry
                If Err.Number <> 0 Then
                    AddRow "", "BOM " & bi, "", "", "csv", "", "error", _
                           Err.Number & " - " & Err.Description
                    failCount = failCount + 1
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        Next bi
    End If

    Dim asmName As String
    asmName = StripExtension(Mid(asmPath, InStrRev(asmPath, "\") + 1))

    Dim logPath As String
    logPath = asmDir & asmName & "_export_log.csv"

    WriteCsv logPath

    Dim elapsed As Double
    elapsed = Timer - startTime
    If elapsed < 0 Then elapsed = elapsed + 86400   ' the run crossed midnight
    elapsed = elapsed - pickSeconds                 ' waiting on the BOM list isn't work
    If elapsed < 0 Then elapsed = 0

    Dim perFile As String
    If okCount > 0 Then perFile = "  (" & Format(elapsed / okCount, "0.0") & "s per file)"

    MsgBox okCount & " file(s) exported, " & failCount & " problem(s)." & vbCrLf & _
           "Took " & FormatElapsed(elapsed) & perFile & vbCrLf & vbCrLf & _
           "Output was saved next to each model file." & vbCrLf & "Log: " & logPath, _
           IIf(failCount = 0, vbInformation, vbExclamation)
End Sub

' ---- Settings dialog ------------------------------------------------

Private Sub LoadDefaultSettings()
    optExportStep = DEF_EXPORT_STEP
    optAssemblyStep = DEF_ASSEMBLY_STEP
    optAssemblyStepTopOnly = DEF_ASSEMBLY_STEP_TOP_ONLY
    optExportPng = DEF_EXPORT_PNG
    optExportPdf = DEF_EXPORT_DRAWING_PDF
    optExportBom = DEF_EXPORT_BOM
    optIncludeSubassemblies = DEF_INCLUDE_SUBASSEMBLIES
    optIncludeTopAssembly = DEF_INCLUDE_TOP_ASSEMBLY
    optConfigInName = DEF_CONFIG_IN_NAME
    optUseSubfolders = DEF_USE_SUBFOLDERS
    optStopOnError = DEF_STOP_ON_ERROR
    optExportFolderName = DEF_EXPORT_FOLDER_NAME
End Sub

' Shows the settings dialog. Returns False if the user cancels, in which
' case nothing is exported.
'
' frmSettings is a blank UserForm carrying only two buttons, cmdOK and
' cmdCancel. Every other control is built at run time by
' BuildSettingsControls below, because a UserForm's controls live in a
' binary .frx file that can't be written by hand.
Private Function ShowSettingsDialog(isAssembly As Boolean) As Boolean
    Dim f As frmSettings
    Set f = New frmSettings

    ' Seed the dialog from the current settings.
    f.Controls("chkStep").Value = optExportStep
    f.Controls("chkAsmStep").Value = optAssemblyStep
    f.Controls("chkTopOnly").Value = optAssemblyStepTopOnly
    f.Controls("chkPng").Value = optExportPng
    f.Controls("chkPdf").Value = optExportPdf
    f.Controls("chkBom").Value = optExportBom
    f.Controls("chkSubAsm").Value = optIncludeSubassemblies
    f.Controls("chkTopAsm").Value = optIncludeTopAssembly
    f.Controls("chkCfgName").Value = optConfigInName
    f.Controls("chkFolders").Value = optUseSubfolders
    f.Controls("chkDebug").Value = optStopOnError
    f.Controls("txtFolder").Text = optExportFolderName

    If isAssembly Then
        f.Controls("lblNote").Caption = "Running on an assembly: each model is exported " & _
                                        "in the configuration the assembly uses."
    Else
        f.Controls("lblNote").Caption = "Running on a part: every configuration of it " & _
                                        "will be exported."
    End If

    f.Show vbModal

    If Not f.Confirmed Then
        Unload f
        Exit Function
    End If

    ' Read them back.
    optExportStep = f.Controls("chkStep").Value
    optAssemblyStep = f.Controls("chkAsmStep").Value
    optAssemblyStepTopOnly = f.Controls("chkTopOnly").Value
    optExportPng = f.Controls("chkPng").Value
    optExportPdf = f.Controls("chkPdf").Value
    optExportBom = f.Controls("chkBom").Value
    optIncludeSubassemblies = f.Controls("chkSubAsm").Value
    optIncludeTopAssembly = f.Controls("chkTopAsm").Value
    optConfigInName = f.Controls("chkCfgName").Value
    optUseSubfolders = f.Controls("chkFolders").Value
    optStopOnError = f.Controls("chkDebug").Value
    optExportFolderName = SafeFileName(Trim(f.Controls("txtFolder").Text))

    Unload f

    ' A folder name is required once sorting into subfolders is on.
    If optUseSubfolders And optExportFolderName = "" Then
        optExportFolderName = DEF_EXPORT_FOLDER_NAME
    End If

    ShowSettingsDialog = True
End Function

' Builds every control on the settings form. Called from the form's
' Initialize event. Keeping it here means the form's own code module stays
' short enough to paste in one go.
Public Sub BuildSettingsControls(f As Object)
    Const PAD As Single = 12
    Const ROW As Single = 15
    Const GAP As Single = 8
    Const CW As Single = 330          ' content width

    Dim y As Single
    y = PAD

    f.Caption = "Export Settings"

    AddHeading f, "lblH1", "What to export", PAD, y, CW:                        y = y + 16
    AddCheck f, "chkStep", "STEP file for each part", PAD + 6, y, CW:           y = y + ROW
    AddCheck f, "chkAsmStep", "STEP file for each assembly (always one file)", _
                PAD + 6, y, CW:                                                 y = y + ROW
    AddCheck f, "chkTopOnly", "Top-level assembly only, not every subassembly", _
                PAD + 22, y, CW - 22:                                           y = y + ROW
    AddCheck f, "chkPng", "Isometric PNG for each model", PAD + 6, y, CW:       y = y + ROW
    AddCheck f, "chkPdf", "PDF of the drawing, where one exists", PAD + 6, y, CW: y = y + ROW
    AddCheck f, "chkBom", "BOM tables from drawings, as CSV (pick them later)", _
                PAD + 6, y, CW
    y = y + ROW + GAP

    AddHeading f, "lblH2", "What to include", PAD, y, CW:                       y = y + 16
    AddCheck f, "chkSubAsm", "Subassemblies", PAD + 6, y, CW:                   y = y + ROW
    AddCheck f, "chkTopAsm", "The top-level assembly itself", PAD + 6, y, CW
    y = y + ROW + GAP

    AddHeading f, "lblH3", "Naming and folders", PAD, y, CW:                    y = y + 16
    AddCheck f, "chkCfgName", "Add the configuration to file names", _
                PAD + 6, y, CW:                                                 y = y + ROW
    AddCheck f, "chkFolders", "Sort output into Export\STEP, \PNG and \PDF", _
                PAD + 6, y, CW
    y = y + ROW + 6

    AddLabel f, "lblFolder", "Export folder", PAD + 6, y + 3, 110
    AddText f, "txtFolder", PAD + 118, y, CW - 118
    y = y + 22 + GAP

    AddHeading f, "lblH4", "Troubleshooting", PAD, y, CW:                       y = y + 16
    AddCheck f, "chkDebug", "Stop on the offending line instead of logging it", _
                PAD + 6, y, CW
    y = y + ROW + GAP

    AddNote f, "lblNote", "", PAD, y, CW
    y = y + 26

    ' The two buttons already exist on the form; just place them.
    With f.Controls("cmdOK")
        .Caption = "Export"
        .Width = 72: .Height = 22
        .Left = PAD + CW - 72 - 78
        .Top = y
        .Default = True
    End With
    With f.Controls("cmdCancel")
        .Caption = "Cancel"
        .Width = 72: .Height = 22
        .Left = PAD + CW - 72
        .Top = y
        .Cancel = True
    End With

    y = y + 22 + PAD

    f.Width = CW + PAD * 2 + 8
    f.Height = y + 24                  ' leave room for the title bar
End Sub

Private Sub AddCheck(f As Object, name As String, caption As String, _
                     x As Single, y As Single, w As Single)
    Dim c As Object
    Set c = f.Controls.Add("Forms.CheckBox.1", name, True)
    c.Caption = caption
    c.Left = x: c.Top = y: c.Width = w: c.Height = 14
    c.WordWrap = False
End Sub

Private Sub AddHeading(f As Object, name As String, caption As String, _
                       x As Single, y As Single, w As Single)
    Dim c As Object
    Set c = f.Controls.Add("Forms.Label.1", name, True)
    c.Caption = caption
    c.Left = x: c.Top = y: c.Width = w: c.Height = 14
    c.Font.Bold = True
End Sub

Private Sub AddLabel(f As Object, name As String, caption As String, _
                     x As Single, y As Single, w As Single)
    Dim c As Object
    Set c = f.Controls.Add("Forms.Label.1", name, True)
    c.Caption = caption
    c.Left = x: c.Top = y: c.Width = w: c.Height = 14
End Sub

Private Sub AddNote(f As Object, name As String, caption As String, _
                    x As Single, y As Single, w As Single)
    Dim c As Object
    Set c = f.Controls.Add("Forms.Label.1", name, True)
    c.Caption = caption
    c.Left = x: c.Top = y: c.Width = w: c.Height = 24
    c.Font.Size = 8
    c.ForeColor = RGB(100, 100, 100)
End Sub

Private Sub AddText(f As Object, name As String, _
                    x As Single, y As Single, w As Single)
    Dim c As Object
    Set c = f.Controls.Add("Forms.TextBox.1", name, True)
    c.Left = x: c.Top = y: c.Width = w: c.Height = 18
End Sub

' Shows the harvested BOM tables and returns the ones to write. Nothing is
' ticked to begin with, so a CSV is only written where it's asked for.
' Cancelling returns an empty list, which just means no CSVs.
Private Function PickBoms(found As Collection) As Collection
    Dim result As New Collection
    Set PickBoms = result

    Dim f As frmPickBoms
    Set f = New frmPickBoms

    Dim lst As Object
    Set lst = f.Controls("lstBoms")

    Dim i As Long
    Dim rec As Variant
    For i = 1 To found.Count
        ' Same reason as above: take the item into a variable rather than
        ' writing found(i)(0), which VBA won't parse reliably.
        rec = found(i)
        lst.AddItem rec(0)
    Next i

    f.Controls("lblCount").Caption = found.Count & _
        IIf(found.Count = 1, " BOM table found.", " BOM tables found.") & _
        "  Tick the ones you want as CSV."

    f.Show vbModal

    If f.Confirmed Then
        For i = 1 To found.Count
            If lst.Selected(i - 1) Then result.Add found(i)
        Next i
    End If

    Unload f
End Function

' Builds the BOM picker's controls. Like the settings form, frmPickBoms is
' a blank UserForm carrying only cmdOK and cmdCancel.
Public Sub BuildBomPickerControls(f As Object)
    Const PAD As Single = 12
    Const CW As Single = 380

    f.Caption = "BOM Tables"

    AddLabel f, "lblCount", "", PAD, PAD, CW

    Dim c As Object
    Set c = f.Controls.Add("Forms.ListBox.1", "lstBoms", True)
    c.Left = PAD
    c.Top = PAD + 20
    c.Width = CW
    c.Height = 200
    c.MultiSelect = 1        ' fmMultiSelectMulti
    c.ListStyle = 1          ' fmListStyleOption - draws a checkbox per row

    Dim y As Single
    y = PAD + 20 + 200 + 10

    With f.Controls("cmdOK")
        .Caption = "Export CSVs"
        .Width = 88: .Height = 22
        .Left = PAD + CW - 88 - 94
        .Top = y
        .Default = True
    End With
    With f.Controls("cmdCancel")
        .Caption = "Skip"
        .Width = 88: .Height = 22
        .Left = PAD + CW - 88
        .Top = y
        .Cancel = True
    End With

    f.Width = CW + PAD * 2 + 8
    f.Height = y + 22 + PAD + 24
End Sub

' Every configuration of a single model, for when the macro is run on a
' part directly rather than on an assembly.
Private Function CollectConfigs(swMdl As SldWorks.ModelDoc2) As Collection
    Dim result As New Collection
    Set CollectConfigs = result

    Dim vNames As Variant
    vNames = swMdl.GetConfigurationNames

    If IsEmpty(vNames) Then
        ' Shouldn't happen, but fall back to whatever is active.
        result.Add Array(swMdl, swMdl.ConfigurationManager.ActiveConfiguration.Name, swMdl.GetTitle)
        Exit Function
    End If

    Dim i As Long
    For i = 0 To UBound(vNames)
        result.Add Array(swMdl, CStr(vNames(i)), swMdl.GetTitle)
    Next i
End Function

' Walks every component and records the model/config pairs worth exporting.
Private Function CollectItems(swAssy As SldWorks.AssemblyDoc) As Collection
    Dim result As New Collection
    Set CollectItems = result

    Dim done As Object
    Set done = CreateObject("Scripting.Dictionary")

    ' The top-level assembly itself is not one of its own components,
    ' so add it explicitly, first.
    If optIncludeTopAssembly Then
        Dim rootModel As SldWorks.ModelDoc2
        Set rootModel = swAssy
        Dim rootCfg As String
        rootCfg = rootModel.ConfigurationManager.ActiveConfiguration.Name
        result.Add Array(rootModel, rootCfg, rootModel.GetTitle)
        done.Add LCase(rootModel.GetPathName & "|" & rootModel.GetTitle) & "|" & rootCfg, True
    End If

    Dim vComps As Variant
    vComps = swAssy.GetComponents(False)      ' False = every level
    If IsEmpty(vComps) Then Exit Function

    Dim i As Long
    Dim swComp As SldWorks.Component2
    Dim compModel As SldWorks.ModelDoc2
    Dim cfgName As String, key As String, cPath As String
    Dim docType As Long

    For i = 0 To UBound(vComps)
        On Error Resume Next

        Set swComp = vComps(i)
        If Err.Number <> 0 Then Err.Clear: GoTo NextComp
        If swComp Is Nothing Then GoTo NextComp

        ' Skip suppressed components, and anything inside a suppressed subassembly.
        If IsSuppressedAnywhere(swComp) Then GoTo NextComp

        cPath = swComp.GetPathName

        Set compModel = swComp.GetModelDoc2
        If compModel Is Nothing Then
            ' Not something we can reach (e.g. file missing).
            If LCase(Right(cPath, 7)) = ".sldprt" Or LCase(Right(cPath, 7)) = ".sldasm" Then
                AddRow "", swComp.Name2, "", cPath, "", "", "skipped", "could not load"
                failCount = failCount + 1
            End If
            GoTo NextComp
        End If

        docType = compModel.GetType
        If docType <> swDocumentTypes_e.swDocPART Then
            If docType <> swDocumentTypes_e.swDocASSEMBLY Then GoTo NextComp
            If Not optIncludeSubassemblies Then GoTo NextComp
        End If

        cfgName = swComp.ReferencedConfiguration
        If cfgName = "" Then cfgName = compModel.ConfigurationManager.ActiveConfiguration.Name

        key = LCase(cPath & "|" & compModel.GetTitle) & "|" & cfgName
        If done.Exists(key) Then GoTo NextComp
        done.Add key, True

        result.Add Array(compModel, cfgName, compModel.GetTitle)

NextComp:
        If Err.Number <> 0 Then
            AddRow "", "component " & i, "", "", "", "", "error", _
                   Err.Number & " - " & Err.Description
            failCount = failCount + 1
            Err.Clear
        End If
        On Error GoTo 0
    Next i
End Function

Private Sub ExportModel(compModel As SldWorks.ModelDoc2, cfgName As String, fallbackDir As String)
    Dim errs As Long, warns As Long
    Dim wasVisible As Boolean

    curStep = "reading model visibility"
    wasVisible = compModel.Visible

    curStep = "reading the model path"
    Dim modelPath As String
    Dim isVirtual As Boolean
    Dim isAssembly As Boolean
    Dim srcType As String
    modelPath = compModel.GetPathName
    isVirtual = (modelPath = "" Or InStr(compModel.GetTitle, "^") > 0)
    isAssembly = (compModel.GetType = swDocumentTypes_e.swDocASSEMBLY)
    srcType = IIf(isAssembly, "assembly", "part")

    ' Bring the model up in its own window (needed for the PNG).
    curStep = "activating the model"
    Dim swMdl As SldWorks.ModelDoc2
    Set swMdl = swApp.ActivateDoc3(compModel.GetTitle, False, swRebuildOnActivation_e.swDontRebuildActiveDoc, errs)
    If swMdl Is Nothing And Not isVirtual Then
        curStep = "opening the model file"
        Dim openType As Long
        If isAssembly Then
            openType = swDocumentTypes_e.swDocASSEMBLY
        Else
            openType = swDocumentTypes_e.swDocPART
        End If
        Set swMdl = swApp.OpenDoc6(modelPath, openType, _
                                   swOpenDocOptions_e.swOpenDocOptions_Silent, cfgName, errs, warns)
        If Not swMdl Is Nothing Then
            Set swMdl = swApp.ActivateDoc3(swMdl.GetTitle, False, swRebuildOnActivation_e.swDontRebuildActiveDoc, errs)
        End If
    End If
    If swMdl Is Nothing Then
        AddRow srcType, compModel.GetTitle, cfgName, modelPath, "", "", "failed", "could not open"
        failCount = failCount + 1
        Exit Sub
    End If

    Dim title As String
    title = swMdl.GetTitle

    ' Switch to the configuration the assembly uses.
    curStep = "reading the active configuration"
    Dim origCfg As String
    origCfg = swMdl.ConfigurationManager.ActiveConfiguration.Name
    If origCfg <> cfgName Then
        curStep = "switching configuration to " & cfgName
        If Not swMdl.ShowConfiguration2(cfgName) Then
            AddRow srcType, title, cfgName, modelPath, "", "", "failed", "could not activate configuration"
            failCount = failCount + 1
            GoTo Cleanup
        End If
        ' Only rebuild when the configuration actually changed; forcing a
        ' rebuild on a large assembly for no reason is slow.
        curStep = "rebuilding after the configuration switch"
        swMdl.ForceRebuild3 True
    End If

    ' Save next to the model file; virtual parts go next to the assembly.
    Dim outDir As String
    If isVirtual Then
        outDir = fallbackDir
    Else
        outDir = Left(modelPath, InStrRev(modelPath, "\"))
    End If

    ' Build a unique, file-safe base name within that folder.
    curStep = "building the output name"
    Dim baseName As String
    Dim cfgCount As Long
    cfgCount = swMdl.GetConfigurationCount
    baseName = ModelBaseName(swMdl)
    ' Only worth naming the configuration when there's more than one to tell apart.
    If optConfigInName And cfgCount > 1 Then baseName = baseName & "-" & cfgName

    ' Does this model get a STEP? Parts and assemblies are controlled separately,
    ' and assembly STEPs can be limited to the top-level assembly.
    Dim doStep As Boolean
    If isAssembly Then
        doStep = optAssemblyStep
        If doStep And optAssemblyStepTopOnly Then
            doStep = (modelPath <> "" And LCase(modelPath) = LCase(rootPath))
        End If
    Else
        doStep = optExportStep
    End If

    ' Reserve only the extensions this model will actually write, so a PDF
    ' from the drawing can still use the same base name.
    Dim wantExts As String
    If doStep Then wantExts = "step"
    If optExportPng Then wantExts = wantExts & IIf(wantExts = "", "", ",") & "png"
    baseName = UniqueName(outDir, SafeFileName(baseName), wantExts)

    Dim opts As Long
    opts = swSaveAsOptions_e.swSaveAsOptions_Silent Or swSaveAsOptions_e.swSaveAsOptions_Copy

    If doStep Then
        curStep = "saving the STEP file"
        Dim stepPath As String
        stepPath = OutFolder(outDir, "step") & baseName & ".step"
        errs = 0: warns = 0
        If swMdl.Extension.SaveAs3(stepPath, swSaveAsVersion_e.swSaveAsCurrentVersion, _
                                   opts, Nothing, Nothing, errs, warns) Then
            AddRow srcType, title, cfgName, modelPath, "step", stepPath, "ok", ""
            okCount = okCount + 1
        Else
            AddRow srcType, title, cfgName, modelPath, "step", stepPath, "failed", "SaveAs error " & errs
            failCount = failCount + 1
        End If
    End If

    If optExportPng Then
        curStep = "setting the isometric view"
        swMdl.ShowNamedView2 "*Isometric", swStandardViews_e.swIsometricView
        swMdl.ViewZoomtofit2
        swMdl.GraphicsRedraw2
        curStep = "saving the PNG file"
        Dim pngPath As String
        pngPath = OutFolder(outDir, "png") & baseName & ".png"
        errs = 0: warns = 0
        If swMdl.Extension.SaveAs3(pngPath, swSaveAsVersion_e.swSaveAsCurrentVersion, _
                                   opts, Nothing, Nothing, errs, warns) Then
            AddRow srcType, title, cfgName, modelPath, "png", pngPath, "ok", ""
            okCount = okCount + 1
        Else
            AddRow srcType, title, cfgName, modelPath, "png", pngPath, "failed", "SaveAs error " & errs
            failCount = failCount + 1
        End If
    End If

Cleanup:
    ' Put the model's active configuration back the way it was.
    curStep = "restoring the original configuration"
    If swMdl.ConfigurationManager.ActiveConfiguration.Name <> origCfg Then
        swMdl.ShowConfiguration2 origCfg
    End If
    ' Close the window only if we opened it (the model stays loaded by the assembly).
    curStep = "closing the model window"
    If Not wasVisible Then CloseModelSafely swMdl

    ' A drawing covers the whole model, so it's handled once per model file,
    ' not once per configuration. The drawing is opened a single time and
    ' both the PDF and the BOM harvest happen while it's open. Trapped
    ' separately so a bad drawing cannot take the model export down with it.
    If (optExportPdf Or optExportBom) And Not isVirtual Then
        If Not drawingsDone.Exists(LCase(modelPath)) Then
            drawingsDone.Add LCase(modelPath), True
            curStep = "handling the drawing"
            If optStopOnError Then
                ProcessDrawing modelPath, srcType, title, cfgName
            Else
                On Error Resume Next
                ProcessDrawing modelPath, srcType, title, cfgName
                If Err.Number <> 0 Then
                    AddRow srcType, title, cfgName, modelPath, "pdf", "", "error", _
                           Err.Number & " - " & Err.Description
                    failCount = failCount + 1
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    End If
End Sub

' Opens the model's drawing once and does everything that needs it: the
' PDF, and reading any BOM tables into memory so they can be picked from a
' list at the end of the run without reopening the file. Does nothing if
' there's no drawing.
Private Sub ProcessDrawing(modelPath As String, srcType As String, title As String, cfgName As String)
    Dim drwPath As String
    drwPath = FindDrawing(modelPath)

    ' Logged rather than passed over in silence, so a missing PDF or BOM can
    ' be told apart from a drawing that was never there to begin with.
    If drwPath = "" Then
        AddRow srcType, title, cfgName, modelPath, "", "", "skipped", _
               "no drawing named like the model beside it"
        Exit Sub
    End If

    Dim modelDir As String, baseName As String
    modelDir = Left(modelPath, InStrRev(modelPath, "\"))
    baseName = StripExtension(Mid(modelPath, InStrRev(modelPath, "\") + 1))

    ' If the drawing is already open, leave it open and untouched afterwards.
    Dim alreadyOpen As Boolean
    Dim swDrw As SldWorks.ModelDoc2
    On Error Resume Next
    Set swDrw = swApp.GetOpenDocumentByName2(drwPath)
    Err.Clear
    On Error GoTo 0
    alreadyOpen = Not (swDrw Is Nothing)

    Dim errs As Long, warns As Long
    If Not alreadyOpen Then
        Set swDrw = swApp.OpenDoc6(drwPath, swDocumentTypes_e.swDocDRAWING, _
                                   swOpenDocOptions_e.swOpenDocOptions_Silent Or _
                                   swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", errs, warns)
    End If

    ' The drawing is the source of these rows, not the model, so it's the
    ' drawing's path that goes in the source columns.
    If swDrw Is Nothing Then
        AddRow srcType, title, cfgName, drwPath, "pdf", "", "failed", "could not open drawing"
        failCount = failCount + 1
        Exit Sub
    End If

    If optExportPdf Then
        curStep = "saving the drawing PDF"
        Dim pdfPath As String
        ' Named against the model's folder so it can't clash with the STEP or
        ' PNG, but written into the PDF subfolder when those are turned on.
        pdfPath = OutFolder(modelDir, "pdf") & UniqueName(modelDir, SafeFileName(baseName), "pdf") & ".pdf"

        ' Ask for every sheet rather than whatever the current PDF export option says.
        Dim pdfData As SldWorks.ExportPdfData
        Set pdfData = swApp.GetExportFileData(swExportDataFileType_e.swExportPdfData)
        If Not pdfData Is Nothing Then
            pdfData.SetSheets swExportDataSheetsToExport_e.swExportData_ExportAllSheets, Empty
            pdfData.ViewPdfAfterSaving = False
        End If

        Dim opts As Long
        opts = swSaveAsOptions_e.swSaveAsOptions_Silent Or swSaveAsOptions_e.swSaveAsOptions_Copy

        Dim e As Long, w As Long
        If swDrw.Extension.SaveAs3(pdfPath, swSaveAsVersion_e.swSaveAsCurrentVersion, _
                                   opts, pdfData, Nothing, e, w) Then
            AddRow srcType, title, cfgName, drwPath, "pdf", pdfPath, "ok", ""
            okCount = okCount + 1
        Else
            AddRow srcType, title, cfgName, drwPath, "pdf", pdfPath, "failed", "SaveAs error " & e
            failCount = failCount + 1
        End If
    End If

    If optExportBom Then
        curStep = "reading BOM tables from the drawing"
        Dim bomsBefore As Long
        bomsBefore = bomList.Count
        HarvestBoms swDrw, drwPath, modelPath, srcType, title, cfgName
        ' Also worth a line: the drawing existed but held no BOM, which is
        ' different from the drawing being missing.
        If bomList.Count = bomsBefore Then
            AddRow srcType, title, cfgName, drwPath, "", "", "skipped", _
                   "drawing has no BOM table"
        End If
    End If

    If Not alreadyOpen Then CloseModelSafely swDrw
End Sub

' Reads every BOM table on a drawing into memory. Nothing is written here -
' the tables are added to bomList so they can be offered as a list once the
' run is over, which is why the drawing only has to be opened once.
Private Sub HarvestBoms(swDrw As SldWorks.ModelDoc2, drwPath As String, modelPath As String, _
                        srcType As String, title As String, cfgName As String)
    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDrw

    Dim origSheetName As String
    On Error Resume Next
    origSheetName = swDrawing.GetCurrentSheet.GetName
    Err.Clear
    On Error GoTo 0

    Dim vSheets As Variant
    vSheets = swDrawing.GetSheetNames
    If IsEmpty(vSheets) Then Exit Sub

    Dim drwName As String
    drwName = StripExtension(Mid(drwPath, InStrRev(drwPath, "\") + 1))

    Dim si As Long, ti As Long
    Dim sheetName As String
    Dim swView As SldWorks.View
    Dim vTables As Variant
    Dim swTable As SldWorks.TableAnnotation
    Dim foundOnSheet As Long

    For si = 0 To UBound(vSheets)
        sheetName = CStr(vSheets(si))

        ' Counted per sheet, not per drawing: the sheet name is in the file
        ' name now, so only a second table on the SAME sheet needs a suffix.
        foundOnSheet = 0

        On Error Resume Next
        swDrawing.ActivateSheet sheetName
        Err.Clear
        On Error GoTo 0

        ' GetFirstView returns the sheet itself, which can own tables too,
        ' so it is included rather than skipped.
        Set swView = swDrawing.GetFirstView
        Do While Not swView Is Nothing
            vTables = Empty
            On Error Resume Next
            vTables = swView.GetTableAnnotations
            Err.Clear
            On Error GoTo 0

            If Not IsEmpty(vTables) Then
                If IsArray(vTables) Then
                    For ti = 0 To UBound(vTables)
                        Set swTable = Nothing
                        On Error Resume Next
                        Set swTable = vTables(ti)
                        Err.Clear
                        On Error GoTo 0

                        If Not swTable Is Nothing Then
                            If IsBomTable(swTable) Then
                                foundOnSheet = foundOnSheet + 1
                                StoreBom swTable, drwPath, drwName, modelPath, _
                                         sheetName, foundOnSheet, srcType, title, cfgName
                            End If
                        End If
                    Next ti
                End If
            End If

            Set swView = swView.GetNextView
        Loop
    Next si

    ' Put the drawing back on the sheet it was showing.
    If origSheetName <> "" Then
        On Error Resume Next
        swDrawing.ActivateSheet origSheetName
        Err.Clear
        On Error GoTo 0
    End If
End Sub

' True if this table annotation is a bill of materials. Tested by trying to
' assign it to a BomTableAnnotation rather than by comparing against a type
' constant, which keeps it working whatever the enum is called.
Private Function IsBomTable(swTable As SldWorks.TableAnnotation) As Boolean
    Dim swBom As SldWorks.BomTableAnnotation
    On Error Resume Next
    Set swBom = swTable
    IsBomTable = (Err.Number = 0) And (Not swBom Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

' Copies a table's cells into memory and records it for the picker.
Private Sub StoreBom(swTable As SldWorks.TableAnnotation, drwPath As String, drwName As String, _
                     modelPath As String, sheetName As String, ByVal indexOnSheet As Long, _
                     srcType As String, title As String, cfgName As String)
    Dim r As Long, c As Long, nRows As Long, nCols As Long

    On Error Resume Next
    nRows = swTable.RowCount
    nCols = swTable.ColumnCount
    If Err.Number <> 0 Or nRows < 1 Or nCols < 1 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Dim cells() As String
    ReDim cells(nRows - 1, nCols - 1)

    For r = 0 To nRows - 1
        For c = 0 To nCols - 1
            On Error Resume Next
            cells(r, c) = swTable.Text(r, c)
            If Err.Number <> 0 Then cells(r, c) = "": Err.Clear
            On Error GoTo 0
        Next c
    Next r

    ' Name the file after the drawing and the sheet the table sits on, so
    ' a multi-sheet drawing gives one clearly named CSV per sheet. The
    ' suffix only appears when a single sheet holds more than one table.
    Dim fileBase As String
    fileBase = drwName & "-" & sheetName
    If indexOnSheet > 1 Then fileBase = fileBase & "-BOM" & indexOnSheet

    Dim label As String
    label = drwName & "   [" & sheetName & "]   " & _
            (nRows - 1) & " rows x " & nCols & " cols"
    If indexOnSheet > 1 Then label = label & "   (table " & indexOnSheet & ")"

    ' label, cells, where the CSV goes, and what to put in the log
    bomList.Add Array(label, cells, modelPath, drwPath, fileBase, srcType, title, cfgName)
End Sub

' Writes one harvested BOM out as a CSV next to its model.
Private Sub WriteBomCsv(ByVal rec As Variant)
    Dim cells As Variant
    cells = rec(1)

    Dim modelPath As String, drwPath As String, fileBase As String
    modelPath = rec(2)
    drwPath = rec(3)
    fileBase = rec(4)

    ' Unpacked into typed variables: AddRow takes Strings, and handing it
    ' rec(5) straight from the Variant array is a ByRef type mismatch.
    Dim srcType As String, title As String, cfgName As String
    srcType = rec(5)
    title = rec(6)
    cfgName = rec(7)

    Dim modelDir As String
    modelDir = Left(modelPath, InStrRev(modelPath, "\"))

    Dim csvPath As String
    csvPath = OutFolder(modelDir, "csv") & UniqueName(modelDir, SafeFileName(fileBase), "csv") & ".csv"

    Dim f As Integer
    Dim r As Long, c As Long
    Dim line As String

    On Error Resume Next
    f = FreeFile
    Open csvPath For Output As #f
    If Err.Number <> 0 Then
        AddRow srcType, title, cfgName, drwPath, "csv", csvPath, "failed", _
               "could not write: " & Err.Description
        failCount = failCount + 1
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    For r = LBound(cells, 1) To UBound(cells, 1)
        line = ""
        For c = LBound(cells, 2) To UBound(cells, 2)
            If c > LBound(cells, 2) Then line = line & ","
            ' Same quoting as the export log, so commas inside a description
            ' don't split the row.
            line = line & CsvField(CStr(cells(r, c)))
        Next c
        Print #f, line
    Next r

    Close #f

    AddRow srcType, title, cfgName, drwPath, "csv", csvPath, "ok", ""
    okCount = okCount + 1
End Sub

' True while the document being exported is still open.
Private Function RootStillOpen() As Boolean
    On Error Resume Next
    Dim d As SldWorks.ModelDoc2
    Set d = swApp.GetOpenDocumentByName2(rootPath)
    RootStillOpen = Not (d Is Nothing)
    If Err.Number <> 0 Then
        Err.Clear
        RootStillOpen = True      ' can't tell, so carry on
    End If
    On Error GoTo 0
End Function

' Closes a document by pointer rather than by name. CloseDoc takes a
' document title, and a drawing usually shares its title with the model it
' documents, so closing by name can shut the wrong document - including the
' assembly this macro is walking. The assembly is never closed.
Private Sub CloseModelSafely(doc As SldWorks.ModelDoc2)
    If doc Is Nothing Then Exit Sub

    On Error Resume Next

    Dim p As String, t As String
    p = doc.GetPathName
    t = doc.GetTitle

    ' Never close the assembly being exported.
    If p <> "" And rootPath <> "" Then
        If LCase(p) = LCase(rootPath) Then Err.Clear: On Error GoTo 0: Exit Sub
    End If

    ' QuitDoc takes the document itself, so there is no name to confuse.
    swApp.QuitDoc doc
    If Err.Number <> 0 Then
        Err.Clear
        ' Fall back to closing by title, but only when that title is
        ' unambiguous - i.e. it is not the assembly's title.
        If LCase(t) <> LCase(rootTitle) Then swApp.CloseDoc t
        Err.Clear
    End If

    On Error GoTo 0
End Sub

' Returns the path of the drawing sitting beside the model and named after
' it, or "" if there isn't one. Works for both .sldprt and .sldasm.
Private Function FindDrawing(modelPath As String) As String
    Dim candidate As String
    candidate = StripExtension(modelPath) & ".slddrw"
    If Dir(candidate) <> "" Then FindDrawing = candidate
End Function

' True if this component, or any subassembly above it, is suppressed.
' A component inside a suppressed subassembly reports its own state as
' resolved, so the parent chain has to be checked as well.
Private Function IsSuppressedAnywhere(swComp As SldWorks.Component2) As Boolean
    Dim c As SldWorks.Component2
    Dim guard As Long

    Set c = swComp
    Do While Not c Is Nothing
        If c.IsSuppressed Then
            IsSuppressedAnywhere = True
            Exit Function
        End If
        Set c = c.GetParent
        guard = guard + 1
        If guard > 100 Then Exit Do          ' paranoia against a bad parent chain
    Loop
    IsSuppressedAnywhere = False
End Function

Private Function ModelBaseName(swMdl As SldWorks.ModelDoc2) As String
    Dim p As String
    p = swMdl.GetPathName
    If p <> "" And InStr(swMdl.GetTitle, "^") = 0 Then
        p = Mid(p, InStrRev(p, "\") + 1)
    Else
        p = swMdl.GetTitle          ' virtual parts
    End If
    ModelBaseName = StripExtension(p)
End Function

' Seconds as something readable: "8.4s", "1m 47s", "1h 04m 09s".
' Whole seconds are worked out first, so a remainder of 59.6 can't round
' up and produce "1m 60s".
Private Function FormatElapsed(secs As Double) As String
    If secs < 60 Then
        FormatElapsed = Format(secs, "0.0") & "s"
        Exit Function
    End If

    Dim t As Long, h As Long, m As Long, s As Long
    t = CLng(Int(secs + 0.5))
    h = t \ 3600
    m = (t - h * 3600) \ 60
    s = t - h * 3600 - m * 60

    If h = 0 Then
        FormatElapsed = m & "m " & Format(s, "00") & "s"
    Else
        FormatElapsed = h & "h " & Format(m, "00") & "m " & Format(s, "00") & "s"
    End If
End Function

' Where a given output type belongs, given the model's own folder.
' With optUseSubfolders on that's <baseDir>\Export\STEP\ (or PNG, PDF),
' created on demand. Returns a path with a trailing backslash. If a folder
' can't be created - read-only drive, a file already using the name - the
' model's own folder is returned so the export still happens.
Private Function OutFolder(baseDir As String, kind As String) As String
    OutFolder = baseDir
    If Not optUseSubfolders Then Exit Function
    If baseDir = "" Then Exit Function

    Dim root As String, leaf As String
    root = baseDir
    If Right(root, 1) <> "\" Then root = root & "\"
    root = root & optExportFolderName
    leaf = root & "\" & UCase(kind)

    On Error Resume Next

    If Dir(root, vbDirectory) = "" Then
        Err.Clear
        MkDir root
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    End If
    Err.Clear

    If Dir(leaf, vbDirectory) = "" Then
        Err.Clear
        MkDir leaf
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    End If
    Err.Clear

    On Error GoTo 0
    OutFolder = leaf & "\"
End Function

' Drops a trailing file extension, if there is one.
Private Function StripExtension(s As String) As String
    Dim dotPos As Long, slashPos As Long
    dotPos = InStrRev(s, ".")
    slashPos = InStrRev(s, "\")
    If dotPos > slashPos And dotPos > 0 Then
        StripExtension = Left(s, dotPos - 1)
    Else
        StripExtension = s
    End If
End Function

Private Function SafeFileName(s As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", "^")
    For Each ch In bad
        s = Replace(s, ch, "_")
    Next ch
    SafeFileName = Trim(s)
End Function

' Avoids two exports in the same folder landing on the same file this run.
' extList is the comma-separated extensions this name will be used with,
' e.g. "step,png" or "pdf". Names are reserved per extension, so a part's
' STEP and PNG don't push its drawing's PDF onto a "_2" name.
Private Function UniqueName(folder As String, s As String, extList As String) As String
    If extList = "" Then
        UniqueName = s
        Exit Function
    End If

    Dim exts As Variant
    exts = Split(extList, ",")

    Dim candidate As String, n As Long, i As Long, clash As Boolean
    candidate = s
    n = 2
    Do
        clash = False
        For i = 0 To UBound(exts)
            If usedNames.Exists(LCase(folder & candidate & "." & Trim(CStr(exts(i))))) Then clash = True
        Next i
        If Not clash Then Exit Do
        candidate = s & "_" & n
        n = n + 1
    Loop

    For i = 0 To UBound(exts)
        usedNames.Add LCase(folder & candidate & "." & Trim(CStr(exts(i)))), True
    Next i

    UniqueName = candidate
End Function

' ---- CSV logging ----------------------------------------------------
' One row per output file. Columns:
'   timestamp      when the row was written
'   source_type    part or assembly
'   source_name    file name of the source, with extension, no path - the
'                  model for step and png rows, the drawing for pdf rows
'   configuration  the configuration the assembly uses
'   source_path    folder holding that source file, no trailing slash
'   output_type    step, png, pdf or csv (a BOM table)
'   output_path    the output file, relative to source_path
'   status         ok, failed, skipped or error
'   message        detail, mostly for failures
' Callers pass full paths; the splitting happens here.
Private Sub AddRow(ByVal srcType As String, ByVal srcTitle As String, ByVal cfgName As String, _
                   ByVal srcFullPath As String, ByVal outType As String, ByVal outFullPath As String, _
                   ByVal status As String, ByVal message As String)
    Dim srcName As String, srcDir As String, outRel As String

    If srcFullPath <> "" Then
        srcName = Mid(srcFullPath, InStrRev(srcFullPath, "\") + 1)
        srcDir = Left(srcFullPath, InStrRev(srcFullPath, "\") - 1)
    Else
        srcName = srcTitle          ' virtual parts have no file of their own
        srcDir = ""
    End If

    outRel = RelativeTo(srcDir, outFullPath)

    rows.Add Array(Format(Now, "yyyy-mm-dd hh:nn:ss"), srcType, srcName, cfgName, srcDir, _
                   outType, outRel, status, message)
    Debug.Print status & " | " & srcName & " | " & outType & " | " & outRel & " | " & message
End Sub

' Expresses fullPath relative to baseDir. Same folder gives just the file
' name, a nearby folder gives a ..\ path, and anything else stays absolute.
Private Function RelativeTo(baseDir As String, fullPath As String) As String
    If fullPath = "" Then Exit Function
    If baseDir = "" Then
        RelativeTo = fullPath
        Exit Function
    End If

    Dim b As String, f As String
    b = baseDir
    If Right(b, 1) = "\" Then b = Left(b, Len(b) - 1)
    f = fullPath

    ' Same folder is the common case.
    If LCase(Left(f, Len(b) + 1)) = LCase(b & "\") Then
        Dim rest As String
        rest = Mid(f, Len(b) + 2)
        If InStr(rest, "\") = 0 Then
            RelativeTo = rest
            Exit Function
        End If
    End If

    Dim bParts As Variant, fParts As Variant
    bParts = Split(b, "\")
    fParts = Split(f, "\")

    ' Different drives cannot be expressed relatively.
    If LCase(bParts(0)) <> LCase(fParts(0)) Then
        RelativeTo = fullPath
        Exit Function
    End If

    Dim common As Long, i As Long
    common = 0
    Do While common <= UBound(bParts) And common < UBound(fParts)
        If LCase(bParts(common)) <> LCase(fParts(common)) Then Exit Do
        common = common + 1
    Loop

    Dim out As String
    For i = common To UBound(bParts)
        out = out & "..\"
    Next i
    For i = common To UBound(fParts)
        out = out & fParts(i)
        If i < UBound(fParts) Then out = out & "\"
    Next i

    RelativeTo = out
End Function

Private Sub WriteCsv(path As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f

    Print #f, "timestamp,source_type,source_name,configuration,source_path," & _
              "output_type,output_path,status,message"

    Dim i As Long, j As Long
    Dim r As Variant
    Dim line As String

    For i = 1 To rows.Count
        r = rows(i)
        line = ""
        For j = LBound(r) To UBound(r)
            If j > LBound(r) Then line = line & ","
            line = line & CsvField(CStr(r(j)))
        Next j
        Print #f, line
    Next i

    Close #f
End Sub

' Quotes a field if it contains a comma, quote or line break.
Private Function CsvField(s As String) As String
    Dim needsQuotes As Boolean
    needsQuotes = (InStr(s, ",") > 0) Or (InStr(s, """") > 0) Or _
                  (InStr(s, vbCr) > 0) Or (InStr(s, vbLf) > 0)

    s = Replace(s, """", """""")
    If needsQuotes Then
        CsvField = """" & s & """"
    Else
        CsvField = s
    End If
End Function
