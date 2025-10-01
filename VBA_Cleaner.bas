Option Explicit

'==========================================================================
'   This module contains two tools, "VBA-Cleaner" and "VBA-DeepClean" (Excel, 64-bit safe):
'   These are simple 64-bit safe alternatives to tools such as the 32-bit "VBA CodeCleaner" written by Rob Bovery (http://www.appspro.com/Utilities/CodeCleaner.htm), and the commercial tool Ribbon Commander (https://www.spreadsheet1.com/vba-project-code-cleaner-for-access-excel-powerpoint-word.html).
'   This code is provided as is, with no warranty of any kind.
'   Author: Peter.Schild@OsloMet.no
'
'   1) VBA-Cleaner
'   - Use VBA-Cleaner (export/remove/import + clear/add for document modules) often during development. It’s fast and fixes VBA-side issues.
'   - Exports all components
'   - Removes/rebuilds non-document modules
'   - Refreshes document modules (ThisWorkbook/Sheets) in-place
'   - Works only on open workbooks (so that you can unlock password first, if necessary)
'   - Uses late binding. No need to add Tools > References
'   - It's recommended to take a backup of the file before running VBA-Cleaner on it
'   - Preserves VBA-project password-protecion
'   - Can clean both .xlsm and .xlsb files
'
'   2) VBA-DeepClean
'   - Use VBA-DeepClean when file size stays huge after VBA-Cleaner, compilation glitches persist, or you inherit a workbook with years of cruft (names/styles/UsedRange issues)
'   - Exports all code (as in VBA-Cleaner).
'   - Creates a new empty .xlsm.
'   - DeepClean gives every sheet a fresh document module and a fresh workbook container
'   - Copies each worksheet from old > new (ws.Copy), which creates fresh document module streams for those sheets.
'   - Imports all Std/Class modules and Forms into the new workbook.
'   - For ThisWorkbook (and any sheet modules if you prefer), clear code and AddFromFile to ensure a fresh code stream.
'   - Saves the new workbook with suffix _DeepCleaned. Does not change the original file.
'   - VBA-project password-protecion (if any) is removed
'   - DeepClean outputs only .xlsm files, which seem the most reliable. I found a bug in .xlsb files that difficult to reconnect OnAct macros (buttons etc)
'
'   Instructions (for both):
'   - Before running: In Excel, enable File > Options > Trust Center > Trust Center Settings… > Macro Settings > [] Trust access to the VBA project object model.
'   - Open the file you wish to clean, and unprotect the code it its VBA project is password-protected
'   - key ALT+F8 to run subroutine "VBA_Cleaner" or "VBA_DeepClean"
'==========================================================================

Private Const vbext_ct_StdModule& = 1&
Private Const vbext_ct_ClassModule& = 2&
Private Const vbext_ct_MSForm& = 3&
Private Const vbext_ct_ActiveXDesigner& = 11&
Private Const vbext_ct_Document& = 100&
Private Const vbext_pp_locked& = 1&
Private Const TEMP_FOLDER_NAME$ = "VbaProjectCleanerTemp"

Public Sub VBA_Cleaner() '<= RUN THIS
    Dim vbeObj As Object 'Application.VBE
    Dim proj As Object   'VBProject (late-bound)

    If Not IsTrustAccessEnabled() Then
        MsgBox "Please enable: File > Options > Trust Center > Trust Center Settings… > Macro Settings > Trust access to the VBA project object model", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    Set vbeObj = Application.VBE
    If vbeObj Is Nothing Then
        MsgBox "VBE not available.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Set proj = PickOpenVBProject(vbeObj)
    If proj Is Nothing Then Exit Sub
    
    If proj.Protection = vbext_pp_locked Then
        MsgBox "Project '" & proj.Name & "' is locked. Please unlock and try again.", vbExclamation
        Exit Sub
    End If

    Clean_VBA_project proj
End Sub

Private Function PickOpenVBProject(ByVal vbeObj As Object) As Object
    'Tiny project picker (open projects only)
    'Called by VBA_Cleaner
    Dim i&, mini&, maxi&, badi&
    Dim info$
    Dim proj As Object
    Dim choice As Variant
    Dim hostName$
    Dim wb As Workbook
    
    info = "Open VBA projects:" & vbCrLf & vbCrLf
    badi = 0
    maxi = 0
    mini = vbeObj.VBProjects.Count
    For i = 1 To vbeObj.VBProjects.Count
        Set proj = vbeObj.VBProjects(i)
        For Each wb In Application.Workbooks
            If Not (wb Is Nothing) Then
                If wb.VBProject Is proj Then
                    hostName = wb.Name
                    Exit For
                End If
            End If
        Next wb
        If hostName = ThisWorkbook.Name Then 'avoid VBA-Clean workbook cleaning itself - it would crash
            badi = i
        Else
            If i < mini Then mini = i
            If maxi < i Then maxi = i
            info = info & i & ") " & proj.Name
            If 0 < LenB(hostName) Then info = info & "  —  " & hostName
            info = info & vbCrLf
        End If
    Next i
    info = info & vbCrLf & "Which # to clean?  (Enter a number " & mini & "-" & maxi & "):"

    choice = Application.InputBox(Prompt:=info, Title:="VBA Cleaner — Project Picker", Type:=1)
    If choice = False Then Exit Function 'cancel
    
    On Error Resume Next
    If Len(choice & vbNullString) > 0 Then
        If mini <= choice And choice <= maxi And choice <> badi Then Set PickOpenVBProject = vbeObj.VBProjects(choice)
    End If
    On Error GoTo 0
    If PickOpenVBProject Is Nothing Then MsgBox "Invalid selection. VBA_Clean has aborted.", vbExclamation
End Function

Private Sub Clean_VBA_project(ByVal Project As Object)
    'This is the core code for VBA-Cleaner, with late-binding
    'Called by VBA_Cleaner
    Dim i&
    Dim totalLines&
    Dim f$
    Dim base$
    Dim codeText$
    Dim ext$
    Dim fileExt$
    Dim tempPath$
    Dim wasEvents As Boolean
    Dim wasAlerts As Boolean
    Dim exportMap As Object 'Scripting.Dictionary
    Dim comp As Object
    Dim codeMod As Object
    Dim appWasUpdating As Boolean

    appWasUpdating = Application.ScreenUpdating
    wasEvents = Application.EnableEvents
    wasAlerts = Application.DisplayAlerts
    
    On Error GoTo quitSub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    tempPath = Environ$("TEMP")
    If Len(tempPath) = 0 Then tempPath = CurDir$
    tempPath = tempPath & "\" & TEMP_FOLDER_NAME
    
    EnsureEmptyFolder tempPath
    Set exportMap = CreateObject("Scripting.Dictionary") 'docName -> exported path

    '1) Export every component
    For Each comp In Project.VBComponents
        fileExt = ExportExtensionFor(comp)
        f = tempPath & "\" & SafeFileName(comp.Name) & fileExt
        comp.Export f
        If comp.Type = vbext_ct_Document Then exportMap(comp.Name) = f
        Debug.Print "Export", comp.Name & fileExt
    Next comp

    '2) Remove all NON-document components
    For i = Project.VBComponents.Count To 1 Step -1
        Set comp = Project.VBComponents(i)
        If comp.Type <> vbext_ct_Document Then
            Project.VBComponents.Remove comp
            Debug.Print "Remove", i
        End If
    Next i

    '3) Re-import non-document files (.bas/.cls/.frm/.pag)
    f = Dir$(tempPath & "\*.*")
    Do While LenB(f) <> 0
        base = Left$(f, InStrRev(f, ".") - 1)
        ext = LCase$(Mid$(f, InStrRev(f, ".")))
        If ext = ".bas" Or ext = ".cls" Or ext = ".frm" Or ext = ".pag" Then
            If Not exportMap.Exists(base) Then
                Project.VBComponents.Import tempPath & "\" & f
                Debug.Print "Re-import", f
            End If
        End If
        f = Dir$
    Loop

    '4) Refresh document modules in-place (clear + AddFromFile)
    For Each comp In Project.VBComponents
        If comp.Type = vbext_ct_Document Then
            If exportMap.Exists(comp.Name) Then
                Set codeMod = comp.CodeModule
                totalLines = codeMod.CountOfLines
                If totalLines > 0 Then codeMod.DeleteLines 1, totalLines
                codeText = LoadCodeBodyFromExport(exportMap(comp.Name))
                If codeText <> vbNullString Then codeMod.AddFromString codeText
                Debug.Print "Refresh", comp.Name
            End If
        End If
    Next comp

    MsgBox "VBA Project '" & Project.Name & "' successfully cleaned.", vbInformation
    Debug.Print "...Successfully finished"
quitSub:
    Application.DisplayAlerts = wasAlerts
    Application.EnableEvents = wasEvents
    Application.ScreenUpdating = appWasUpdating
    If Err.Number <> 0 Then
        MsgBox "Error while cleaning '" & Project.Name & "': " & Err.Number & " - " & Err.Description, vbCritical
        'We don't delete the temp files on failure. They might need salvaging.
        Debug.Print "Temp files are in " & tempPath
    Else
        On Error Resume Next
        EnsureFolderDeleted tempPath 'clear temp files
        On Error GoTo 0
    End If
End Sub

'=====================================
'   VBA-DeepClean
'=====================================

Public Sub VBA_DeepClean() '<= RUN THIS
    Dim wb As Workbook

    If Not IsTrustAccessEnabled() Then
        MsgBox "Please enable: File > Options > Trust Center > Trust Center Settings… > Macro Settings > Trust access to the VBA project object model", vbCritical
        Exit Sub
    End If
    
    Set wb = PickWorkbookForDeepClean()
    If wb Is Nothing Then Exit Sub
    
    'Last chance check for lock
    On Error Resume Next
    If wb.VBProject.Protection = vbext_pp_locked Then
        MsgBox "The VBProject in '" & wb.Name & "' is locked. Please unlock it, then run again.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    DeepClean_Workbook wb
End Sub

Private Function PickWorkbookForDeepClean() As Workbook
    'Tiny picker: choose an open workbook
    '- Shows index, workbook name, project name, and lock state
    '- Blank/Cancel safely exits
    '- Blank input = ActiveWorkbook
    Dim i&, mini&, maxi&, badi&
    Dim msg$
    Dim choice As Variant
    Dim wb As Workbook
    Dim projName$
    Dim lockTxt$

    msg = "Open workbooks:" & vbCrLf & vbCrLf
    badi = 0
    maxi = 0
    mini = Application.Workbooks.Count
    For i = 1 To Application.Workbooks.Count
        Set wb = Application.Workbooks(i)
        If wb Is ThisWorkbook Then 'avoid VBA-Cleaner workbook cleaning itself - it would crash
            badi = i
        Else
            If i < mini Then mini = i
            If maxi < i Then maxi = i
            projName = SafeProjName(wb)
            lockTxt = ProjLockText(wb)
            msg = msg & i & ") " & wb.Name & vbCrLf
        End If
    Next i

    If maxi = 0 Then
        MsgBox "No workbooks are open.", vbInformation
        Exit Function
    End If

    msg = msg & vbCrLf & "Which # to deep-clean?  (Enter a number " & mini & "-" & maxi & "):"
    choice = Application.InputBox(Prompt:=msg, Title:="DeepClean — Workbook Picker", Type:=1)
    If choice = False Then Exit Function 'Cancel

    On Error Resume Next
    If 0 < Len(choice & vbNullString) Then
        If mini <= choice And choice <= maxi And choice <> badi Then Set PickWorkbookForDeepClean = Application.Workbooks(CLng(choice))
    End If
    On Error GoTo 0
    If PickWorkbookForDeepClean Is Nothing Then MsgBox "Invalid selection. VBA_DeepClean has aborted", vbInformation
End Function

Private Function SafeProjName$(ByVal wb As Workbook)
    'Best-effort VBProject name without requiring VBIDE reference
    'Helper function to PickWorkbookForDeepClean
    On Error Resume Next
    SafeProjName = wb.VBProject.Name
    If LenB(SafeProjName) = 0 Then SafeProjName = "(unknown)"
    On Error GoTo 0
End Function

Private Function ProjLockText$(ByVal wb As Workbook)
    'Lock status text (helps you unlock before running)
    'Helper function to PickWorkbookForDeepClean
    On Error Resume Next
    If wb.VBProject.Protection = vbext_pp_locked Then ProjLockText = "   [LOCKED]" Else ProjLockText = vbNullString
    On Error GoTo 0
End Function

Private Sub DeepClean_Workbook(Optional ByVal SrcWb As Workbook)
    'This is the core code for VBA-DeepClean, with late-binding
    Dim i&
    Dim k&
    Dim base$
    Dim docText$
    Dim destCodeName$
    Dim ext$ 'file extension
    Dim f$ 'filename
    Dim projName$ 'VBProject name to restore
    Dim refs$() 'VBA project references
    Dim tempPath$
    Dim twText$
    Dim flag As Boolean
    Dim firstDone As Boolean
    Dim tabk As Variant
    Dim shp As Shape 'e.g. macro button
    Dim ch As Chart
    Dim ws As Worksheet 'temp worksheet
    Dim src As Workbook 'source workbook
    Dim dst As Workbook 'destination workbook (cleaned)
    Dim wbk As Workbook 'temp workbook
    Dim projSrc As Object 'source VBAProject
    Dim projDst As Object 'destination VBAProject
    Dim comp As Object
    Dim codeMod As Object
    Dim tabToFile As Object
    Dim tabToKind As Object

    On Error GoTo quitSub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    'Mapping from TAB NAME -> (exported .cls file path, kind "WS"/"CH")
    Set tabToFile = CreateObject("Scripting.Dictionary")
    Set tabToKind = CreateObject("Scripting.Dictionary")

    '1) Remember references, so that they can be restored later
    If SrcWb Is Nothing Then Set src = ThisWorkbook Else Set src = SrcWb
    Set projSrc = src.VBProject
    
    projName = projSrc.Name
    ReDim refs(projSrc.References.Count, 3)
    For Each comp In projSrc.References
        i = i + 1
        refs(i, 1) = CStr(comp.GUID)
        refs(i, 2) = CStr(comp.Major)
        refs(i, 3) = CStr(comp.Minor)
    Next

    '2) Temp folder + export all components
    tempPath = Environ$("TEMP"): If Len(tempPath) = 0 Then tempPath = CurDir$
    tempPath = tempPath & "\VbaDeepCleanTemp"
    EnsureEmptyFolder tempPath

    For Each comp In projSrc.VBComponents
        comp.Export tempPath & "\" & SafeFileName(comp.Name) & ExportExtensionFor(comp)
    Next

    'Build TAB NAME -> source exported file mapping (document modules)
    '(Locale-agnostic: use actual sheet/chartsheet names)
    For Each ws In src.Worksheets
        tabToFile(ws.Name) = tempPath & "\" & ws.CodeName & ".cls"
        tabToKind(ws.Name) = "WS"
    Next
    For Each ch In src.Charts
        tabToFile(ch.Name) = tempPath & "\" & ch.CodeName & ".cls"
        tabToKind(ch.Name) = "CH"
    Next

    '3) Create destination by copying the FIRST sheet to a NEW workbook (no placeholder)
    firstDone = False
    For Each ws In src.Worksheets
        ws.Copy 'no destination -> creates a new workbook with this sheet as the only sheet
        Set dst = ActiveWorkbook
        firstDone = True
        Exit For
    Next
    'If there were no worksheets, start with a blank and we’ll add charts only
    If Not firstDone Then Set dst = Application.Workbooks.Add(xlWBATWorksheet)

    'Copy remaining worksheets (preserves tab names)
    i = 0
    For Each ws In src.Worksheets
        i = i + 1
        If i > 1 Then ws.Copy After:=dst.Sheets(dst.Sheets.Count)
    Next
    'Copy chartsheets too
    For Each ch In src.Charts
        ch.Copy After:=dst.Sheets(dst.Sheets.Count)
    Next

    '4) Rebuild VBA in destination (late-bound)
    Set projDst = dst.VBProject
    If projDst.Protection = vbext_pp_locked Then Err.Raise vbObjectError + 1, , "Destination project is locked."
    projDst.Name = projName 'reinstate project name

    'Reinstate references
    For i = 1 To UBound(refs, 1)
        On Error Resume Next
        projDst.References.AddFromGuid refs(i, 1), CLng(refs(i, 2)), CLng(refs(i, 3))
        On Error GoTo quitSub
    Next

    'Import all NON-document components
    f = Dir$(tempPath & "\*.*")
    Do While LenB(f) <> 0
        ext = LCase$(Mid$(f, InStrRev(f, ".")))
        base = Left$(f, InStrRev(f, ".") - 1)
        If ext = ".bas" Or ext = ".cls" Or ext = ".frm" Or ext = ".pag" Then
            'Skip any .cls that corresponds to document modules (their .cls will be injected as plain code)
            If Not IsDocumentFileForAnyTab(base, tabToFile) And base <> "ThisWorkbook" Then projDst.VBComponents.Import tempPath & "\" & f
        End If
        f = Dir$
    Loop

    'Refresh ThisWorkbook code (strip headers)
    Set comp = Nothing
    On Error Resume Next
    Set comp = projDst.VBComponents("ThisWorkbook")
    On Error GoTo quitSub
    If Not comp Is Nothing Then
        Set codeMod = comp.CodeModule
        If codeMod.CountOfLines > 0 Then codeMod.DeleteLines 1, codeMod.CountOfLines
        f = tempPath & "\ThisWorkbook.cls"
        If Dir$(f) <> vbNullString Then
            twText = LoadCodeBodyFromExport(f)
            If LenB(twText) > 0 Then codeMod.AddFromString twText
        End If
    End If

    'Refresh each document module by TAB NAME (locale-safe; ignores CodeName renumbering)
    For Each tabk In tabToFile.Keys
        f = CStr(tabToFile(tabk))
        If Dir$(f) <> vbNullString Then
            'Find destination component by tab name (worksheet or chart)
            destCodeName = DestCodeNameByTab(dst, CStr(tabk), CStr(tabToKind(tabk)))
            If LenB(destCodeName) > 0 Then
                Set comp = Nothing
                On Error Resume Next
                Set comp = projDst.VBComponents(destCodeName)
                On Error GoTo quitSub
                If Not comp Is Nothing Then
                    Set codeMod = comp.CodeModule
                    If codeMod.CountOfLines > 0 Then codeMod.DeleteLines 1, codeMod.CountOfLines
                    docText = LoadCodeBodyFromExport(f)
                    If LenB(docText) > 0 Then codeMod.AddFromString docText
                End If
            End If
        End If
    Next

    'Reconnect all buttons to macros local file instead of link to original file. Finds form controls, shapes, and charts with OnAction
    For Each ws In dst.Worksheets
        For Each shp In ws.Shapes
            RewireShapeOnAction shp, dst.Name
        Next shp
    Next ws
    For Each ch In dst.Charts
        For Each shp In ch.Shapes
            RewireShapeOnAction shp, dst.Name
        Next shp
    Next ch

    '5) Save, tidy, close
    i = InStrRev(src.Name, ".")
    f = Left(src.Name, i - 1) & "_DeepCleaned.xlsm"
    Do
        tabk = Application.GetSaveAsFilename(f, "Macro-Enabled Workbook (*.xlsm),*.xlsm", , "Choose where to save the DeepCleaned workbook") 'Save as macro-enabled
        If VarType(tabk) = vbBoolean And tabk = False Then GoTo quitSub
        f = CStr(tabk)
        flag = IsLikelyUrl(f)
        If flag Then MsgBox "Cannot save on online folders. Please try again, and select a local file path.", vbExclamation
    Loop Until Not flag
    'If the target is already open in this Excel instance, close it (no save) so that it can be overwritten
    For Each wbk In Application.Workbooks
        If StrComp(wbk.FullName, f, vbTextCompare) = 0 Then
            wbk.Close SaveChanges:=False
            Exit For
        End If
    Next
    'If a file already exists at path, remove read-only flag and delete it
    If Dir$(f, vbNormal) <> vbNullString Then
        On Error Resume Next
        SetAttr f, vbNormal  'clear read-only if set
        Kill f               'delete existing file
        On Error GoTo quitSub
    End If
    dst.CheckCompatibility = False
    Debug.Print "Saving to file: " & f
    dst.SaveAs Filename:=f, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    'Release any refs into the destination VBProject BEFORE closing
    Set codeMod = Nothing
    Set comp = Nothing
    Set projDst = Nothing
    'We close the new file here because for user to check macro functionality, they have to open it, which triggers events Auto_Open and Workbook_Open
    dst.Saved = True 'Ensure that Excel believes it's saved
    dst.Close SaveChanges:=False 'Close without re-saving (we already saved)

    MsgBox "DeepClean successfully completed:" & vbLf & f, vbInformation
    Debug.Print "...DeepClean finished"
quitSub:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Deep Clean failed: " & Err.Number & " - " & Err.Description, vbCritical
        'We don't delete the temp files on failure. They might need salvaging.
        Debug.Print "Temp files are in " & tempPath
    Else
        On Error Resume Next
        EnsureFolderDeleted tempPath
        On Error GoTo 0
    End If
End Sub

Private Function IsLikelyUrl(ByVal p$) As Boolean
    'Helper function to DeepClean_Workbook. To check whether Excel is trying to save to Sharepoint/OneDrive or mirrored folder on local hard disk
    p = LCase$(Trim$(p))
    IsLikelyUrl = (Left$(p, 7) = "http://" Or Left$(p, 8) = "https://")
End Function

Private Function IsDocumentFileForAnyTab(ByVal baseName$, ByVal tabToFile As Object) As Boolean
    'Returns True if the given base filename (without extension) matches any exported document .cls
    'Helper function to DeepClean_Workbook
    Dim tabk As Variant
    Dim p$

    For Each tabk In tabToFile.Keys
        p = CStr(tabToFile(tabk))
        If LCase$(baseName) = LCase$(Left$(Mid$(p, InStrRev(p, "\") + 1), InStrRev(p, ".") - InStrRev(p, "\") - 1)) Then
            IsDocumentFileForAnyTab = True
            Exit Function
        End If
    Next
End Function

Private Function DestCodeNameByTab$(ByVal wb As Workbook, ByVal tabName$, ByVal kind$)
    'Find the destination component CodeName by tab name (and kind). This is more robust because Worksheet tab name is preserved on copy
    'Helper function to DeepClean_Workbook
    On Error Resume Next
    If kind = "WS" Then
        Dim ws As Worksheet
        Set ws = wb.Worksheets(tabName)
        If Not ws Is Nothing Then DestCodeNameByTab = ws.CodeName
    Else
        Dim ch As Chart
        Set ch = wb.Charts(tabName)
        If Not ch Is Nothing Then DestCodeNameByTab = ch.CodeName
    End If
    On Error GoTo 0
End Function

Private Sub RewireShapeOnAction(ByVal shp As Shape, ByVal targetWbName$)
    'Reconnects buttons by renaming links "'Book Name.xlsm'!Module1.Macro1" or "Book.xlsm!Macro1" to "'Book Name_DeepClean.xlsm'!Module1.Macro1" or "Book_DeepClean.xlsm!Macro1".
    Dim p& 'pointer
    Dim act$ 'OnAction
    Dim tail$ 'Macro link tail: Module.Proc or Proc name

    On Error Resume Next
    act = Trim(shp.OnAction)
    If LenB(act) = 0 Then Exit Sub
        
    p = InStrRev(act, "!")
    If 0 < p Then tail = Mid$(act, p + 1) Else tail = act
    If Left$(tail, 1) = "'" And Right$(tail, 1) = "'" Then tail = Mid$(tail, 2, Len(tail) - 2) 'Remove wrapping single quotes (Excel sometimes quotes the whole thing)
    tail = Trim$(tail)
    If LenB(tail) = 0 Then Exit Sub

    shp.OnAction = "'" & targetWbName & "'!" & tail 'Always qualify with destination workbook name (safe solution)
    Debug.Print "#:", targetWbName 'fixo
    On Error GoTo 0
End Sub

'=========================================================
'   Helper code used by both VBA-Cleaner and VBA-DeepClean
'=========================================================

Private Function IsTrustAccessEnabled() As Boolean
    'Helper function to VBA_Cleaner and VBA_DeepClean
    Dim n&
    On Error Resume Next
    n = Application.VBE.VBProjects.Count
    IsTrustAccessEnabled = (Err.Number = 0)
    Err.Clear
End Function

Private Function SafeFileName$(ByVal s$)
    'Helper function to Clean_VBA_project and DeepClean_Workbook
    Dim bad As Variant
    Dim ch As Variant
    bad = Array("<", ">", ":", """", "/", "\", "|", "?", "*")
    For Each ch In bad
        s = Replace$(s, CStr(ch), "_")
    Next ch
    SafeFileName = s
End Function

Private Function ExportExtensionFor$(ByVal comp As Object)
    'Helper function to Clean_VBA_project and DeepClean_Workbook
    Select Case comp.Type
        Case vbext_ct_StdModule:        ExportExtensionFor = ".bas"
        Case vbext_ct_ClassModule:      ExportExtensionFor = ".cls"
        Case vbext_ct_MSForm:           ExportExtensionFor = ".frm"
        Case vbext_ct_ActiveXDesigner:  ExportExtensionFor = ".pag"
        Case vbext_ct_Document:         ExportExtensionFor = ".cls" 'document modules export as .cls
        Case Else:                      ExportExtensionFor = ".txt"
    End Select
End Function

Private Sub EnsureEmptyFolder(ByVal path$)
    'Called by Clean_VBA_project and DeepClean_Workbook
    On Error Resume Next
    If Len(Dir$(path, vbDirectory)) = 0 Then
        MkDir path
    Else
        Kill path & "\*.*"
        Kill path & "\*.frx"
    End If
    On Error GoTo 0
End Sub

Private Sub EnsureFolderDeleted(ByVal path$)
    'Called by Clean_VBA_project and DeepClean_Workbook
    On Error Resume Next
    Kill path & "\*.*"
    Kill path & "\*.frx"
    RmDir path 'remove directory
    On Error GoTo 0
End Sub

Private Function LoadCodeBodyFromExport$(ByVal filePath$)
    'Called by Clean_VBA_project and DeepClean_Workbook
    Dim f%
    Dim i&
    Dim line$
    Dim buf$
    Dim inHeaderBlock As Boolean
    
    f = FreeFile
    On Error GoTo CleanFail
    Open filePath For Input As #f
    Do While Not EOF(f)
        i = i + 1
        Line Input #f, line
        If i = 1 And LCase(line) Like "version*" Then 'Start of class header, e.g VERSION 1.0 CLASS
            inHeaderBlock = True
        ElseIf inHeaderBlock Then 'Ignore everything inside the header block: skip until it closes with a standalone "END"
            If LCase(line) = "end" Then inHeaderBlock = False
        ElseIf LCase$(line) Like "attribute*" Then
            'Skip Attribute lines (e.g., Attribute VB_Name = ..., VB_PredeclaredId, etc.)
        Else
            buf = buf & line & vbCrLf 'Keep everything else (including Option Explicit, code, comments)
        End If
    Loop
    Close #f
    LoadCodeBodyFromExport = buf
    Exit Function

CleanFail:
    On Error Resume Next
    If f <> 0 Then Close #f
    LoadCodeBodyFromExport = vbNullString
End Function

Private Sub reboot() 'Not called by any code. Can be run manually enable events/screen in case any VBA code crashed
    With Application
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End With
End Sub
