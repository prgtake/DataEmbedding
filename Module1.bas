Attribute VB_Name = "Module1"
' =====================================================
'  データ埋込マクロ (Data Embedding Macro)
'  Copyright (c) 2026 Datan (データン)
'  Licensed under the MIT License.
' =====================================================
Option Explicit

' --- Version Management ---
Public Const APP_VERSION As String = "1.1.0"

' --- Startup ---
Sub auto_open()
    Dim thisBookName As String
    thisBookName = ActiveWorkbook.Name

    Dim b As Integer, e As Integer
    If Workbooks.Count > 1 Then
        e = 0
        For b = 1 To Workbooks.Count
            If Workbooks(b - e).Name <> thisBookName Then
                Workbooks(b - e).Close SaveChanges:=True
                e = e + 1
            End If
        Next b
    End If

    UserForm1.Show
End Sub

' --- Helper to get Local Path (Handles OneDrive/SharePoint URLs) ---
Private Function GetLocalPath(ByVal folderPath As String) As String
    If Left(folderPath, 8) <> "https://" Then
        GetLocalPath = folderPath
        Exit Function
    End If

    Dim oneDrivePath As String
    Dim relativePath As String
    Dim slashPos As Long

    ' Get local OneDrive root from environment variables
    ' Handles both Personal and Business OneDrive
    oneDrivePath = Environ("OneDrive")
    If oneDrivePath = "" Then oneDrivePath = Environ("OneDriveCommercial")

    ' "https://d.docs.live.net/xxxxxxxx/" - Cut the ID part
    ' Find the 3rd slash after the domain
    slashPos = InStr(9, folderPath, "/") ' Search after https://
    slashPos = InStr(slashPos + 1, folderPath, "/") ' Slash after the ID

    If slashPos > 0 Then
        ' Extract path after the slash (e.g., /Desktop/MacroDir)
        relativePath = Mid(folderPath, slashPos)
        ' Convert slashes to backslashes and join
        GetLocalPath = oneDrivePath & Replace(relativePath, "/", "\")
    Else
        GetLocalPath = folderPath ' Return as is if conversion fails
    End If
End Function

' --- Helper to Log Messages ---
Private Sub LogMessage(ByVal msg As String, ByVal outputDir As String)
    On Error Resume Next
    Dim fso As Object
    Dim ts As Object
    Dim logFile As String
    logFile = outputDir & "\process_log.txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Open for appending (8), create if not exists (True)
    Set ts = fso.OpenTextFile(logFile, 8, True)
    ts.WriteLine Format(Now, "yyyy/mm/dd hh:mm:ss") & " : " & msg
    ts.Close
    On Error GoTo 0
End Sub

' --- Main Logic for Data Embedding (Injection Method) ---
Public Sub ProcessDataEmbedding(ByVal targetPath As String, _
                               ByVal templateName As String, _
                               ByVal dataName As String, _
                               ByVal startRow As Long, _
                               ByVal sheetsPerBook As Long)

    Dim targetWb As Workbook
    Dim dataWs As Worksheet
    Dim templateWs As Worksheet
    Dim newFileName As String
    Dim ry As Long
    Dim thisBookName As String
    Dim b As Integer, c As Integer
    Dim outputDir As String

    Dim timestampStr As String
    Dim fileIndex As Integer
    Dim outWb As Workbook
    Dim sheetCount As Integer
    Dim currentRow As Long
    Dim defaultDataRow As Long

    Dim backupPath As String
    Dim fso As Object

    ' --- OneDrive Process Kill (Match BindM.bas) ---
    On Error Resume Next
    Shell "taskkill /F /IM OneDrive.exe", vbHide
    On Error GoTo ErrorHandler

    ' Wait for OS to release locks (2 seconds)
    Application.Wait [Now() + "0:00:02"]
    ' -----------------------------------------------

    ' Set output directory relative to this workbook
    outputDir = GetLocalPath(ThisWorkbook.Path) & "\output_results"
    
    ' Initialization of logging
    On Error Resume Next
    MkDir outputDir
    On Error GoTo ErrorHandler
    LogMessage "=== Process Started ===", outputDir
    LogMessage "TargetPath: " & targetPath, outputDir

    backupPath = targetPath & ".bak"

    MsgBox "Starting data generation. Output folder: " & outputDir, vbInformation, "Start"

    thisBookName = ThisWorkbook.Name
    c = 0
    For b = 1 To Workbooks.Count
        If Workbooks(b - c).Name <> thisBookName Then
            Workbooks(b - c).Close SaveChanges:=False
            c = c + 1
        End If
    Next b

    Application.ScreenUpdating = False

    On Error Resume Next
    ChDir outputDir
    Kill outputDir & "\*.xlsx"
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrorHandler
    
    ' Create backup
    LogMessage "Creating backup...", outputDir
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(backupPath) Then fso.DeleteFile backupPath
    fso.CopyFile targetPath, backupPath

    ' Excel Settings
    With Application
        .Calculation = xlAutomatic
        .MaxChange = 0.001
        .SheetsInNewWorkbook = 1
    End With

    ' Open Master Workbook as ReadOnly to prevent AutoSave/OneDrive locks
    LogMessage "Opening master workbook as ReadOnly...", outputDir
    Application.EnableEvents = False
    Set targetWb = Workbooks.Open(Filename:=targetPath, ReadOnly:=True)
    
    ' (AutoSaveOn removed to ensure compatibility with older Excel versions)
    
    On Error Resume Next
    Set dataWs = targetWb.Worksheets(dataName)
    Set templateWs = targetWb.Worksheets(templateName)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrorHandler
    
    If dataWs Is Nothing Or templateWs Is Nothing Then
        LogMessage "Error: Specified sheets were not found.", outputDir
        MsgBox "Specified sheets were not found.", vbCritical, "Error"
        GoTo Cleanup
    End If

    ' Identify default data row
    defaultDataRow = startRow + 1

    ' --- Duplicate Check (Optimized with Dictionary) ---
    LogMessage "Running duplicate check...", outputDir
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    currentRow = defaultDataRow
    
    Do While dataWs.Cells(currentRow, 1).Value <> ""
        Dim keyVal As String
        keyVal = CStr(dataWs.Cells(currentRow, 1).Value)
        
        If dict.exists(keyVal) Then
            LogMessage "Error: Duplicate key found: " & keyVal, outputDir
            MsgBox "Duplicate keys found: " & keyVal & " (Row: " & currentRow & ")", vbCritical, "Error"
            GoTo Cleanup
        Else
            dict.Add keyVal, currentRow
        End If
        currentRow = currentRow + 1
    Loop
    Set dict = Nothing

    ' Initialization
    timestampStr = Format(Now, "yyyymmddhhmmss")
    fileIndex = 1
    currentRow = defaultDataRow

    ' --- Main Loop ---
    LogMessage "Starting main loop...", outputDir
    Do While dataWs.Cells(currentRow, 1).Value <> ""
        
        Dim tempPath As String
        Dim targetExt As String
        targetExt = fso.GetExtensionName(targetPath)
        
        ' Using a temporary file to preserve workbook properties (Subtraction Method)
        tempPath = outputDir & "\tmp_" & timestampStr & "_" & fileIndex & "." & targetExt
        
        LogMessage "Loop Index " & fileIndex & ": Creating temp copy...", outputDir
        ' 1. Create a copy of the target workbook to preserve all properties/styles
        targetWb.SaveCopyAs tempPath
        
        ' 2. Open the copy
        LogMessage "Opening temp copy: " & tempPath, outputDir
        Application.EnableEvents = False
        Set outWb = Workbooks.Open(tempPath)
        outWb.Activate
        
        ' (AutoSaveOn removed for compatibility)
        
        Dim outDataWs As Worksheet
        Dim outTemplateWs As Worksheet
        On Error Resume Next
        Set outDataWs = outWb.Worksheets(dataName)
        Set outTemplateWs = outWb.Worksheets(templateName)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo ErrorHandler
        
        ' Record original sheets to delete them later (Subtraction)
        Dim originalSheets As Object
        Set originalSheets = CreateObject("Scripting.Dictionary")
        Dim wsObj As Object
        For Each wsObj In outWb.Sheets
            ' Do NOT add the template sheet to the delete list, we will keep it as a result sheet
            If wsObj.Name <> templateName Then
                originalSheets.Add wsObj.Name, True
            End If
        Next wsObj
        
        sheetCount = 0
        ' Inner loop for sheets per book
        Do While sheetCount < sheetsPerBook
            newFileName = dataWs.Cells(currentRow, 1).Value
            If newFileName = "" Then Exit Do
            
            LogMessage "Processing record: " & newFileName, outputDir
            
            ' Injection Method: Directly assign values to row 2 of the OUTPUT data sheet
            ' This is more reliable than Copy/Paste and triggers formulas correctly
            Dim lastCol As Long
            lastCol = dataWs.Cells(1, dataWs.Columns.Count).End(xlToLeft).Column
            outDataWs.Range(outDataWs.Cells(2, 1), outDataWs.Cells(2, lastCol)).Value = _
                dataWs.Range(dataWs.Cells(currentRow, 1), dataWs.Cells(currentRow, lastCol)).Value
            
            ' Ensure formulas are fully updated across all sheets
            Application.CalculateFull
            DoEvents
            
            Dim workingWs As Worksheet
            ' If this is the last record for this book, use the original template to preserve links
            If sheetCount = sheetsPerBook - 1 Or dataWs.Cells(currentRow + 1, 1).Value = "" Then
                Set workingWs = outTemplateWs
                LogMessage "Using original template sheet to preserve links.", outputDir
            Else
                ' Otherwise, copy the template
                LogMessage "Copying template sheet...", outputDir
                outTemplateWs.Copy After:=outWb.Sheets(outWb.Sheets.Count)
                Set workingWs = outWb.Sheets(outWb.Sheets.Count)
            End If
            
            ' Set Sheet Name (Clean invalid characters)
            Dim cleanName As String
            cleanName = newFileName
            Dim invalidChars As Variant
            invalidChars = Array("\", "/", ":", "?", "*", "[", "]")
            Dim charIdx As Integer
            For charIdx = LBound(invalidChars) To UBound(invalidChars)
                cleanName = Replace(cleanName, invalidChars(charIdx), "")
            Next charIdx
            
            On Error Resume Next
            workingWs.Name = Left(cleanName, 31)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo ErrorHandler
            
            ' Paste Values to break formula links but keep the result
            LogMessage "Pasting values and freezing shapes for sheet: " & workingWs.Name, outputDir
            outWb.Activate
            workingWs.Activate
            
            ' Freeze Shapes/Textboxes
            Dim shp As Object
            For Each shp In workingWs.Shapes
                On Error Resume Next
                If shp.DrawingObject.Formula <> "" Then
                    shp.DrawingObject.Formula = ""
                End If
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo ErrorHandler
            Next shp
            
            workingWs.Cells.Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            workingWs.Cells(1, 1).Select
            
            currentRow = currentRow + 1
            sheetCount = sheetCount + 1
            
            ' End of data check
            If dataWs.Cells(currentRow, 1).Value = "" Then Exit Do
        Loop
        
        ' 3. Subtract: Delete original master sheets (except the template which was repurposed)
        LogMessage "Subtracting (deleting) original sheets...", outputDir
        Application.DisplayAlerts = False
        Dim k As Variant, i As Long
        k = originalSheets.Keys
        For i = 0 To originalSheets.Count - 1
            On Error Resume Next
            ' Check if the sheet still exists under its original name before deleting
            Dim wsToDelete As Worksheet
            Set wsToDelete = Nothing
            Set wsToDelete = outWb.Worksheets(CStr(k(i)))
            
            If Not wsToDelete Is Nothing Then
                wsToDelete.Delete
                LogMessage "Deleted original sheet: " & k(i), outputDir
            End If
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo ErrorHandler
        Next i
        Application.DisplayAlerts = True
        
        ' 4. Save output workbook as .xlsx (removes VBA and keeps properties)
        If Not outWb Is Nothing Then
            Application.DisplayAlerts = False
            Dim finalFileName As String
            If sheetsPerBook = 1 Then
                finalFileName = outputDir & "\" & outWb.Sheets(1).Name & ".xlsx"
            Else
                finalFileName = outputDir & "\" & timestampStr & "_" & templateName & "_" & fileIndex & ".xlsx"
            End If
            
            LogMessage "Saving final workbook: " & finalFileName, outputDir
            outWb.SaveAs Filename:=finalFileName, _
                FileFormat:=51, ReadOnlyRecommended:=False ' 51 = xlOpenXMLWorkbook
            Application.DisplayAlerts = True
            
            outWb.Close SaveChanges:=False
            fileIndex = fileIndex + 1
        End If
        
        ' Cleanup temp file
        If fso.FileExists(tempPath) Then fso.DeleteFile tempPath
        Application.EnableEvents = True
    Loop
 
    LogMessage "=== Process Completed Successfully ===", outputDir
    Beep
    MsgBox "Process completed successfully.", vbInformation, "Done"

Cleanup:
    On Error Resume Next
    LogMessage "Running Cleanup...", outputDir
    Application.EnableEvents = True
    If Not targetWb Is Nothing Then
        Application.DisplayAlerts = False
        targetWb.Close SaveChanges:=False
        Application.DisplayAlerts = True
    End If
    
    ' --- Restore from Backup and Cleanup ---
    If Not fso Is Nothing Then
        If fso.FileExists(backupPath) Then
            fso.CopyFile backupPath, targetPath, True
            fso.DeleteFile backupPath
        End If
    End If
    
    ' --- OneDrive Restart (Match BindM.bas) ---
    Dim shellApp As Object
    Set shellApp = CreateObject("WScript.Shell")
    On Error Resume Next
    shellApp.Run """" & Environ("LocalAppData") & "\Microsoft\OneDrive\OneDrive.exe""", 1, False
    If Err.Number <> 0 Then LogMessage "Warning: Failed to restart OneDrive.", outputDir
    On Error GoTo 0
    ' -------------------------------------------
    
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Dim errDesc As String
    errDesc = "Error " & Err.Number & ": " & Err.Description
    LogMessage "FATAL ERROR: " & errDesc, outputDir
    MsgBox "An unexpected error occurred: " & errDesc, vbCritical, "System Error"
    Resume Cleanup
End Sub
