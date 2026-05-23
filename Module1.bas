Attribute VB_Name = "Module1"
' =====================================================
'  データ埋込マクロ (Data Embedding Macro)
'  Copyright (c) 2026 Datan (データン)
'  Licensed under the MIT License.
' =====================================================
Option Explicit

' --- Version Management ---
Public Const APP_VERSION As String = "1.0.0"

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
    MkDir outputDir
    ChDir outputDir
    Kill outputDir & "\*.xlsx"
    On Error GoTo ErrorHandler
    
    ' Create backup
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
    Set targetWb = Workbooks.Open(Filename:=targetPath, ReadOnly:=True)
    
    ' Disable AutoSave if applicable (OneDrive/SharePoint)
    On Error Resume Next
    targetWb.AutoSaveOn = False
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Set dataWs = targetWb.Worksheets(dataName)
    Set templateWs = targetWb.Worksheets(templateName)
    On Error GoTo ErrorHandler
    
    If dataWs Is Nothing Or templateWs Is Nothing Then
        MsgBox "Specified sheets were not found.", vbCritical, "Error"
        GoTo Cleanup
    End If

    ' Identify default data row
    defaultDataRow = startRow + 1

    ' --- Duplicate Check (Optimized with Dictionary) ---
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    currentRow = defaultDataRow
    
    Do While dataWs.Cells(currentRow, 1).Value <> ""
        Dim keyVal As String
        keyVal = CStr(dataWs.Cells(currentRow, 1).Value)
        
        If dict.exists(keyVal) Then
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
    Do While dataWs.Cells(currentRow, 1).Value <> ""
        
        Set outWb = Nothing
        sheetCount = 0
        
        ' Loop for sheets per book
        Do While sheetCount < sheetsPerBook
            newFileName = dataWs.Cells(currentRow, 1).Value
            If newFileName = "" Then Exit Do
            
            ' Injection Method: Copy current record to the default row
            If currentRow <> defaultDataRow Then
                dataWs.Rows(currentRow).Copy Destination:=dataWs.Rows(defaultDataRow)
            End If
            
            ' Copy updated template
            If outWb Is Nothing Then
                templateWs.Copy
                Set outWb = ActiveWorkbook
            Else
                templateWs.Copy After:=outWb.Sheets(outWb.Sheets.Count)
            End If
            
            ' Set Sheet Name
            On Error Resume Next
            ActiveSheet.Name = Left(Replace(Replace(Replace(Replace(Replace(Replace(Replace(newFileName, "\", ""), "/", ""), ":", ""), "?", ""), "*", ""), "[", ""), "]", ""), 31)
            On Error GoTo ErrorHandler
            
            ' Paste Values
            Cells.Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            Cells(1, 1).Select
            
            currentRow = currentRow + 1
            sheetCount = sheetCount + 1
            
            ' End of data check
            If dataWs.Cells(currentRow, 1).Value = "" Then Exit Do
        Loop
        
        ' Save output workbook
        If Not outWb Is Nothing Then
            Application.DisplayAlerts = False
            If sheetsPerBook = 1 Then
                outWb.SaveAs Filename:=outputDir & "\" & outWb.Sheets(1).Name & ".xlsx", _
                    FileFormat:=xlOpenXMLWorkbook, ReadOnlyRecommended:=False
            Else
                outWb.SaveAs Filename:=outputDir & "\" & timestampStr & "_" & templateName & "_" & fileIndex & ".xlsx", _
                    FileFormat:=xlOpenXMLWorkbook, ReadOnlyRecommended:=False
            End If
            Application.DisplayAlerts = True
            outWb.Close SaveChanges:=False
            fileIndex = fileIndex + 1
        End If
        
    Loop
 
    Beep
    MsgBox "Process completed successfully.", vbInformation, "Done"

Cleanup:
    On Error Resume Next
    If Not targetWb Is Nothing Then
        Application.DisplayAlerts = False
        targetWb.Close SaveChanges:=False
        Application.DisplayAlerts = True
    End If
    
    ' --- Restore from Backup and Cleanup ---
    If Not fso Is Nothing Then
        If fso.FileExists(backupPath) Then
            ' Attempt to restore. If targetPath is locked, this might fail, but since we opened as ReadOnly, 
            ' targetPath is likely unmodified anyway.
            fso.CopyFile backupPath, targetPath, True
            fso.DeleteFile backupPath
        End If
    End If
    
    ' --- OneDrive Restart (Match BindM.bas) ---
    Dim shellApp As Object
    Set shellApp = CreateObject("WScript.Shell")
    On Error Resume Next
    shellApp.Run """" & Environ("LocalAppData") & "\Microsoft\OneDrive\OneDrive.exe""", 1, False
    On Error GoTo 0
    ' -------------------------------------------
    
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "System Error"
    Resume Cleanup
End Sub
