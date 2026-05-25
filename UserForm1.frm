VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Data Embedding Macro"
   ClientHeight    =   3380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =====================================================
'  データ埋込マクロ - UserForm1
'  Copyright (c) 2026 Datan (データン)
'  Licensed under the MIT License.
' =====================================================

' --- Execute Button ---
Private Sub CommandButton1_Click()
    If (TextBox1.Text <> "" And TextBox2.Text <> "" And TextBox3.Text <> "" _
        And ComboBox1.Text <> "" And ComboBox2.Text <> "") Then

        Call ProcessDataEmbedding( _
            TextBox1.Text, _
            ComboBox1.Text, _
            ComboBox2.Text, _
            CLng(TextBox2.Text), _
            CLng(TextBox3.Text))

    Else
        MsgBox "Please fill in all fields.", vbCritical, "Input Error"
        
        If TextBox1.Text = "" Then
            CommandButton2.SetFocus
        ElseIf ComboBox1.Text = "" Then
            ComboBox1.SetFocus
        ElseIf ComboBox2.Text = "" Then
            ComboBox2.SetFocus
        ElseIf TextBox2.Text = "" Then
            TextBox2.SetFocus
        Else
            TextBox3.SetFocus
        End If
    End If
End Sub

' --- Browse Button ---
Private Sub CommandButton2_Click()
    Dim filePath As Variant
    
    ' Select File
    filePath = Application.GetOpenFilename("Excel Files (*.xls*),*.xls*", , "Select Workbook")

    If filePath <> False Then
        UserForm1.TextBox1.Value = filePath
        Application.ScreenUpdating = False
        
        Dim wb As Workbook
        Set wb = Workbooks.Open(Filename:=filePath, ReadOnly:=True)
        
        UserForm1.ComboBox1.Clear
        UserForm1.ComboBox2.Clear
        
        Dim i As Integer
        For i = 1 To wb.Sheets.Count
            UserForm1.ComboBox1.AddItem wb.Sheets(i).Name
            UserForm1.ComboBox2.AddItem wb.Sheets(i).Name
        Next i
        
        wb.Close SaveChanges:=False
        Application.ScreenUpdating = True
    End If
End Sub

' --- Form Initialize ---
Private Sub UserForm_Initialize()
    ' Set window title with version
    Me.Caption = "Data Embedding Macro v" & Module1.APP_VERSION

    ' Ensure the sheets per book field is editable
    With TextBox3
        .Enabled = True
        .Locked = False
        .Value = "1"
    End With
    
    ' Initialize Start Row (Default to 2 for headers in row 1)
    If TextBox2.Value = "" Then TextBox2.Value = "2"

    ' Attempt to "delete" the copyright label
    ' Note: VBA design-time controls cannot be truly "Removed" at runtime, 
    ' but we hide and clear it to make it effectively non-existent.
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Label" Then
            If InStr(ctrl.Caption, "Copyright") > 0 Then
                On Error Resume Next
                ctrl.Caption = ""
                ctrl.Visible = False
                Me.Controls.Remove ctrl.Name ' This may fail for design-time controls
                On Error GoTo 0
            End If
        End If
    Next ctrl
End Sub

' --- Reset Button ---
Private Sub CommandButton3_Click()
    UserForm1.TextBox1.Value = ""
    UserForm1.ComboBox1.Value = ""
    UserForm1.ComboBox2.Value = ""
    UserForm1.TextBox3.Value = "1"
    UserForm1.Hide
End Sub

' --- Numeric Validation ---
Private Sub TextBox2_Change()
    CheckNumber TextBox2
End Sub

Private Sub TextBox3_Change()
    CheckNumber TextBox3
End Sub

Private Sub CheckNumber(ctrl As MSForms.TextBox)
    If ctrl.Text <> "" Then
        If Not IsNumeric(ctrl.Text) Or Val(ctrl.Text) < 1 Or Val(ctrl.Text) > 1048576 Then
            MsgBox "Please enter a valid number.", vbCritical, "Validation Error"
            ctrl.Text = "1"
            ctrl.SetFocus
        End If
    End If
End Sub
