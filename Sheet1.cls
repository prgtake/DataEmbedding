VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    ' Close all workbooks except this one
    Dim ThisBook As String
    ThisBook = ActiveWorkbook.Name
    
    Dim b As Integer, e As Integer
    If Workbooks.Count > 1 Then
        e = 0
        For b = 1 To Workbooks.Count
            If Workbooks(b - e).Name <> ThisBook Then
                Workbooks(b - e).Close SaveChanges:=True
                e = e + 1
            End If
        Next b
    End If
    
    UserForm1.Show
End Sub
