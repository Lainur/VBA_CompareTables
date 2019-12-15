VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TableMatch_frm 
   Caption         =   "Tables matching"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5970
   OleObjectBlob   =   "TableMatch_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TableMatch_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public TableRange1 As Range
Public TableRange2 As Range

Private Sub Cb_Datas_Click()
    If Me.Cb_Datas.Value = False Then
        Cb_DataPatterns.Enabled = False
        Cb_DataPatterns.Value = False
    Else
        Cb_DataPatterns.Enabled = True
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim r As Range
    On Error Resume Next
    Set r = Application.InputBox("Select a cell in or whole Table 1", "Table 1", , Type:=8)
    If Not (r Is Nothing) Then
        If r.cells.count = 1 Then
            Set TableRange1 = r.CurrentRegion
        Else
            Set TableRange1 = r
        End If
        Me.TextBox1 = "'" & TableRange1.Parent.Name & "'!" & TableRange1.Address
        Me.TextBox2 = 1
        If Me.TextBox5 = vbNullString Then Me.TextBox5 = TableRange1.Parent.Name
    End If
    On Error GoTo 0
End Sub
Private Sub CommandButton2_Click()
    On Error Resume Next
    Dim Ok As Boolean, KeyStrings As String
    If Not (Me.TextBox1 = vbNullString) Then
        TableRange1.Parent.Activate
        Dim r As Range
        Set r = Application.InputBox("Key Columns in Table 1", "Key Columns", , Type:=8)
        If Not (r Is Nothing) Then
            Ok = True
            For i = 1 To r.Areas.count
                If r.Areas(i).cells(1).Column <= TableRange1.cells(1).Column And r.Areas(i).cells(1).Column >= TableRange1.cells(TableRange1.cells.count).Column Then
                    Ok = False: Exit For
                Else
                    KeyStrings = KeyStrings & IIf(KeyStrings = "", "", " ") & CStr(1 + r.Areas(i).cells(1).Column - TableRange1.cells(1).Column)
                End If
            Next
            If Ok Then
                Me.TextBox2 = KeyStrings
            Else
                MsgBox "Some key columns outside the table area !"
            End If
        End If
    
    Else
        MsgBox "Select Table range first"
    End If
    On Error GoTo 0
End Sub

Private Sub CommandButton3_Click()
    Dim r As Range
    On Error Resume Next
    Set r = Application.InputBox("Select a cell in or whole Table 2", "Table 2", , Type:=8)
    If Not (r Is Nothing) Then
        If r.cells.count = 1 Then
            Set TableRange2 = r.CurrentRegion
        Else
            Set TableRange2 = r
        End If
        Me.TextBox3 = "'" & TableRange2.Parent.Name & "'!" & TableRange2.Address
        Me.TextBox4 = 1
        If Me.TextBox6 = vbNullString Then Me.TextBox6 = TableRange2.Parent.Name
    End If
    On Error GoTo 0
End Sub
Private Sub CommandButton4_Click()
    On Error Resume Next
    Dim Ok As Boolean, KeyStrings As String
    If Not (Me.TextBox3 = vbNullString) Then
        TableRange2.Parent.Activate
        Dim r As Range
        Set r = Application.InputBox("Key Columns in Table 2", "Key Columns", , Type:=8)
        If Not (r Is Nothing) Then
            Ok = True
            For i = 1 To r.Areas.count
                If r.Areas(i).cells(1).Column <= TableRange2.cells(1).Column And r.Areas(i).cells(1).Column >= TableRange2.cells(TableRange2.cells.count).Column Then
                    Ok = False: Exit For
                Else
                    KeyStrings = KeyStrings & IIf(KeyStrings = "", "", " ") & CStr(1 + r.Areas(i).cells(1).Column - TableRange2.cells(1).Column)
                End If
            Next
            If Ok Then
                Me.TextBox4 = KeyStrings
            Else
                MsgBox "Some key columns outside the table area !"
            End If
        End If
    
    Else
        MsgBox "Select Table range first"
    End If
    On Error GoTo 0
End Sub

Private Sub CommandButton5_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    If ActiveCell.CurrentRegion.cells.count > 100 Then
        ActiveCell.CurrentRegion.cells.Select
        Set TableRange1 = Selection
        Me.TextBox1 = "'" & TableRange1.Parent.Name & "'!" & TableRange1.Address
        Me.TextBox2 = 1
        If Me.TextBox5 = vbNullString Then Me.TextBox5 = TableRange1.Parent.Name
    End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then End
End Sub
