VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "UserForm6"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15300
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim menu As String
Private col As Collection
Public WithEvents btEvents As MSForms.CommandButton
Attribute btEvents.VB_VarHelpID = -1
Dim picPath() As Variant
Dim ws As Worksheet


Private Sub CommandButton1_Click()
menu = "SEMUA"
CreateButtons

End Sub

Private Sub CommandButton2_Click()
menu = "MAKANAN"
CreateButtons
End Sub

Private Sub CommandButton3_Click()
menu = "MINUMAN"
CreateButtons
End Sub

Private Sub UserForm_AddControl(ByVal Control As MSForms.Control)
On Error Resume Next
Dim cl As Class2
Dim ctl As MSForms.CommandButton
Set col = New Collection
For Each ctl In Me.Controls
    Set cl = New Class2
    Set cl.btEvents = ctl
    col.Add cl
Next ctl
End Sub

Sub CreateButtons()
  On Error Resume Next
  Dim btn As MSForms.CommandButton
  Set ws = Sheets(menu)
  ws.Activate
  lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
  lastrowSemua = Sheets("SEMUA").Cells(Rows.Count, "A").End(xlUp).Row
  
  
  LeftPos = 0
  TopPos = 50
  
  For k = 1 To lastrowSemua
    Me.Controls.Remove ("Btn" & k)
  Next k
  'i = 2
  For i = 2 To lastRow
    
    Set btn = Me.Controls.Add("Forms.CommandButton.1")
    
    With btn
       
      .Caption = Sheets(menu).Range("A" & i).Value
      '.Caption = "Btn" & i
      .name = "Btn" & i
        .Height = 100
        .Width = 100
        .Left = LeftPos
        .Top = TopPos
        '.Picture = LoadPicture("C:\Users\agus\Documents\ALL\Bingka Ambon.jpg")
        .Picture = LoadPicture(ws.Cells(i, "E").Value)
        '.Picture = LoadPicture(picPath)
        LeftPos = LeftPos + .Width + 5
        'TopPos = LeftPos + .Width + 1
        Pic = Pic + 1
    End With
    If Pic = 3 Then
        TopPos = TopPos + btn.Height + 5
        LeftPos = 0
        Pic = 0
    End If
  Next i
  ListBox1.Clear
End Sub

Private Sub UserForm_Initialize()
menu = "SEMUA"
CreateButtons
Set ws = Sheets("PESANAN")
ws.Select
lastrow2 = Sheets("PESANAN").Cells(Rows.Count, 1).End(xlUp).Row
Me.ListBox1.ColumnCount = 5
If lastrow2 = 1 Then
    'Me.ListBox1.ColumnHeads = False
    Me.ListBox1.RowSource = ""
Else
    Me.ListBox1.ColumnHeads = True
    Me.ListBox1.RowSource = "A2:E" & lastrow2
End If



End Sub
