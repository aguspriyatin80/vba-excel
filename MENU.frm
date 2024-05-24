VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "UserForm6"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15300
   OleObjectBlob   =   "MENU.frx":0000
   ShowModal       =   0   'False
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
menu = "MENU"
CreateButtons
refreshTabelTransaksi

End Sub

Private Sub CommandButton2_Click()
menu = "MAKANAN"
CreateButtons
refreshTabelTransaksi
End Sub

Private Sub CommandButton3_Click()
menu = "MINUMAN"
CreateButtons
refreshTabelTransaksi
End Sub

Private Sub CommandButton6_Click()
Set ws = Sheets("TRANSAKSI")
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 1 To lastRow
    ws.Range("A2:F" & lastRow).ClearContents
Next i
refreshTabelTransaksi
End Sub

Sub RoundDownValue()
    Set ws = Sheets("MENU")
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    hasil = (lastRow * 100) / 3
    Me.Label1.Caption = WorksheetFunction.RoundDown(hasil, 0)
End Sub
Private Sub ScrollBar1_Change()
On Error Resume Next
Set ws = Sheets("MENU")
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Me.Label1.Caption = Me.ScrollBar1.Value
For i = 1 To lastRow
    If Val(Me.Label1.Caption) = 0 Then
        Me.Controls("Btn" & i).Top = 50
    ElseIf Val(Me.Label1.Caption) > 0 Then
        Me.Controls("Btn" & i).Top = Me.Controls("Btn" & i).Top - (Val(Me.Label1.Caption) / 10)
    End If
    
    
Next i
'Call RoundDownValue

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
  'ws.Activate
  lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
  lastrowSemua = Sheets("MENU").Cells(Rows.Count, "A").End(xlUp).Row
    
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
        .Picture = LoadPicture(ws.Cells(i, "D").Value)
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
    If btn.Top + btn.Height > 300 Then
        Me.ScrollBars = fmScrollBarsVertical
        Me.ScrollHeight = btn.Top + btn.Height + 5
    Else
        Me.ScrollBars = fmScrollBarsNone
    End If
  Next i
  tblTransaksi.Clear
End Sub

Private Sub UserForm_Initialize()
menu = "MENU"
CreateButtons
refreshTabelTransaksi

End Sub

Sub refreshTabelTransaksi()
Set ws = Sheets("TRANSAKSI")
'ws.Select
lastrow2 = Sheets("TRANSAKSI").Cells(Rows.Count, 1).End(xlUp).Row
Me.tblTransaksi.ColumnCount = 6
If lastrow2 <= 1 Then
    Me.tblTransaksi.ColumnHeads = False
    Me.tblTransaksi.RowSource = "A1:f" & lastrow2
Else
    Me.tblTransaksi.ColumnHeads = True
    Me.tblTransaksi.RowSource = "A2:f" & lastrow2
End If
End Sub
