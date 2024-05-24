VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "UserForm6"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15300
   OleObjectBlob   =   "galleryFix.frx":0000
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
Dim picPath() As Variant
Dim ws As Worksheet

'Set dt = Sheets("MENU")
'lastRow = dt.Cells(Rows.Count, 1).End(xlUp).Row
Dim DPD(1 To 1000) As New Class2

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


Sub isi()
id = Me.btEvents.Caption
jumlah = 1
akhir = Sheets("TRANSAKSI").Cells(Rows.Count, 1).End(xlUp).Row
akhir1 = Sheets("SEMENTARA").Cells(Rows.Count, 1).End(xlUp).Row
akhir2 = Sheets("MENU").Cells(Rows.Count, 1).End(xlUp).Row
Set Datanya = Sheets("MENU").Range("a2:a" & akhir2).Find(what:=id, LookIn:=xlValues, LookAt:=xlWhole)
If Datanya Is Nothing Then
'MsgBox "ID barang tidak ditemukan"
Exit Sub
ElseIf jumlah = 0 Then
MsgBox "Jumlah tidak boleh kosong"
Exit Sub
Else
        
    Set cekKode = Sheets("TRANSAKSI").Range("b2:B" & akhir).Find(what:=id, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cekKode Is Nothing Then
        Sheets("TRANSAKSI").Range("a" & akhir + 1).Value = "=Row()-1"
        Sheets("TRANSAKSI").Range("c" & akhir + 1).Value = Datanya.Offset(0, 1).Value
        Sheets("TRANSAKSI").Range("b" & akhir + 1).Value = Datanya.Offset(0, 0).Value
        Sheets("TRANSAKSI").Range("d" & akhir + 1).Value = jumlah
        
        'Sheets("TRANSAKSI").Range("e" & akhir + 1).Value = Datanya.Offset(0, 2).Value * jumlah
        Sheets("TRANSAKSI").Range("e" & akhir + 1).Value = jumlah * Datanya.Offset(0, 2).Value
        Sheets("TRANSAKSI").Range("f" & akhir + 1).Value = jumlah * Datanya.Offset(0, 2).Value * jumlah
        
        Sheets("SEMENTARA").Range("a" & akhir1 + 1).Value = "=Row()-1"
        Sheets("SEMENTARA").Range("b" & akhir1 + 1).Value = Datanya.Offset(0, 0).Value
        Sheets("SEMENTARA").Range("c" & akhir1 + 1).Value = Datanya.Offset(0, 1).Value
        Sheets("SEMENTARA").Range("d" & akhir1 + 1).Value = jumlah
        Sheets("SEMENTARA").Range("e" & akhir1 + 1).Value = Datanya.Offset(0, 2).Value
   Else
    cekKode.Offset(0, 2).Value = cekKode.Offset(0, 2).Value + 1
    'cekKode.Offset(0, 4).Value = jumlah * cekKode.Offset(0, 3).Value
    cekKode.Offset(0, 4).Value = cekKode.Offset(0, 2).Value * cekKode.Offset(0, 3).Value
   End If
End If
refreshTabelTransaksi
End Sub

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
If lastRow = 1 Then Me.tblTransaksi.RowSource = ""
For i = 2 To lastRow
    ws.Range("A" & i & ":F" & lastRow).ClearContents
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





Private Sub UserForm_Initialize()
menu = "MENU"
refreshTabelTransaksi
CreateButtons
End Sub

Sub refreshTabelTransaksi()
Set ws = Sheets("TRANSAKSI")
'ws.Select
lastrow2 = Sheets("TRANSAKSI").Cells(Rows.Count, 1).End(xlUp).Row
Me.tblTransaksi.ColumnCount = 6
If lastrow2 <= 1 Then
    Me.tblTransaksi.ColumnHeads = False
    'Me.tblTransaksi.RowSource = "A1:f" & lastrow2
    Me.tblTransaksi.RowSource = ""
Else
    Me.tblTransaksi.ColumnHeads = True
    Me.tblTransaksi.RowSource = "A2:f" & lastrow2
End If
End Sub

Sub CreateButtons()
  On Error Resume Next
  Dim btn As MSForms.CommandButton
  Dim frm As MSForms.Frame
  Set frm = Me.Controls.Add("Forms.Frame.1")
   frm.Name = "Frame1"
    frm.Top = 50
    frm.Left = 10
    frm.Width = 400
    frm.Height = 300
    frm.SpecialEffect = fmSpecialEffectFlat
    
    
  Set ws = Sheets(menu)
  'ws.Activate
  lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
  lastrowSemua = Sheets("MENU").Cells(Rows.Count, "A").End(xlUp).Row
    
  LeftPos = 0
  TopPos = 20
  
  For k = 1 To lastrowSemua
    frm.Controls.Remove ("Btn" & k)
  Next k
  'i = 2
  For i = 2 To lastRow
    
    
    
    Set btn = frm.Controls.Add("Forms.CommandButton.1")
    
    With btn
       
      .Caption = Sheets(menu).Range("A" & i).Value
      '.Caption = "Btn" & i
      .Name = "Btn" & i
        .Height = 100
        .Width = 100
        .Left = LeftPos
        .Top = TopPos
        '.Picture = LoadPicture("C:\Users\agus\Documents\ALL\Bingka Ambon.jpg")
        .Picture = LoadPicture(ws.Cells(i, "D").Value)
        .TakeFocusOnClick = False
        '.Picture = LoadPicture(picPath)
        LeftPos = LeftPos + .Width + 5
        'TopPos = LeftPos + .Width + 1
        Pic = Pic + 1
        Set DPD(i).btEvents = btn
    End With
    If Pic = 3 Then
        TopPos = TopPos + btn.Height + 5
        LeftPos = 0
        Pic = 0
    End If
    If btn.Top + btn.Height > 300 Then
        frm.ScrollBars = fmScrollBarsVertical
        frm.ScrollHeight = btn.Top + btn.Height + 5
    Else
        frm.ScrollBars = fmScrollBarsNone
    End If
  Next i
  tblTransaksi.Clear
End Sub

