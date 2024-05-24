VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents btEvents As MSForms.CommandButton
Attribute btEvents.VB_VarHelpID = -1
Dim menu As String
Private Sub btEvents2_click()
On Error Resume Next
    'Dim btEx As MSForms.CommandButton
    'Set btEx = UserForm7.Controls.Add("Forms.CommandButton.1")
    
    'With btEx
    '    .Top = 30 * UserForm7.Controls.Count
    '    .Left = 30
    '    .Width = 72
    '    .Height = 36
    '    .Caption = "Click Me"
    'End With

Dim CompId As Range
Dim id As String

id = Me.btEvents.Caption
menu = "MENU"
Set CompId = Sheets(menu).Range("A:A").Find(what:=id, LookIn:=xlValues, LookAt:=xlWhole)
If Not CompId Is Nothing Then
    lastRow = Sheets("TRANSAKSI").Cells(Rows.Count, 1).End(xlUp).Row
    qty = 1

    id2 = id
    Set cekPesanan = Sheets("TRANSAKSI").Range("A2:A" & lastRow).Find(what:=id2, LookIn:=xlValues, LookAt:=xlWhole)
    If Not cekPesanan Is Nothing Then
        Sheets("TRANSAKSI").Offset(0, 2).Value = Sheets("TRANSAKSI").Offset(0, 2).Value + 1
        cekPesanan.Offset(0, 4).Value = cekPesanan.Offset(0, 2).Value * cekPesanan.Offset(0, 3).Value
    Else

        Sheets("TRANSAKSI").Range("A" & lastRow + 1).Value = cekPesanan.Offset(0, 0).Value
        Sheets("TRANSAKSI").Range("B" & lastRow + 1).Value = CompId.Offset(0, 1).Value
        Sheets("TRANSAKSI").Range("C" & lastRow + 1).Value = qty + 1
        Sheets("TRANSAKSI").Range("D" & lastRow + 1).Value = CompId.Offset(0, 3).Value
        Sheets("TRANSAKSI").Range("E" & lastRow + 1).Value = Sheets("TRANSAKSI").Range("C" & lastRow + 1).Value * Sheets("TRANSAKSI").Range("D" & lastRow + 1).Value

    End If
    
End If
End Sub

Private Sub btEvents_click()

'jumlah = InputBox("Masukkan Jumlah Barang", "APLIKASI KASIR")
'Sheets("MENU").AutoFilterMode = False
isi

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
UserForm6.refreshTabelTransaksi
End Sub

