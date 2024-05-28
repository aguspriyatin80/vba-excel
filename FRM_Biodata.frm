VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_Biodata 
   Caption         =   "BIODATA"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13440
   OleObjectBlob   =   "FRM_Biodata.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_Biodata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB_Batal_Click()
KondisiAwal
End Sub

Private Sub CB_Hapus_Click()
On Error Resume Next
If Me.LB_Siswa.ListIndex = -1 Then
    MsgBox "Klik data yang akan dihapus"
Else
Dim nama As String
nama = Me.LB_Siswa.List(LB_Siswa.ListIndex, 1)
lastRow = Sheets("SISWA").Cells(Rows.Count, 1).End(xlUp).Row
Set dtSiswa = Sheets("SISWA").Range("A2:A" & lastRow).Find(what:=nama, LookIn:=xlValues, LookAt:=xlWhole)
Set dtTransaksi = Sheets("V_SISWA").Range("b2:b" & lastRow).Find(what:=nama, LookIn:=xlValues, LookAt:=xlWhole)
Baris = dtTransaksi.Row
    If MsgBox("Apakah anda yakin akan menghapus transaksi " & nama & "?", vbYesNo + vbQuestion, "MICROSHOP") = vbYes Then
        Sheets("SISWA").Cells(Baris, "a").Delete Shift:=xlUp
        Sheets("SISWA").Cells(Baris, "b").Delete Shift:=xlUp
        Sheets("SISWA").Cells(Baris, "c").Delete Shift:=xlUp
        Sheets("V_SISWA").Cells.ClearContents
        RefreshTabel
    End If
End If
End Sub
Private Sub CB_Simpan_Click()
Dim ws As Worksheet
Set ws = Sheets("SISWA")
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Range("A" & lastRow + 1).Value = Me.TB_Nama
ws.Range("B" & lastRow + 1).Value = Me.TB_Alamat
RefreshTabel
KosongkanForm
End Sub
Sub BuatViewSiswa()
Dim newSheetName As String
Dim newSheet As Worksheet
newSheetName = "V_SISWA"
If sheetExists(newSheetName) = False Then
    With ThisWorkbook
        .Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = newSheetName
    End With
End If
IsiViewSiswa
lastRow = Sheets(newSheetName).Cells(Rows.Count, 1).End(xlUp).Row
Me.LB_Siswa.ColumnCount = 3
Me.LB_Siswa.ColumnHeads = True
Me.LB_Siswa.RowSource = "A2:C" & lastRow
End Sub
Private Sub IsiViewSiswa()
On Error Resume Next
lastRow = Sheets("SISWA").Cells(Rows.Count, 1).End(xlUp).Row
Sheets("SISWA").Range("A1:C" & lastRow).Copy
'Activate the destination worksheet
Sheets("V_SISWA").Activate
'Select the target range
Range("B1").Select
'Paste in the target destination
ActiveSheet.Paste
lastRow = Sheets("V_SISWA").Cells(Rows.Count, 2).End(xlUp).Row
Sheets("V_SISWA").Range("A1").Value = "No."
Sheets("V_SISWA").Range("A1").Font.Bold
For i = 2 To lastRow
    Sheets("V_SISWA").Range("A" & i).Value = "=ROW()-ROW($A$1)"
Next i
Application.CutCopyMode = False
Range("A1").Select
End Sub
Private Sub LB_Siswa_Click()
Me.CB_Hapus.Enabled = True
Me.TB_Nama.Value = Me.LB_Siswa.Column(1)
Me.TB_Alamat.Value = Me.LB_Siswa.Column(2)
Me.CB_Simpan.Caption = "UPDATE"
End Sub
Private Sub LB_Siswa_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
MsgBox Me.LB_Siswa.Column(0)
End Sub
Private Sub UserForm_Activate()
If Me.LB_Siswa.ListIndex = -1 Then
    Me.CB_Hapus.Enabled = False
Else
    Me.CB_Hapus.Enabled = True
End If
End Sub
Sub KondisiAwal()
KosongkanForm
Me.LB_Siswa.ListIndex = -1
Me.CB_Hapus.Enabled = False
Me.CB_Simpan.Caption = "SIMPAN"
RefreshTabel
End Sub
Private Sub UserForm_Initialize()
KondisiAwal
End Sub
Sub RefreshTabel()
BuatViewSiswa
HitungJumlahSiswa
End Sub
Sub KosongkanForm()
Me.TB_Nama.Value = ""
Me.TB_Alamat.Value = ""
Me.TB_Nama.SetFocus
End Sub
Sub HitungJumlahSiswa()
JmlSiswa = Me.LB_Siswa.ListCount
Me.LBL_JumlahSiswa.Caption = JmlSiswa
End Sub
Function sheetExists(sheetToFind As String, Optional InWorkbook As Workbook) As Boolean
    If InWorkbook Is Nothing Then Set InWorkbook = ThisWorkbook
    On Error Resume Next
    sheetExists = Not InWorkbook.Sheets(sheetToFind) Is Nothing
End Function
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim newSheetName As String
newSheetName = "V_SISWA"
If sheetExists(newSheetName) = True Then
    Application.DisplayAlerts = False
    Sheets(newSheetName).Delete
    Application.DisplayAlerts = True
End If
End Sub
