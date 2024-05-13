VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LOOKUP 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6810
   OleObjectBlob   =   "LOOKUP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LOOKUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function wsExits(shName As String) As Boolean
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    If ws.Name = shName Then
        wsExits = True
        Exit Function
    End If
Next ws

wsExits = False
End Function

Sub AddSheetBarangSupplier()
Dim wsNew As Worksheet
If Not wsExits("BARANG_SUPPLIER") Then
    Set wsNew = ThisWorkbook.Worksheets.Add
    wsNew.Name = "BARANG_SUPPLIER"
    
Else
    Application.DisplayAlerts = False
    Sheets("BARANG_SUPPLIER").Delete
    Set wsNew = ThisWorkbook.Worksheets.Add
    wsNew.Name = "BARANG_SUPPLIER"
    Application.DisplayAlerts = True
End If

'Dim i As Long, newWsName As String
'Do
'    i = i + 1
'    newWsName = "BARANG_SUPPLIER_" & i
'    If Not wsExits(newWsName) Then
'        Set wsNew = ThisWorkbook.Sheets.Add
'        wsNew.Name = newWsName
'        Exit Sub
'    End If
'Loop
End Sub

Private Sub UserForm_Initialize()
Dim ws As Worksheet
Dim lastRow As Long
Dim namaCustomer As Range
Dim namaPerusahaan As Range
Dim myData As Range
Dim idBarang As Range
Set ws = Sheets("BARANG")

ws.Activate

'lastRow = ws.Cells(Rows.Count, "E").End(xlUp).Row

AddSheetBarangSupplier
ws.Range("a1:e5").Copy Destination:=Sheets("BARANG_SUPPLIER").Range("A1")
Sheets("BARANG_SUPPLIER").Range("F1").Value = "NAMA_CUSTOMER"
Sheets("BARANG_SUPPLIER").Range("G1").Value = "PERUSAHAAN"

lastRow = Sheets("BARANG_SUPPLIER").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastRow
    Set myData = Sheets("SUPPLIER").Range("B1:d4")
    Set idSup = Sheets("BARANG_SUPPLIER").Range("E" & i)
    Set namaCustomer = Sheets("BARANG_SUPPLIER").Range("F" & i)
    Set namaPerusahaan = Sheets("BARANG_SUPPLIER").Range("G" & i)
    namaCustomer = Application.WorksheetFunction.VLookup(idSup, myData, 2, False)
    namaPerusahaan = Application.WorksheetFunction.VLookup(idSup, myData, 3, False)
    Sheets("BARANG_SUPPLIER").Range("F" & i) = namaCustomer
    Sheets("BARANG_SUPPLIER").Range("G" & i) = namaPerusahaan
    Sheets("BARANG_SUPPLIER").Range("D" & i).NumberFormat = "#,##0"
Next i

Me.ListBox1.ColumnCount = 7
Me.ListBox1.ColumnHeads = True
Me.ListBox1.RowSource = "BARANG_SUPPLIER!a2:g5"
Me.ListBox1.ColumnWidths = ";;;;0;;"

End Sub
