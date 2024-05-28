VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_MultiselectListBox 
   Caption         =   "UserForm1"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11205
   OleObjectBlob   =   "FRM_MultiselectListBox.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_MultiselectListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myHobbies As String

Sub massupdate()
Dim SelectedRows()
idx = 0
ReDim SelectedRows(1 To 1)
With Me.ListBox1 'name of listbox
  For i = 0 To .ListCount - 1
    If .Selected(i) Then
      idx = idx + 1
      ReDim Preserve SelectedRows(1 To idx)
      SelectedRows(idx) = i
    End If
  Next i
  Set yyy = Range(.RowSource)
  myId = Me.TB_ID.Value 'recruiter employee id and recruiter name are linked
  myName = Me.TB_Name.Value 'recruiter name and recruiter employee id are linked
   
For Each SelectedRow In SelectedRows
   yyy.Cells(SelectedRow + 1, 1).Value = cmbRecruiterEmplID 'for each selection made in the listbox - update the recruiter emplid selected in the combobox to 13th column in spreadsheet
   yyy.Cells(SelectedRow + 1, 2).Value = cmbRecruiterName  'for each selection made in the listbox - update the recruiter employeename selected in the combobox to 14th column in spreadsheet
  Next SelectedRow 'make updates for each selection in listbox

End With

Function GetHobbies()
Dim SelectedRows()
idx = 0
ReDim SelectedRows(1 To 1)
With Me.ListBox2
  For i = 0 To .ListCount - 1
    If .Selected(i) Then
      idx = idx + 1
      ReDim Preserve SelectedRows(1 To idx)
      SelectedRows(idx) = i
    End If
  Next i
  Set yyy = Range(.RowSource)
  textcheckdate = hobbies
  For Each SelectedRow In SelectedRows
   yyy.Cells(SelectedRow + 1, 3).Value = textcheckdate
   GetHobbies = textcheckdate
  Next SelectedRow
  'reinstate selected items:
  'For Each SelectedRow In SelectedRows
  '  ListBox1.Selected(SelectedRow) = True
  'Next SelectedRow
End With
End Function

Sub refreshTables()
lastRow = Sheets("BUAH").Cells(Rows.Count, 1).End(xlUp).Row
Me.ListBox1.RowSource = "BUAH!B2:B" & lastRow
lastRow2 = Sheets("ORANG").Cells(Rows.Count, 1).End(xlUp).Row
Me.ListBox2.ColumnCount = 3
Me.ListBox2.ColumnHeads = True
Me.ListBox2.RowSource = "ORANG!A2:C" & lastRow2
End Sub
Sub selectHobbies()
Dim StringValue As String
StringValue = Me.ListBox2.Column(2)
Dim SingleValue() As String
SingleValue = Split(StringValue, ", ")
'ukuranLB1 = UBound(SingleValue)
'ukuranLB2 = Me.ListBox1.ListCount
Dim stVar As String
    For i = 0 To Me.ListBox1.ListCount - 1
        For j = 0 To UBound(SingleValue)
            'stVar = SingleValue(j)
            If Me.ListBox1.Column(0, i) = SingleValue(j) Then
              Me.ListBox1.Selected(i) = True
              'Exit Sub
              'MsgBox SingleValue
            End If
        Next j
    Next
End Sub

Function namaBuah()
Dim nama As Variant
Set ws = Sheets("BUAH")
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 0 To Me.ListBox1.ListCount - 1
    nama = Me.ListBox1.List(i, 0)
    'MsgBox Me.ListBox1.Column(0)
Next i
namaBuah = nama
End Function


Sub kodeOtomatis()
Dim aktifSheet As Worksheet
Set aktifSheet = Sheets("ORANG")
lastRow = aktifSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
kode_otomatis = "CUS" & Format(Right(aktifSheet.Cells(lastRow - 1, 1), 3) + 1, "0##")
Me.TB_ID.Text = kode_otomatis
End Sub

Sub kosongkan()
Me.TB_Nama.Value = ""
End Sub
Function hobbies()
myVar = ""
For x = 0 To ListBox1.ListCount - 1
    If Me.ListBox1.Selected(x) = True Then
        If myVar = "" Then
            myVar = Me.ListBox1.List(x, 0)
        Else
            myVar = myVar & ", " & Me.ListBox1.List(x, 0)
        End If
    End If
Next x
hobbies = myVar
End Function

Function hobbies2()
myVar = ""
For x = 0 To ListBox1.ListCount - 1
    If Me.ListBox1.Selected(x) = True Then
        If myVar = "" Then
            myVar = Me.ListBox1.List(x, 0)
        Else
            myVar = myVar & ", " & Me.ListBox1.List(x, 0)
        End If
    End If
Next x
hobbies2 = myVar
End Function

Private Sub cmdSimpan_Click()

Dim Baris As Integer
Dim aktifSheet As Worksheet
Dim newHobbies As String
Set aktifSheet = Sheets("ORANG")
Set colId = Sheets("ORANG").Range("A:A")
Set ketemu = colId.Find(Me.TB_ID.Value, LookIn:=xlValues, lookat:=xlWhole)
'MsgBox hobbies
'Stop
hobbiesUpdated = hobbies
lastRow = aktifSheet.Cells(Rows.Count, 1).End(xlUp).Row
If Not ketemu Is Nothing Then
    Baris = ketemu.Row
    colId.Cells(Baris, 2) = Me.TB_Nama.Value
    'colId.Cells(Baris, 3) = "UPDATED"
    'For i = 0 To Me.ListBox1.ListCount - 1
    
        'Me.ListBox1.SetFocus
        'If Me.ListBox1.Selected(i) Then
            'colId.Cells(Baris, 3).Value = Me.ListBox1.List(i)
            
            aktifSheet.Cells(Baris, 3).Value = hobbiesUpdated
            'Baris = Baris + 1
        'End If
    'Next i
        
Else
    aktifSheet.Cells(lastRow + 1, 1) = Me.TB_ID.Value
    aktifSheet.Cells(lastRow + 1, 2) = Me.TB_Nama.Value
    aktifSheet.Cells(lastRow + 1, 3) = hobbies
    
End If

For i = 0 To ListBox1.ListCount - 1
Me.ListBox1.Selected(i) = False
Next i

For j = 0 To ListBox2.ListCount - 1
Me.ListBox2.Selected(j) = False
Next j
lastRow2 = Sheets("BUAH").Cells(Rows.Count, 1).End(xlUp).Row
Me.ListBox1.RowSource = "BUAH!B2:B" & lastRow2
refreshTables
kodeOtomatis
kosongkan
Me.TB_Nama.SetFocus
End Sub

Private Sub CommandButton1_Click()
Me.TB_Nama = ""
For i = 0 To ListBox1.ListCount - 1
    Me.ListBox1.Selected(i) = False
Next i

For i = 0 To ListBox2.ListCount - 1
    Me.ListBox2.Selected(i) = False
Next i
kodeOtomatis
Me.TB_Nama.SetFocus
End Sub




Private Sub UpdateHobbies()
Dim SelectedRows()
idx = 0
ReDim SelectedRows(1 To 1)
With Me.ListBox1
  For i = 0 To .ListCount - 1
    If .Selected(i) Then
      idx = idx + 1
      ReDim Preserve SelectedRows(1 To idx)
      SelectedRows(idx) = i
    End If
  Next i
  Set yyy = Range(.RowSource)
  textcheckdate = hobbies
  For Each SelectedRow In SelectedRows
   yyy.Cells(SelectedRow + 1, 2).Value = textcheckdate
  Next
End With
End Sub

Function cariHobi() As Variant
Dim myArr As Variant
For x = 1 To Me.ListBox1.ListCount - 1
    If Me.ListBox1.Selected(x) Then
        MsgBox Me.ListBox1.List(x, 0)
        'MsgBox cariHobi
    End If
Next x
End Function

Private Sub ListBox1_Change()
'MsgBox hobbies
'hobbies
'GetHobbies
'MsgBox myHobbies
End Sub

Private Sub ListBox1_Click()


End Sub

Private Sub ListBox2_Click()
Dim ws As Worksheet
Set ws = Sheets("BUAH")
ws.Activate
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ListBox1.RowSource = ""
ListBox1.RowSource = "B2:B" & lastRow
Me.TB_ID = Me.ListBox2.Column(0)
Me.TB_Nama = Me.ListBox2.Column(1)
selectHobbies
'Me.cmdSimpan.Caption = "UPDATE"
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim ws As Worksheet
Set ws = Sheets("BUAH")
ws.Activate
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ListBox1.RowSource = ""
ListBox1.RowSource = "B2:B" & lastRow
Me.TB_ID = Me.ListBox2.Column(0)
Me.TB_Nama = Me.ListBox2.Column(1)
selectHobbies
End Sub

Private Sub UserForm_Activate()
kodeOtomatis
End Sub

Private Sub UserForm_Initialize()


End Sub

Private Sub UserForm_Layout()
refreshTables
End Sub
