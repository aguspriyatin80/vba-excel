VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGallery 
   Caption         =   "FORM GALLERY"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8310.001
   OleObjectBlob   =   "frmGallery.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGallery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FilePath() As String
Dim MyFileName As String
Dim MyFileNameWOExt As String
Dim FileFold As String
Dim TotPhoto As Long

Sub TampilkanCekList()
'On Error GoTo Salah
On Error Resume Next
If FileFold = Empty Then
    Exit Sub
End If
'If Me.SelPhoto.Caption = "SELECT MENU" Then
    'Me.SelPhoto.Caption = "DESELECT MENU"
    'BersihkanCekList
    'For ChkBtn = 1 To TotPhoto
     '   Set ChkBox = Me.Controls.Add("Forms.CheckBox.1")
      '      ChkBox.Name = "CheckBox" & ChkBtn
       '     ChkBox.Width = Me.Controls("Image" & ChkBtn).Width
        '    ChkBox.Left = Me.Controls("Image" & ChkBtn).Left
         '   ChkBox.Top = Me.Controls("Image" & ChkBtn).Top + Me("Image" & ChkBtn).Height - 10
          '  frmGallery.Controls("CheckBox" & ChkBtn).BackColor = vbRed
           ' frmGallery.Controls("CheckBox" & ChkBtn).BackStyle = 1
            'ControlName = frmGallery.Controls("CheckBox" & ChkBtn).Name & " - " & frmGallery.Controls("Image" & ChkBtn).Name
            'MsgBox ControlName
    'Next ChkBtn
'Else
    'Me.SelPhoto.Caption = "SELECT MENU"
    For ChkBtn = 1 To Me.Controls.Count
        Me.Controls.Remove ("CheckBox" & ChkBtn)
    Next ChkBtn
'End If
End Sub

Sub BersihkanCekList()
For ChkBtn = 0 To TotPhoto - 1
    Me.Controls.Remove ("CheckBox" & ChkBtn)
Next ChkBtn
End Sub
Private Sub InsertImgFood()
TotPhoto = 0
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
FileFold = "FOOD"
Set ObjFolder = ObjFSO.getFolder(FileFold)
'finding total no of image file
For Each ObjFiles In ObjFolder.Files
    If ObjFiles.Type = "JPG File" Then
        TotPhoto = TotPhoto + 1
    End If
Next
ReDim FilePath(TotPhoto)
FileArr = 0
For Each ObjFiles In ObjFolder.Files
    If ObjFiles.Type = "JPG File" Then
        FilePath(FileArr) = ObjFiles.Path
        FileArr = FileArr + 1
    End If
Next
'Insert Image
LeftPos = 10
TopPos = 100
    For InsImg = 1 To TotPhoto
        Set Img = frmGallery.Controls.Add("Forms.Image.1")
            With Img
                .Left = LeftPos
                .Top = TopPos
                .Picture = LoadPicture(FilePath(InsImg - 1))
                .PictureSizeMode = 1
                LeftPos = LeftPos + .Width + 5
                Pic = Pic + 1
            End With
            If Pic = 5 Then
                TopPos = TopPos + Img.Height + 5
                LeftPos = 10
                Pic = 0
            End If
            If Img.Top + Img.Height > 300 Then
                Me.ScrollBars = fmScrollBarsVertical
                Me.ScrollHeight = Img.Top + Img.Height + 5
            Else
                Me.ScrollBars = fmScrollBarsNone
            End If
    MyFilePath = FilePath(InsImg - 1)
    MyFileName = ObjFSO.GetFileName(MyFilePath)
    MyFileNameWOExt = Left(MyFileName, InStr(MyFileName, ".") - 1)
    Me.ComboBox1.AddItem MyFileNameWOExt
    Next InsImg
End Sub
Private Sub InsertImgBeverage()
TotPhoto = 0
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
FileFold = "BEVERAGE"
Set ObjFolder = ObjFSO.getFolder(FileFold)
'finding total no of image file
For Each ObjFiles In ObjFolder.Files
    If ObjFiles.Type = "JPG File" Then
        TotPhoto = TotPhoto + 1
    End If
Next
ReDim FilePath(TotPhoto)
FileArr = 0
For Each ObjFiles In ObjFolder.Files
    If ObjFiles.Type = "JPG File" Then
        FilePath(FileArr) = ObjFiles.Path
        FileArr = FileArr + 1
    End If
Next
'Insert Image
LeftPos = 10
TopPos = 100
    For InsImg = 1 To TotPhoto
        Set Img = frmGallery.Controls.Add("Forms.Image.1")
            With Img
                .Left = LeftPos
                .Top = TopPos
                .Picture = LoadPicture(FilePath(InsImg - 1))
                .PictureSizeMode = 1
                LeftPos = LeftPos + .Width + 5
                Pic = Pic + 1
            End With
            If Pic = 5 Then
                TopPos = TopPos + Img.Height + 5
                LeftPos = 10
                Pic = 0
            End If
            If Img.Top + Img.Height > 300 Then
                Me.ScrollBars = fmScrollBarsVertical
                Me.ScrollHeight = Img.Top + Img.Height + 5
            Else
                Me.ScrollBars = fmScrollBarsNone
            End If
    MyFilePath = FilePath(InsImg - 1)
    MyFileName = ObjFSO.GetFileName(MyFilePath)
    'Me.ComboBox1.AddItem MyFileName
    MyFileNameWOExt = Left(MyFileName, InStr(MyFileName, ".") - 1)
    Me.ComboBox1.AddItem MyFileNameWOExt
    Next InsImg
End Sub
Private Sub InsertImg()
TotPhoto = 0
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
FileFold = "ALL"
Set ObjFolder = ObjFSO.getFolder(FileFold)
'finding total no of image file
For Each ObjFiles In ObjFolder.Files
    If ObjFiles.Type = "JPG File" Then
        TotPhoto = TotPhoto + 1
    End If
Next
ReDim FilePath(TotPhoto)
FileArr = 0
For Each ObjFiles In ObjFolder.Files
    If ObjFiles.Type = "JPG File" Then
        FilePath(FileArr) = ObjFiles.Path
        FileArr = FileArr + 1
    End If
Next
'Insert Image
LeftPos = 10
TopPos = 100
    For InsImg = 1 To TotPhoto
        Set Img = frmGallery.Controls.Add("Forms.Image.1")
            Img.BackColor = vbRed
            
            With Img
                .Name = "Image" & InsImg
                .Left = LeftPos
                .Top = TopPos
                .Picture = LoadPicture(FilePath(InsImg - 1))
                .PictureSizeMode = 1
                
                LeftPos = LeftPos + .Width + 5
                Pic = Pic + 1
            End With
            If Pic = 5 Then
                TopPos = TopPos + Img.Height + 5
                LeftPos = 10
                Pic = 0
            End If
            If Img.Top + Img.Height > 300 Then
                Me.ScrollBars = fmScrollBarsVertical
                Me.ScrollHeight = Img.Top + Img.Height + 5
            Else
                Me.ScrollBars = fmScrollBarsNone
            End If
    MyFilePath = FilePath(InsImg - 1)
    MyFileName = ObjFSO.GetFileName(MyFilePath)
    'Me.ComboBox1.AddItem MyFileName
    MyFileNameWOExt = Left(MyFileName, InStr(MyFileName, ".") - 1)
    Me.ComboBox1.AddItem MyFileNameWOExt
    Next InsImg
End Sub
Private Sub chkOrder_Click()
If Me.chkOrder.Value = True Then
    Me.txtQty.Enabled = True
    Me.txtQty.Value = 1
    Me.cmdPlus.Enabled = True
    Me.cmdProses.Visible = True
Else
    Me.cmdMinus.Enabled = False
    Me.txtQty.Enabled = False
    Me.txtQty.Value = 0
    Me.cmdPlus.Enabled = False
    Me.cmdProses.Visible = False
End If
End Sub

Private Sub cmdBackToMenu_Click()
SelFolder_Click

End Sub

Private Sub cmdMinus_Click()
If Me.txtQty.Value > 1 Then
    Me.txtQty.Value = Me.txtQty - 1
Else
    Me.cmdMinus.Enabled = False
End If
End Sub
Private Sub cmdPlus_Click()
Me.cmdMinus.Enabled = True
Me.txtQty.Value = Me.txtQty + 1
End Sub
Private Sub cmdProses_Click()
MsgBox Me.ComboBox1.Value & " sudah masuk dalam pesanan "
SelFolder_Click
End Sub
Private Sub ComboBox1_Change()
Dim imagePath As String, imageName As String
imagePath = ThisWorkbook.Path & "\ALL\"
imageName = Me.ComboBox1.Value & ".jpg"
On Error Resume Next
For RemImg = 0 To TotPhoto
    Me.Controls.Remove ("Image" & RemImg)
    Me.Controls.Remove ("CheckBox" & RemImg)
Next RemImg
Me.ImgMenu.Visible = True
Me.cmdMinus.Visible = True
Me.txtQty.Visible = True
Me.cmdPlus.Visible = True

Me.cmdMinus.Enabled = True
Me.txtQty.Enabled = True
Me.cmdPlus.Enabled = True


Me.cmdProses.Visible = True
Me.cmdBackToMenu.Visible = True
Me.ImgMenu.Picture = LoadPicture(imagePath & imageName)
Me.ScrollBars = fmScrollBarsNone
Me.SelPhoto.Caption = "SELECT MENU"
If Me.ComboBox1.Value = "" Then
    Me.SelPhoto.Enabled = True
Else
    Me.SelPhoto.Enabled = False
End If
End Sub

Private Sub DelBtn_Click()
'TotPhoto = Me.Controls("Forms.Image.1").Count
On Error Resume Next
Dim valsum As Integer
sumTrue = 0
  
If TotPhoto > 0 Then
    
    For ChkBtn = 0 To TotPhoto
        sumTrue = sumTrue + IIf(Me("CheckBox" & ChkBtn).Value = True, 1, 0)
        If Me("CheckBox" & ChkBtn).Value = True Then
            Me.Controls.Remove ("CheckBox" & ChkBtn)
            Me.Controls.Remove ("Image" & ChkBtn)
            Kill FilePath(ChkBtn - 1)
        End If
    Next ChkBtn
    If sumTrue = 0 Then
        MsgBox "Foto belum dipilih"
    End If
End If
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
myString = ObjFSO.getFolder(FileFold)
If ContainsSubString("FOOD", myString) = True Then
    If TotPhoto <= 1 Then
        Exit Sub
    ElseIf TotPhoto > 1 And sumTrue > 1 Then
        Exit Sub
    Else
        SelFood_Click
    End If
ElseIf ContainsSubString("BEVERAGE", myString) = True Then
    If TotPhoto <= 1 Then
        Exit Sub
    ElseIf TotPhoto > 1 And sumTrue > 1 Then
        Exit Sub
    Else
        SelBeverage_Click
    End If
    
ElseIf ContainsSubString("ALL", myString) = True Then
    If TotPhoto <= 1 Then
        Exit Sub
    ElseIf TotPhoto > 1 And sumTrue > 1 Then
        Exit Sub
    Else
        SelFolder_Click
    End If
End If

End Sub

Private Sub SelBeverage_Click()
'Set Filedig = Application.FileDialog(msoFileDialogFolderPicker)
'With Filedig
'    .Title = "Select folder for image"
'    If .Show <> -1 Then GoTo NoFolder
'    FileFold = .SelectedItems(1)
'End With
'remove remove existing controls
On Error Resume Next
TotPhoto = Me.Controls.Count
If TotPhoto = 0 Then Exit Sub
For RemImg = 1 To TotPhoto
    Me.Controls.Remove ("Image" & RemImg)
Next RemImg
Me.ComboBox1.Clear
Me.ImgMenu.Visible = False
Me.txtQty.Visible = False
Me.cmdMinus.Visible = False
Me.cmdPlus.Visible = False
Me.cmdProses.Visible = False
BersihkanCekList
InsertImgBeverage
Me.SelPhoto.Caption = "SELECT MENU"
For ChkBtn = 1 To Me.Controls.Count
    Me.Controls.Remove ("CheckBox" & ChkBtn)
Next ChkBtn
SelPhoto_Click
NoFolder:

End Sub

Private Sub SelFolder_Click()
'Set Filedig = Application.FileDialog(msoFileDialogFolderPicker)
'With Filedig
'    .Title = "Select folder for image"
'    If .Show <> -1 Then GoTo NoFolder
'    FileFold = .SelectedItems(1)
'End With
'remove remove existing controls
On Error Resume Next
TotPhoto = Me.Controls.Count
For RemImg = 1 To TotPhoto
    Me.Controls.Remove ("Image" & RemImg)
Next RemImg
Me.ComboBox1.Clear
Me.ImgMenu.Visible = False
'Me.chkOrder.Visible = False
Me.txtQty.Visible = False
Me.cmdMinus.Visible = False
Me.cmdPlus.Visible = False
Me.cmdProses.Visible = False
Me.cmdBackToMenu.Visible = False
InsertImg
Me.SelPhoto.Caption = "SELECT MENU"
For ChkBtn = 1 To Me.Controls.Count
    Me.Controls.Remove ("CheckBox" & ChkBtn)
Next ChkBtn
SelPhoto_Click
NoFolder:

End Sub

Private Sub SelFood_Click()
'Set Filedig = Application.FileDialog(msoFileDialogFolderPicker)
'With Filedig
'    .Title = "Select folder for image"
'    If .Show <> -1 Then GoTo NoFolder
'    FileFold = .SelectedItems(1)
'End With
'remove remove existing controls
On Error Resume Next
TotPhoto = Me.Controls.Count
For RemImg = 1 To TotPhoto
    Me.Controls.Remove ("Image" & RemImg)
Next RemImg
Me.ComboBox1.Clear
Me.ImgMenu.Visible = False
Me.txtQty.Visible = False
Me.cmdMinus.Visible = False
Me.cmdPlus.Visible = False
Me.cmdProses.Visible = False
BersihkanCekList
InsertImgFood
Me.SelPhoto.Caption = "SELECT MENU"
For ChkBtn = 1 To Me.Controls.Count
    Me.Controls.Remove ("CheckBox" & ChkBtn)
Next ChkBtn
SelPhoto_Click
NoFolder:

End Sub

Private Sub SelPhoto_Click()
'On Error GoTo Salah
On Error Resume Next
'If FileFold = Empty Then
'    Exit Sub
'End If

    If Me.SelPhoto.Caption = "SELECT MENU" Then
        Me.SelPhoto.Caption = "DESELECT MENU"
        BersihkanCekList
        If TotPhoto = 0 Then
        MsgBox "Belum ada data"
        Exit Sub
        End If
        
        For ChkBtn = 1 To TotPhoto
            Set ChkBox = Me.Controls.Add("Forms.CheckBox.1")
                ChkBox.Name = "CheckBox" & ChkBtn
                ChkBox.Width = Me.Controls("Image" & ChkBtn).Width
                ChkBox.Height = Me.Controls("Image" & ChkBtn).Height
                ChkBox.Left = Me.Controls("Image" & ChkBtn).Left
                ChkBox.BackStyle = 1
                ChkBox.BackColor = red
                'ChkBox.Top = Me.Controls("Image" & ChkBtn).Top + Me("Image" & ChkBtn).Height - 10
                ChkBox.Top = Me.Controls("Image" & ChkBtn).Top
                'frmGallery.Controls("CheckBox" & ChkBtn).BackColor = vbRed
                frmGallery.Controls("CheckBox" & ChkBtn).BackStyle = 0
                'ControlName = frmGallery.Controls("CheckBox" & ChkBtn).Name & " - " & frmGallery.Controls("Image" & ChkBtn).Name
                'MsgBox ControlName
        Next ChkBtn
    Else
        Me.SelPhoto.Caption = "SELECT MENU"
        For ChkBtn = 1 To Me.Controls.Count
            Me.Controls.Remove ("CheckBox" & ChkBtn)
        Next ChkBtn
    End If
End Sub
Private Sub UserForm_Initialize()
'Set Filedig = Application.FileDialog(msoFileDialogFolderPicker)
'With Filedig
'    .Title = "Select folder for image"
'    If .Show <> -1 Then GoTo NoFolder
'    FileFold = .SelectedItems(1)
'End With
'remove remove existing controls
On Error Resume Next
TotPhoto = Me.Controls.Count
For RemImg = 1 To TotPhoto
    Me.Controls.Remove ("Image" & RemImg)
Next RemImg
Me.ComboBox1.Clear
InsertImg
Me.ImgMenu.Visible = False
'Me.chkOrder.Visible = False
Me.txtQty.Visible = False
Me.txtQty.Value = 0
Me.cmdMinus.Visible = False
Me.cmdPlus.Visible = False
Me.cmdBackToMenu.Visible = False
Me.txtQty.Enabled = False
Me.cmdMinus.Enabled = False
Me.cmdPlus.Enabled = False
Me.cmdProses.Visible = False
SelPhoto_Click
NoFolder:
End Sub
