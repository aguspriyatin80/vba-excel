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
Dim MyFilePath As String
Dim MyFileName As String
Dim MyFileNameWOExt As String
Dim FileFold As String
Dim TotPhoto As Long

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
    'Me.ComboBox1.AddItem MyFileName
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

Private Sub CheckBox1_Click()
If Me.CheckBox1.Value = True Then
    Me.txtQty.Enabled = True
    Me.txtQty.Value = 1
    Me.cmdPlus.Enabled = True
Else
    Me.cmdMinus.Enabled = False
    Me.txtQty.Enabled = False
    Me.txtQty.Value = 0
    Me.cmdPlus.Enabled = False
End If
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

Private Sub ComboBox1_Change()
Dim imagePath As String, imageName As String
imagePath = "C:\Users\agus\Documents\ALL\"
imageName = Me.ComboBox1.Value & ".jpg"

On Error Resume Next
TotPhoto = Me.Controls.Count
For RemImg = 0 To TotPhoto
    Me.Controls.Remove ("Image" & RemImg)
Next RemImg
Me.Image1.Visible = True
Me.CheckBox1.Visible = True
Me.cmdMinus.Visible = True
Me.txtQty.Visible = True
Me.cmdPlus.Visible = True
Me.Image1.Picture = LoadPicture(imagePath & imageName)
Me.ScrollBars = fmScrollBarsNone
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

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
For RemImg = 0 To TotPhoto
    Me.Controls.Remove ("Image" & RemImg)
Next RemImg
Me.ComboBox1.Clear
Me.Image1.Visible = False
InsertImgBeverage
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
For RemImg = 0 To TotPhoto
    Me.Controls.Remove ("Image" & RemImg)
Next RemImg
Me.ComboBox1.Clear
Me.Image1.Visible = False
Me.CheckBox1.Visible = False
Me.txtQty.Visible = False
Me.cmdMinus.Visible = False
Me.cmdPlus.Visible = False
InsertImg
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
For RemImg = 0 To TotPhoto
    Me.Controls.Remove ("Image" & RemImg)
Next RemImg
Me.ComboBox1.Clear
Me.Image1.Visible = False
Me.CheckBox1.Visible = False
Me.txtQty.Visible = False
Me.cmdMinus.Visible = False
Me.cmdPlus.Visible = False
InsertImgFood
NoFolder:
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
For RemImg = 0 To TotPhoto
    Me.Controls.Remove ("Image" & RemImg)
Next RemImg
Me.ComboBox1.Clear
InsertImg
Me.Image1.Visible = False
Me.CheckBox1.Visible = False
Me.txtQty.Visible = False
Me.txtQty.Value = 0
Me.cmdMinus.Visible = False
Me.cmdPlus.Visible = False

Me.txtQty.Enabled = False
Me.cmdMinus.Enabled = False
Me.cmdPlus.Enabled = False
NoFolder:
End Sub



