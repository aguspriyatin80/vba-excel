VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CommandButtonWithIcon 
   Caption         =   "UserForm1"
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7410
   OleObjectBlob   =   "CommandButtonWithIcon.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CommandButtonWithIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub TambahSpasi2(cmd As MSForms.CommandButton)
Dim str As String
str = cmd.Caption
Do While Len(str) < 17
    str = str & Chr(160)
Loop
cmd.Caption = Chr(160) & str
str = ""
End Sub

Sub TambahSpasi()
Dim str As String
Dim wk As MSForms.CommandButton
    str = Me.CommandButton2.Caption
    str2 = Me.CommandButton4.Caption

    Do While Len(str) < 17
    str = str & Chr(160)
    Loop
    
    Do While Len(str2) < 17
        str2 = str2 & Chr(160)
    Loop

    Me.CommandButton2.Caption = Chr(160) & str
    Me.CommandButton4.Caption = Chr(160) & str2
    str = ""
    str2 = ""
End Sub


Private Sub CommandButton3_Click()

End Sub

Private Sub CommandButton2_Click()
'TambahSpasi2 Me.CommandButton2
If CommandButton2.Caption = " LAPORAN" Then
    Me.CommandButton2.Caption = " LAPORAN PENJUALAN"
Else
    Me.CommandButton2.Caption = "LAPORAN"
End If

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
'TambahSpasi2 Me.CommandButton2
'TambahSpasi2 Me.CommandButton4
End Sub
