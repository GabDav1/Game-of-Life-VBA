VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} xColors 
   Caption         =   "Color Schema"
   ClientHeight    =   1476
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   5832
   OleObjectBlob   =   "xColors.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "xColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Sheets(2).Cells(4, 1) = 1
xColors.Hide

End Sub

Private Sub CommandButton2_Click()
Sheets(2).Cells(4, 1) = 2
xColors.Hide

End Sub

Private Sub CommandButton3_Click()
Sheets(2).Cells(4, 1) = 3
xColors.Hide

End Sub
