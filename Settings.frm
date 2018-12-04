VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings 
   Caption         =   "Settings"
   ClientHeight    =   3996
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   2964
   OleObjectBlob   =   "Settings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Sheets(2).Cells(1, 1).Value = InputBox("Enter length", "Board length")
Sheets(2).Cells(2, 1).Value = InputBox("Enter height", "Board height")
Settings.Hide

End Sub

Private Sub CommandButton2_Click()
xColors.Show
Settings.Hide

End Sub

Private Sub CommandButton3_Click()
Sheets(2).Cells(5, 1).Value = InputBox("Minimum neighbours", "Min")
Sheets(2).Cells(6, 1).Value = InputBox("Maximum neighbours", "Max")
Settings.Hide
End Sub

Private Sub CommandButton4_Click()
Sheets(2).Cells(3, 1).Value = InputBox("Number of rounds", "Rounds")
Settings.Hide
End Sub
