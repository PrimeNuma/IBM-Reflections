VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CDZ_IMPA 
   Caption         =   "CDZA IMPA (DEV V1.02)"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12810
   OleObjectBlob   =   "CDZ_IMPA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CDZ_IMPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CPTbtn_Click()

Call CPT

End Sub

Private Sub EXWbtn_Click()

Call EXW

End Sub

Private Sub OFTbtn_Click()

Call Get_OFT

End Sub

Private Sub PHBbtn_Click()

Call PHB

End Sub

Private Sub Prepaidbtn_Click()

MsgBox "Discussion to be had in regards to what is in and what is out"

End Sub

Private Sub printINVbtn_Click()

MsgBox "Under Construction"

End Sub

Private Sub Collectbtn_Click()

Call Collectfreight

End Sub

Private Sub UserForm_Click()

End Sub


Private Sub IRMbtn_Click()

Call IRM
'MsgBox "Under Construction"

End Sub
