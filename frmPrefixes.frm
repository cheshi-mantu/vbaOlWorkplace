VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrefixes 
   Caption         =   "choose the prefix"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "frmPrefixes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrefixes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
lstBLPrefixes.Value = ""
frmPrefixes.Hide
End Sub

Private Sub lstBLPrefixes_Change()
    frmPrefixes.Hide
End Sub
Private Sub UserForm_Initialize()
Dim arrBLPefixes
arrBLPefixes = Array("STATUS", "MOM", "REQ", "QUESTION", "APPROVAL", "SITE VISIT")
    For i = 0 To UBound(arrBLPefixes)
        lstBLPrefixes.AddItem (arrBLPefixes(i))
    Next i
    
End Sub

