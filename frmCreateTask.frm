VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateTask 
   Caption         =   "Create task"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   OleObjectBlob   =   "frmCreateTask.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
frmCreateTask.Hide
End Sub

Private Sub UserForm_Initialize()
'initialize form controls
Tools.subSetFolders
    For i = 0 To UBound(Tools.arrDomains)
        cmbCountryList.AddItem (Tools.arrDomains(i))
    Next i
    
    For i = 0 To UBound(Tools.arrCities)
        cmbCity.AddItem (Tools.arrCities(i))
    Next i
    txbxDescription.Value = "Requestor: " & vbCrLf & "Reason: " & vbCrLf & "Priority:"
    
End Sub
