VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SAP_Password 
   Caption         =   "SAP Password"
   ClientHeight    =   1395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5220
   OleObjectBlob   =   "SAP_Password.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SAP_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnCancel_Click()
    SAP_Password.Hide
    End
End Sub
Private Sub txtPassword_AfterUpdate()
    Password = Me.txtPassword
    If Password <> "" And UserName <> "" Then SAP_Password.Hide
End Sub
Private Sub txtUserName_AfterUpdate()
    UserName = Me.txtUserName
    If Password <> "" And UserName <> "" Then SAP_Password.Hide
End Sub

