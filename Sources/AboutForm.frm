VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AboutForm 
   Caption         =   "About"
   ClientHeight    =   3960
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7050
   OleObjectBlob   =   "AboutForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CloseAbout_Click()
    Me.Hide
    Unload Me
End Sub
