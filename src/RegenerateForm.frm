VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegenerateForm 
   Caption         =   "Regenerating"
   ClientHeight    =   1804
   ClientLeft      =   44
   ClientTop       =   374
   ClientWidth     =   3916
   OleObjectBlob   =   "RegenerateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegenerateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' -- RegenerateForm.frm
'
' Form that displays status of batch regeneration
'
' PowerGLE: PowerPoint Add-in for GLE
'
' Author:   Vincent LaBella
' Email:    vlabella@sunypoly.edu
' GLE:      glx.sourceforge.io github.com/vlabella/GLE
' PowerGLE: github.com/vlabella/PowerGLE
'
' Inspired by and based on code from IguanaTeX  www.jonathanleroux.org/software/iguanatex/
'
Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    
End Sub

Private Sub CommandButtonCancel_Click()
    'CheckBoxContinue.Value = False
    RegenerateContinue = False
    Unload RegenerateForm
    'End
End Sub
