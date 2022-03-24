VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BatchEditForm 
   Caption         =   "Batch edit"
   ClientHeight    =   4301
   ClientLeft      =   44
   ClientTop       =   374
   ClientWidth     =   4708
   OleObjectBlob   =   "BatchEditForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BatchEditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' -- BatchEditForm.frm
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
    LoadSettings
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
End Sub

Private Sub LoadSettings()
    TextBoxTempFolder.Text = GetTempPath(False)
    ComboBoxOutputFormat.List = ArrayFromCSVList(OUTPUT_FORMATS)
    ComboBoxOutputFormat.ListIndex = GetArrayIndex(OUTPUT_FORMATS, GetValue(OUTPUT_FORMAT_VALUE_NAME))
    TextBoxDPI.Text = GetValue(BITMAP_DPI_VALUE_NAME)
    checkboxPNGTransparent.value = GetValue(PNG_TRANSPARENT_VALUE_NAME)
    CheckBoxUseCairo.value = GetValue(USE_CAIRO_VALUE_NAME)

    CheckBoxModifyTempFolder.value = False
    CheckBoxModifyOutputFormat.value = False
    CheckBoxModifyDPI.value = False
    CheckBoxModifyUseCairo.value = False
    CheckBoxModifyPNGTransparent.value = False
    CheckBoxReplaceText.value = False
    
    Apply_CheckBoxModifyTempFolder
    Apply_CheckBoxModifyOutputFormat
    Apply_CheckBoxModifyDPI
    Apply_CheckBoxModifyUseCairo
    Apply_CheckBoxModifyPNGTransparent
    Apply_CheckBoxReplaceText
End Sub

Sub ButtonRun_Click()
    BatchEditForm.Hide
    Call RegenerateSelectedDisplays
    Unload BatchEditForm
End Sub

Private Sub ButtonCancel_Click()
    Unload BatchEditForm
End Sub

' Enable/Disable Modifications

Private Sub CheckBoxModifyTempFolder_Click()
    Apply_CheckBoxModifyTempFolder
End Sub

Private Sub CheckBoxModifyOutputFormat_Click()
    Apply_CheckBoxModifyOutputFormat
End Sub

Private Sub CheckBoxModifyDPI_Click()
    Apply_CheckBoxModifyDPI
End Sub

Private Sub CheckBoxModifyUseCairo_Click()
    Apply_CheckBoxModifyUseCairo
End Sub

Private Sub CheckBoxModifyPNGTransparent_Click()
    Apply_CheckBoxModifyPNGTransparent
End Sub

Private Sub CheckBoxReplaceText_Click()
    Apply_CheckBoxReplaceText
End Sub

Private Sub Apply_CheckBoxModifyTempFolder()
    LabelTempFolder.Enabled = CheckBoxModifyTempFolder.value
    TextBoxTempFolder.Enabled = CheckBoxModifyTempFolder.value
End Sub

Private Sub Apply_CheckBoxModifyOutputFormat()
    ComboBoxOutputFormat.Enabled = CheckBoxModifyOutputFormat.value
    LabelOutputFormat.Enabled = CheckBoxModifyOutputFormat.value
End Sub

Private Sub Apply_CheckBoxModifyDPI()
    LabelDPI.Enabled = CheckBoxModifyDPI.value
    TextBoxDPI.Enabled = CheckBoxModifyDPI.value
End Sub

Private Sub Apply_CheckBoxModifyUseCairo()
    CheckBoxUseCairo.Enabled = CheckBoxModifyUseCairo.value
End Sub

Private Sub Apply_CheckBoxModifyPNGTransparent()
    checkboxPNGTransparent.Enabled = CheckBoxModifyPNGTransparent.value
End Sub

Private Sub Apply_CheckBoxReplaceText()
    LabelReplace.Enabled = CheckBoxReplaceText.value
    TextBoxFindText.Enabled = CheckBoxReplaceText.value
    LabelWith.Enabled = CheckBoxReplaceText.value
    TextBoxReplacementText.Enabled = CheckBoxReplaceText.value
End Sub


