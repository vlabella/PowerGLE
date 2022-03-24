VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm 
   Caption         =   "PowerGLE Settings"
   ClientHeight    =   4928
   ClientLeft      =   11
   ClientTop       =   330
   ClientWidth     =   6292
   OleObjectBlob   =   "SettingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' -- SettingsForm.frm
'
' Form that handles changing and viewing of global settings
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
Private Sub ButtonAbsTempPath_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker) 'msoFileDialogFilePicker
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.InitialFileName = AbsPathTextBox.Text

    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            AbsPathTextBox.Text = vrtSelectedItem
        Next vrtSelectedItem
    End If
    Set fd = Nothing
End Sub

Private Sub ButtonCancelTemp_Click()
    Unload SettingsForm
End Sub

Private Sub ButtonEditorPath_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker) 'msoFileDialogFolderPicker
    
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.InitialFileName = TextBoxExternalEditor.Text
    fd.Filters.Clear
    fd.Filters.Add "All Files", "*.*", 1
    
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            TextBoxExternalEditor.Text = vrtSelectedItem
        Next vrtSelectedItem
    End If

    Set fd = Nothing
    TextBoxExternalEditor.SetFocus
End Sub

Private Sub ButtonGLEPath_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker) 'msoFileDialogFolderPicker
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = False
    fd.InitialFileName = TextBoxGLE.Text
    fd.Filters.Clear
    fd.Filters.Add "All Files", "*.*", 1
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            TextBoxGLE.Text = vrtSelectedItem
        Next vrtSelectedItem
    End If
    Set fd = Nothing
    TextBoxGLE.SetFocus
End Sub

Private Sub ButtonReset_Click()
    AbsPathButton.value = GetDefaultValue(USE_ABSOLUTE_TEMP_DIR_VALUE_NAME)
    AbsPathTextBox.Text = GetDefaultValue(ABSOLUTE_TEMP_DIR_VALUE_NAME)
    RelPathTextBox.Text = GetDefaultValue(RELATIVE_TEMP_DIR_VALUE_NAME)
    CheckBoxUTF8.value = GetDefaultValue(USE_UTF8_VALUE_NAME)
    CheckBoxExternalEditor.value = GetDefaultValue(USE_EXTERNAL_EDITOR_VALUE_NAME)
    ComboBoxOutputFormat.ListIndex = GetArrayIndex(OUTPUT_FORMATS, GetDefaultValue(OUTPUT_FORMAT_VALUE_NAME))
    TextBoxGLE.Text = GetDefaultValue(GLE_EXECUTABLE_VALUE_NAME)
    TextBoxDPI.Text = GetDefaultValue(BITMAP_DPI_VALUE_NAME)
    CheckBoxUseCairo.value = GetDefaultValue(USE_CAIRO_VALUE_NAME)
    CheckBoxPreserveTempFiles.value = GetDefaultValue(PRESERVE_TEMP_FILES_VALUE_NAME)
    TextBoxScalingGain.Text = GetDefaultValue(SCALING_GAIN_VALUE_NAME)
    TextBoxTimeOut.Text = GetDefaultValue(TIMEOUT_VALUE_NAME)
    TextBoxFontSize.Text = GetDefaultValue(EDITOR_FONT_SIZE_VALUE_NAME)
    SetAbsRelDependencies
End Sub

Private Sub ButtonOk_Click()
    SetValue USE_ABSOLUTE_TEMP_DIR_VALUE_NAME, AbsPathButton.value
    SetValue ABSOLUTE_TEMP_DIR_VALUE_NAME, AbsPathTextBox.Text
    ' Temp folder
    If Left(RelPathTextBox.Text, 2) = ".\" Then
        RelPathTextBox.Text = Mid(RelPathTextBox.Text, 3, Len(RelPathTextBox.Text) - 2)
    End If
    SetValue RELATIVE_TEMP_DIR_VALUE_NAME, RelPathTextBox.Text
    ' UTF8
    SetValue USE_UTF8_VALUE_NAME, CheckBoxUTF8.value
    ' GLE command
    SetValue GLE_EXECUTABLE_VALUE_NAME, UnQuote(CStr(TextBoxGLE.Text))
    ' Global dpi setting
    SetValue BITMAP_DPI_VALUE_NAME, CInt(TextBoxDPI.Text)
    SetValue USE_CAIRO_VALUE_NAME, CheckBoxUseCairo.value
    SetValue PRESERVE_TEMP_FILES_VALUE_NAME, CheckBoxPreserveTempFiles.value
    ' Path to External Editor
    SetValue EXTERNAL_EDITOR_EXECUTABLE_VALUE_NAME, UnQuote(TextBoxExternalEditor.Text)
    SetValue USE_EXTERNAL_EDITOR_VALUE_NAME, CheckBoxExternalEditor.value
    SetValue SCALING_GAIN_VALUE_NAME, TextBoxScalingGain.Text
    SetValue TIMEOUT_VALUE_NAME, TextBoxTimeOut.Text
    SetValue EDITOR_FONT_SIZE_VALUE_NAME, TextBoxFontSize.Text
    Unload SettingsForm
End Sub

Private Sub AbsPathButton_Click()
    AbsPathButton.value = True
    SetAbsRelDependencies
End Sub

Private Sub LabelDLGLE_Click()
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", GLE_URL)
    If (lSuccess = 0) Then
        MsgBox "Cannot open " & Link
    End If
End Sub

Private Sub LabelDLTextEditor_Click()
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", EXTERNAL_EDITOR_URL)
    If (lSuccess = 0) Then
        MsgBox "Cannot open " & Link
    End If
End Sub

Private Sub RelPathButton_Click()
    AbsPathButton.value = False
    SetAbsRelDependencies
End Sub

Private Sub SetAbsRelDependencies()
    RelPathButton.value = Not AbsPathButton.value
    AbsPathTextBox.Enabled = AbsPathButton.value
    RelPathTextBox.Enabled = RelPathButton.value
End Sub

Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    AbsPathTextBox.Text = GetValue(ABSOLUTE_TEMP_DIR_VALUE_NAME)
    RelPathTextBox.Text = GetValue(RELATIVE_TEMP_DIR_VALUE_NAME)
    AbsPathButton.value = GetValue(USE_ABSOLUTE_TEMP_DIR_VALUE_NAME)
    TextBoxGLE.Text = GetValue(GLE_EXECUTABLE_VALUE_NAME)
    TextBoxDPI.Text = GetValue(BITMAP_DPI_VALUE_NAME)
    CheckBoxUseCairo.value = GetValue(USE_CAIRO_VALUE_NAME)
    CheckBoxPreserveTempFiles.value = GetValue(PRESERVE_TEMP_FILES_VALUE_NAME)
    CheckBoxUTF8.value = GetValue(USE_UTF8_VALUE_NAME)
    TextBoxTimeOut.Text = GetValue(TIMEOUT_VALUE_NAME)
    TextBoxFontSize.Text = GetValue(EDITOR_FONT_SIZE_VALUE_NAME)
    TextBoxScalingGain.Text = GetValue(SCALING_GAIN_VALUE_NAME)
    TextBoxExternalEditor.Text = GetValue(EXTERNAL_EDITOR_EXECUTABLE_VALUE_NAME)
    CheckBoxExternalEditor.value = GetValue(USE_EXTERNAL_EDITOR_VALUE_NAME)
    ComboBoxOutputFormat.List = ArrayFromCSVList(OUTPUT_FORMATS)
    ComboBoxOutputFormat.ListIndex = GetArrayIndex(OUTPUT_FORMATS, GetValue(OUTPUT_FORMAT_VALUE_NAME))
    SetAbsRelDependencies
End Sub

