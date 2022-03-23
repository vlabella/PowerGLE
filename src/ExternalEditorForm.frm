VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExternalEditorForm 
   Caption         =   "External Editor"
   ClientHeight    =   2640
   ClientLeft      =   11
   ClientTop       =   330
   ClientWidth     =   5687
   OleObjectBlob   =   "ExternalEditorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExternalEditorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

















Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
End Sub

Private Sub CmdButtonCancel_Click()
    Unload ExternalEditorForm
End Sub

Private Sub LoadTextIntoGLEForm()
    Dim OutputPath As String
    Dim Filename As String
    
    OutputPath = GetTempPath()
    CreateFolder (OutputPath) ' make sure it exists
    If OutputPath = "" Then
        Exit Sub
    End If
    Dim FigureName As String
    ' this get populated upon init or user changes it
    FigureName = GLEForm.TextBoxFigureName.value
    OutputPath = AddSlash(AddSlash(OutputPath) + FigureName)
    ' does folder exist? need to warn user for new figures only
    ' probably combine this with above folder seletion
    CreateFolder (OutputPath)
    Filename = OutputPath + FigureName + "." + GLE_EXT
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(Filename) Then
        Set DataFile = fs.OpenTextFile(Filename, ForReading)
        GLEForm.TextBoxGLECode.Text = DataFile.ReadAll
        DataFile.Close
    End If
End Sub

Private Sub CmdButtonReload_Click()
    SelStartPos = GLEForm.TextBoxGLECode.SelStart
    Call LoadTextIntoGLEForm
    Unload ExternalEditorForm
    GLEForm.Hide
    GLEForm.Show vbModal
    GLEForm.TextBoxGLECode.SetFocus
    If SelStartPos < Len(GLEForm.TextBoxGLECode.Text) Then
        GLEForm.TextBoxGLECode.SelStart = SelStartPos
    End If
End Sub

Private Sub CmdButtonGenerate_Click()
    SelStartPos = GLEForm.TextBoxGLECode.SelStart
    Call LoadTextIntoGLEForm
    Unload ExternalEditorForm
    GLEForm.TextBoxGLECode.SetFocus
    If SelStartPos < Len(GLEForm.TextBoxGLECode.Text) Then
        GLEForm.TextBoxGLECode.SelStart = SelStartPos
    End If
    Call GLEForm.ButtonGenerate_Click
End Sub
