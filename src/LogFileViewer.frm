VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LogFileViewer 
   Caption         =   "Error in GLE Code"
   ClientHeight    =   6974
   ClientLeft      =   44
   ClientTop       =   330
   ClientWidth     =   8855.001
   OleObjectBlob   =   "LogFileViewer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LogFileViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



































Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
End Sub

Private Sub CloseLogButton_Click()
    
    SelStartPos = GLEForm.TextBoxGLECode.SelStart
    TempPath = GLEForm.TextBoxTempFolder.Text
    
    If Left(TempPath, 1) = "." Then
        Dim sPath As String
        sPath = ActivePresentation.path
        If Len(sPath) > 0 Then
            If Right(sPath, 1) <> "\" Then
                sPath = sPath & "\"
            End If
            TempPath = sPath & TempPath
        Else
            MsgBox "You need to have saved your presentation once to use a relative path."
            Exit Sub
        End If
    End If
    
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
   ' objStream.LoadFromFile (TempPath & GetFilePrefix() & ".tex")
    GLEForm.TextBoxGLECode.Text = objStream.ReadText()

    CloseLogButton.Caption = "Close"
    Unload LogFileViewer
    GLEForm.TextBoxGLECode.SetFocus
    If SelStartPos < Len(GLEForm.TextBoxGLECode.Text) Then
        GLEForm.TextBoxGLECode.SelStart = SelStartPos
    End If
End Sub

Private Sub CmdButtonExternalEditor_Click()
    TempPath = GLEForm.TextBoxTempFolder.Text
    If Left(TempPath, 1) = "." Then
        Dim sPath As String
        sPath = ActivePresentation.path
        If Len(sPath) > 0 Then
            If Right(sPath, 1) <> "\" Then
                sPath = sPath & "\"
            End If
            TempPath = sPath & TempPath
        Else
            MsgBox "You need to have saved your presentation once to use a relative path."
            Exit Sub
        End If
    End If
   ' LogFileViewer.Caption = """" & GetEditorPath() & """ """ & TempPath & GetFilePrefix() & ".tex"""
   ' CloseLogButton.Caption = "Reload modified code"
   ' Shell """" & GetEditorPath() & """ """ & TempPath & GetFilePrefix() & ".tex""", vbNormalFocus
End Sub
