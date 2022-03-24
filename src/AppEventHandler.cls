VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1
'
' -- AppEventHandler.cls
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
Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    ' happens when a selection is changed or copy and pasted
    ' Detect event where PowerGLE figure has been created using copy and paste
    Debug.Print "App_WindowSelectionChange "
    If Sel.Type = ppSelectionShapes Then
        Debug.Print "App_WindowSelectionChange " + CStr(Sel.Type) + " " + CStr(Sel.HasChildShapeRange)
        ' Get all the shapes in the selection that are PowerGLEShapes maybe multiple if in a group
        Dim SelShapes() As Variant
        SelShapes = GetAllShapesInShapeRange(Sel.ShapeRange)
        ' get all the shapes in the current presenation that are PowerGLEShapes
        If (Not SelShapes) = -1 Then
            ' empty array no powergle shapes exit
            Debug.Print "no sel shapes"
            Exit Sub
        End If
        Dim PresShapes() As Variant
        PresShapes = GetAllShapesInPresentation(ActivePresentation)
        If (Not PresShapes) = -1 Then
            ' empty array no powergle shapes in presentation
            Exit Sub
        End If
        ' find shapes that have same uuid but different names - this indictes a COPY event
        For Each SelShape In SelShapes
            Debug.Print "SELECTION: PowerGLE Shape " + SelShape.name + " " + SelShape.Tags(GetShapeTagName(TAG_FIGURE_NAME)) + " " + SelShape.Tags(GetShapeTagName(TAG_FIGURE_UUID))
            For Each PresShape In PresShapes
                Debug.Print "PRESENTATION: PowerGLE Shape " + SelShape.name + " " + SelShape.Tags(GetShapeTagName(TAG_FIGURE_NAME)) + " " + SelShape.Tags(GetShapeTagName(TAG_FIGURE_UUID))
                If PresShape.Tags(GetShapeTagName(TAG_FIGURE_UUID)) = SelShape.Tags(GetShapeTagName(TAG_FIGURE_UUID)) And PresShape.name <> SelShape.name Then
                    Debug.Print "COPY EVENT"
                    With SelShape.Tags
                    .Delete (GetShapeTagName(TAG_FIGURE_UUID))
                    .Add GetShapeTagName(TAG_FIGURE_UUID), GetUUID()
                    .Delete (GetShapeTagName(TAG_TEMP_FOLDER))
                    .Add GetShapeTagName(TAG_TEMP_FOLDER), GetTempPath(False)
                    ' gets the next auto-gen figure name availible.
                    ' copied figure name will revert back to auto-gen on overriding any figure name customization of source figure by the user
                    .Delete (GetShapeTagName(TAG_FIGURE_NAME))
                    Dim next_index As Long
                    ' must get max of next index from both folder and existsing shapes in presentation
                    next_index = max(GetNextFigureIndexInFolder(GetTempPath()), GetNextFigureIndexInShapes(PresShapes))
                    .Add GetShapeTagName(TAG_FIGURE_NAME), GetFigureName(next_index)
                    ' create the figure directory so another copy will get unique name
                    CreateFolder (AddSlash(GetTempPath()) + .Item(GetShapeTagName(TAG_FIGURE_NAME)))
                    End With
                    ' update alt text
                    FormAltTextandTitle SelShape
                End If
            Next
        Next
    End If
End Sub

Private Sub App_AfterDragDropOnSlide(ByVal Sld As slide, ByVal X As Single, ByVal Y As Single)
Debug.Print "App_AfterDragDropOnSlide"
End Sub

Private Sub App_AfterNewPresentation(ByVal Pres As Presentation)
Debug.Print "App_AfterNewPresentation"
End Sub

Private Sub App_AfterPresentationOpen(ByVal Pres As Presentation)
Debug.Print "App_AfterPresentationOpen"
End Sub

Private Sub App_AfterShapeSizeChange(ByVal shp As Shape)
Debug.Print "App_AfterShapeSizeChange"
End Sub

Private Sub App_ColorSchemeChanged(ByVal SldRange As SlideRange)
Debug.Print "App_ColorSchemeChanged"
End Sub

Private Sub App_NewPresentation(ByVal Pres As Presentation)
Debug.Print "App_NewPresentation"
End Sub

Private Sub App_PresentationBeforeClose(ByVal Pres As Presentation, Cancel As Boolean)
Debug.Print "App_PresentationBeforeClose"
End Sub

Private Sub App_PresentationBeforeSave(ByVal Pres As Presentation, Cancel As Boolean)
Debug.Print "App_PresentationBeforeSave"
End Sub

Private Sub App_PresentationClose(ByVal Pres As Presentation)
Debug.Print "App_PresentationClose"
End Sub

Private Sub App_PresentationCloseFinal(ByVal Pres As Presentation)
Debug.Print "App_PresentationCloseFinal"
End Sub

Private Sub App_PresentationNewSlide(ByVal Sld As slide)
Debug.Print "App_PresentationNewSlide"
End Sub

Private Sub App_PresentationOpen(ByVal Pres As Presentation)
Debug.Print "App_PresentationOpen"
End Sub

Private Sub App_PresentationPrint(ByVal Pres As Presentation)
Debug.Print "App_PresentationPrint"
End Sub

Private Sub App_PresentationSave(ByVal Pres As Presentation)
Debug.Print "App_PresentationSave"
End Sub

Private Sub App_PresentationSync(ByVal Pres As Presentation, ByVal SyncEventType As Office.MsoSyncEventType)
Debug.Print "App_PresentationSync"
End Sub

Private Sub App_ProtectedViewWindowActivate(ByVal ProtViewWindow As ProtectedViewWindow)
Debug.Print "App_ProtectedViewWindowActivate"
End Sub

Private Sub App_ProtectedViewWindowBeforeClose(ByVal ProtViewWindow As ProtectedViewWindow, ByVal ProtectedViewCloseReason As PpProtectedViewCloseReason, Cancel As Boolean)
Debug.Print "App_ProtectedViewWindowBeforeClose"
End Sub

Private Sub App_ProtectedViewWindowBeforeEdit(ByVal ProtViewWindow As ProtectedViewWindow, Cancel As Boolean)
Debug.Print "App_ProtectedViewWindowBeforeEdit"
End Sub

Private Sub App_ProtectedViewWindowDeactivate(ByVal ProtViewWindow As ProtectedViewWindow)
Debug.Print "App_ProtectedViewWindowDeactivate"
End Sub

Private Sub App_ProtectedViewWindowOpen(ByVal ProtViewWindow As ProtectedViewWindow)
Debug.Print "App_ProtectedViewWindowOpen"
End Sub

Private Sub App_SlideSelectionChanged(ByVal SldRange As SlideRange)
Debug.Print "App_SlideSelectionChanged"
End Sub

Private Sub App_SlideShowBegin(ByVal Wn As SlideShowWindow)
Debug.Print "App_SlideShowBegin"
End Sub

Private Sub App_SlideShowEnd(ByVal Pres As Presentation)
Debug.Print "App_SlideShowEnd"
End Sub

Private Sub App_SlideShowNextBuild(ByVal Wn As SlideShowWindow)
Debug.Print "App_SlideShowNextBuild"
End Sub

Private Sub App_SlideShowNextClick(ByVal Wn As SlideShowWindow, ByVal nEffect As Effect)
Debug.Print "App_SlideShowNextClick"
End Sub

Private Sub App_SlideShowNextSlide(ByVal Wn As SlideShowWindow)
Debug.Print "App_SlideShowNextSlide"
End Sub

Private Sub App_SlideShowOnNext(ByVal Wn As SlideShowWindow)
Debug.Print "App_SlideShowOnNext"
End Sub

Private Sub App_SlideShowOnPrevious(ByVal Wn As SlideShowWindow)
Debug.Print "App_SlideShowOnPrevious"
End Sub

Private Sub App_WindowActivate(ByVal Pres As Presentation, ByVal Wn As DocumentWindow)
Debug.Print "App_WindowActivate"
End Sub

Private Sub App_WindowBeforeDoubleClick(ByVal Sel As Selection, Cancel As Boolean)
Debug.Print "App_WindowBeforeDoubleClick"
    If (EditGLEFigure()) Then
        Cancel = True
    End If
End Sub

Private Sub App_WindowBeforeRightClick(ByVal Sel As Selection, Cancel As Boolean)
Debug.Print "App_WindowBeforeRightClick"
End Sub

Private Sub App_WindowDeactivate(ByVal Pres As Presentation, ByVal Wn As DocumentWindow)
Debug.Print "App_WindowDeactivate"
End Sub

