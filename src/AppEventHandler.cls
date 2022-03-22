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

Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    ' happens when a selection is changed or copy and pasted
    ' try to implement something to detect a duplicte PowerGLE figure has been created using copy and paste
    Dim lshape As Shape
    Debug.Print "App_WindowSelectionChange " + CStr(Sel.Type) + " " + CStr(Sel.HasChildShapeRange)
    If Sel.Type = ppSelectionShapes Then
        ' try first shape  should go through each shape
        For Each lshape In Sel.ShapeRange
            If IsPowerGLEShape(lshape) Then
                Debug.Print "PowerGLE Shape " + lshape.name+" "+ lshape.Tags(GetShapeTagName(TAG_FIGURE_NAME)) + " " + lshape.Tags(GetShapeTagName(TAG_FIGURE_UUID))
                'does this shape uuid exist in this presentation - if so create a new uuid and change the figure name
                ' it has just been copy and pasted
                For Each mySlide in ActivePresentation.Slides
                    For each myShape in MySlide.Shapes 
                        ' loop on all groups too'
                        ' if a shape is found with the same uuid as the selected shape and a different name
                        ' then the selected shape has just been copied
                        ' change the uuid and figure name of the selected shape.
                        ' do this for all selected shapes
                        Debug.Print myShape.name 
                    Next 
                Next 
            End If
        Next
    End If
End Sub