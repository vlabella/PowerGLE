Attribute VB_Name = "PowerGLE"
'
' -- PowerGLE.bas
'
' Main functions calls and app initialization
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
Public GlobalOldShape As Shape
Public GlobalDataFiles As New Scripting.Dictionary
Public RegenerateContinue As Boolean

Public theAppEventHandler As New AppEventHandler

Sub InitializeApp()
    Debug.Print "initapp"
    Set theAppEventHandler.App = Application
    Set GlobalOldShape = Nothing
End Sub

Private Sub Auto_Open()
    ' Runs when the add-in is loaded
    Debug.Print "Auto Open"
    InitializeApp
    'Load GLEForm
    'Unload GLEForm
End Sub

Public Sub onLoadRibbon(myRibbon As IRibbonUI)
    ' runs when ribbon is loaded
    Debug.Print "Ribbon Loaded"
    InitializeApp
End Sub

Private Sub Auto_Close()
    GLEForm.UnInitializeApp
End Sub

Public Sub RibbonNewGLEFigure(ByVal control)
    NewGLEFigure
End Sub

Public Sub RibbonEditGLEFigure(ByVal control)
    EditGLEFigure
End Sub

Public Sub RibbonShowSettings(ByVal control)
    LoadSettingsForm
End Sub

Public Sub RibbonShowAbout(ByVal control)
    Load AboutBox
    AboutBox.Show
End Sub

Public Sub RibbonRegenerateSelectedDisplays(ByVal control)
    Load BatchEditForm
    BatchEditForm.Show
End Sub

Public Sub SetOldShape(s As Shape)
    Set GlobalOldShape = s
    ' Debug.Print "Set old shape " + GlobalOldShape.Tags(POWER_GLE_FIGURE_TAG)
End Sub

Public Sub ClearOldShape()
    Set GlobalOldShape = Nothing
End Sub

Function IsPowerGLEShape(lshape As Shape) As Boolean
    ' returns true if this is a powergle shape or figure
    IsPowerGLEShape = False
    If Not lshape Is Nothing Then
        If lshape.Tags(GetShapeTagName(TAG_FIGURE)) = POWER_GLE_UUID Then
            ' PowerGLE display
            IsPowerGLEShape = True
            ' For j = 1 To .count
            '    Debug.Print .name(j) & vbTab & .value(j)
            ' Next j
        End If
    End If
End Function

Function GetSelectedShape(ByRef lshape As Shape) As Boolean
    ' if there is a selected shape returns true
    ' and sets the lshape
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    Set lshape = Nothing
    GetSelectedShape = False
    If Sel.Type = ppSelectionShapes Then
        ' First make sure we don't have any shapes with duplicate names on this slide
        Call DeDuplicateShapeNamesInSlide(ActiveWindow.View.slide.SlideIndex)
        If Sel.ShapeRange.count = 1 Then ' if not 1, then multiple objects are selected
            ' Group case: either 1 object within a group, or 1 group corresponding to an EMF display
            If Sel.ShapeRange.Type = msoGroup Then
                If Sel.HasChildShapeRange = False Then ' Maybe an EMF display
                    Set lshape = Sel.ShapeRange(1)
                    GetSelectedShape = True
                ElseIf Sel.ChildShapeRange.count = 1 Then
                    ' 1 object inside a group
                    Set lshape = Sel.ChildShapeRange(1)
                    GetSelectedShape = True
                End If
            Else
                ' Non-group case: only a single object can be selected
                Set lshape = Sel.ShapeRange(1)
                GetSelectedShape = True
            End If
        End If
    End If
End Function


Sub NewGLEFigure()
    Dim Go As Boolean
    Go = False
    ClearOldShape
    GlobalDataFiles.RemoveAll
    ' must have an active presentation and saved at least once
    If Not ActivePresentation Is Nothing Then
        If ActivePresentation.path <> "" Then
            Go = True
        End If
    End If
    If Go Then
        Load GLEForm
        If GetValue(USE_EXTERNAL_EDITOR_VALUE_NAME) Then
            GLEForm.Show vbModeless
            Call GLEForm.CmdButtonExternalEditor_Click
        Else
            GLEForm.Show vbModal
        End If
    Else
        MsgBox "The current presentation must be saved once prior to adding a GLE figure."
    End If
End Sub


Function EditGLEFigure() As Boolean
    ' Check if the user currently has a single GLE figure selected.
    ' If so, display the dialog box. If not, display an error message.
    ' Called when the user clicks the "Edit GLE Figure" menu item.
    EditGLEFigure = False
    GlobalDataFiles.RemoveAll
    Dim MyShape As Shape
    Set MyShape = Nothing
    If GetSelectedShape(MyShape) Then
        If (IsPowerGLEShape(MyShape)) Then
            SetOldShape MyShape
            Load GLEForm
            If GetValue(USE_EXTERNAL_EDITOR_VALUE_NAME) Then
                GLEForm.Show vbModeless
                Call GLEForm.CmdButtonExternalEditor_Click
            Else
                GLEForm.Show vbModal
            End If
            EditGLEFigure = True
        Else
            MsgBox "A single PowerGLE figure must be selected to modify it."
        End If
    Else
        MsgBox "A single PowerGLE figure must be selected to modify it."
    End If
End Function

' Make sure there aren't multiple shapes with the same name prior to processing
Sub DeDuplicateShapeNamesInSlide(SlideIndex As Integer)
    Dim vSh As Shape
    Dim vSl As slide
    Set vSl = ActivePresentation.Slides(SlideIndex)
    
    Dim NameList() As String
    
    Dim dict As New Scripting.Dictionary
    For Each vSh In vSl.Shapes
        If vSh.Type = msoGroup Then
            NameList = CollectGroupedItemList(vSh, True)
        Else
            ReDim NameList(0 To 0) As String
            NameList(0) = vSh.name
        End If
        For n = LBound(NameList) To UBound(NameList)
            Key = NameList(n)
            If Not dict.Exists(Key) Then
                dict.Item(Key) = 1
            Else
                dict.Item(Key) = dict.Item(Key) + 1
            End If
        Next n
    Next vSh
    
    For Each vSh In vSl.Shapes
        Set dict = RenameDuplicateShapes(vSh, dict)
    Next vSh
    Set dict = Nothing
End Sub

Private Function RenameDuplicateShapes(vSh As Shape, dict As Scripting.Dictionary) As Scripting.Dictionary
    If vSh.Type = msoGroup Then
        Dim n As Long
        For n = 1 To vSh.GroupItems.count
            Set dict = RenameDuplicateShapes(vSh.GroupItems(n), dict)
        Next
    Else
        k = vSh.name
        If dict.Item(k) > 1 Then
            shpCount = 1
            Do While dict.Exists(k & " " & shpCount)
                shpCount = shpCount + 1
            Loop
            vSh.name = k & " " & shpCount
            dict.Add k & " " & shpCount, 1
        End If
    End If
    Set RenameDuplicateShapes = dict
End Function


Public Sub RegenerateSelectedDisplays()
    ' called from batch edito form
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    Dim vSh As Shape
    Dim vSl As slide
    Dim SlideIndex As Integer
    RegenerateContinue = True
    Select Case Sel.Type
        Case ppSelectionShapes
            SlideIndex = ActiveWindow.View.slide.SlideIndex
            Call DeDuplicateShapeNamesInSlide(SlideIndex)
            DisplayCount = CountDisplaysInSelection(Sel)
            If DisplayCount > 0 Then
                RegenerateForm.LabelSlideNumber.Caption = 1
                RegenerateForm.LabelTotalSlideNumber.Caption = 1
                RegenerateForm.LabelShapeNumber.Caption = 0
                RegenerateForm.LabelTotalShapeNumberOnSlide.Caption = DisplayCount
                RegenerateForm.Show False
                If Sel.HasChildShapeRange Then ' displays within a group
                    For Each vSh In Sel.ChildShapeRange
                        Call RegenerateOneDisplay(vSh)
                    Next vSh
                Else
                    For Each vSh In Sel.ShapeRange
                        If vSh.Type = msoGroup And Not IsPowerGLEShape(vSh) Then ' grouped displays
                            Call RegenerateGroupedDisplays(vSh, SlideIndex)
                        Else ' single display
                            Call RegenerateOneDisplay(vSh)
                        End If
                    Next vSh
                End If
            Else
                MsgBox "No displays to be regenerated."
            End If
        Case ppSelectionSlides
            RegenerateForm.LabelSlideNumber.Caption = 0
            RegenerateForm.LabelTotalSlideNumber.Caption = Sel.SlideRange.count
            RegenerateForm.LabelShapeNumber.Caption = 0
            RegenerateForm.LabelTotalShapeNumberOnSlide.Caption = 0
            RegenerateForm.Show False
            For Each vSl In Sel.SlideRange
                RegenerateForm.LabelSlideNumber.Caption = RegenerateForm.LabelSlideNumber.Caption + 1
                DisplayCount = CountDisplaysInSlide(vSl)
                RegenerateForm.LabelTotalShapeNumberOnSlide.Caption = DisplayCount
                DoEvents
                If DisplayCount > 0 Then
                    Call RegenerateDisplaysOnSlide(vSl)
                End If
            Next vSl
        Case Else
            MsgBox "A set of shapes or slides must be selected"
    End Select
    
    With RegenerateForm
        .Hide
        .LabelShapeNumber.Caption = 0
        .LabelSlideNumber.Caption = 0
        .LabelTotalSlideNumber.Caption = 0
        .LabelTotalShapeNumberOnSlide.Caption = 0
    End With
    Unload RegenerateForm
End Sub

Sub RegenerateDisplaysOnSlide(vSl As slide)
    vSl.Select
    Call DeDuplicateShapeNamesInSlide(vSl.SlideIndex)
    Dim vSh As Shape
    For Each vSh In vSl.Shapes
        If vSh.Type = msoGroup And Not IsPowerGLEShape(vSh) Then
            Call RegenerateGroupedDisplays(vSh, vSl.SlideIndex)
        Else
            Call RegenerateOneDisplay(vSh)
        End If
    Next vSh
End Sub

Sub RegenerateGroupedDisplays(vGroupSh As Shape, SlideIndex As Integer)
    Dim n As Long
    Dim vSh As Shape
    
    Dim ItemToRegenerateList() As String
    
    ItemToRegenerateList = CollectGroupedItemList(vGroupSh, False)
    
    For n = LBound(ItemToRegenerateList) To UBound(ItemToRegenerateList)
        Set vSh = ActivePresentation.Slides(SlideIndex).Shapes(ItemToRegenerateList(n))
        Call RegenerateOneDisplay(vSh)
    Next

End Sub

Private Function CollectGroupedItemList(vSh As Shape, AllDisplays As Boolean) As Variant
    Dim n As Long
    Dim i As Long
    Dim prev_length As Long
    Dim added_length As Long
    Dim TmpList() As String
    Dim SubList() As String
    prev_length = -1
    For n = 1 To vSh.GroupItems.count
'        If n = 1 Then
'            prev_length = -1
'        Else
'            prev_length = UBound(TmpList)
'        End If
        If vSh.GroupItems(n).Type = msoGroup Then ' this case should never occur, as PPT disregards subgroups. Consider removing.
            SubList = CollectGroupedItemList(vSh.GroupItems(n), AllDisplays)
            added_length = UBound(SubList)
            ReDim Preserve TmpList(0 To prev_length + added_length) As String
            For j = prev_length + 1 To UBound(TmpList)
                TmpList(j) = SubList(j - prev_length - 1)
            Next j
        Else
            If AllDisplays Or IsPowerGLEShape(vSh.GroupItems(n)) Then
            ReDim Preserve TmpList(0 To prev_length + 1) As String
            TmpList(UBound(TmpList)) = vSh.GroupItems(n).name
            End If
        End If
        prev_length = UBound(TmpList)
    Next
    CollectGroupedItemList = TmpList
End Function

Sub RegenerateOneDisplay(vSh As Shape)
    If RegenerateContinue Then
    vSh.Select
    With vSh.Tags
        If .Item(POWER_GLE_FIGURE_TAG) <> "" Then ' we're dealing with an PowerGLE display
            RegenerateForm.LabelShapeNumber.Caption = RegenerateForm.LabelShapeNumber.Caption + 1
            DoEvents
            Load GLEForm
            
            Call GLEForm.RetrieveOldShapeInfo(vSh)

            Apply_BatchEditSettings

            Call GLEForm.ButtonGenerate_Click
            Exit Sub
        End If
    End With
    Else
        Debug.Print "Pressed Cancel"
    End If
End Sub

Private Sub Apply_BatchEditSettings()
    ' copy batch edit settings to GLEForm so when batch modify is run they are changed
    If BatchEditForm.CheckBoxModifyTempFolder.value Then
        GLEForm.TextBoxTempFolder.Text = BatchEditForm.TextBoxTempFolder.Text
    End If
    If BatchEditForm.CheckBoxModifyOutputFormat.value Then
        GLEForm.ComboBoxOutputFormat.ListIndex = BatchEditForm.ComboBoxOutputFormat.ListIndex
    End If
    If BatchEditForm.CheckBoxModifyDPI.value Then
        GLEForm.TextBoxLocalDPI.Text = BatchEditForm.TextBoxDPI.Text
    End If
    If BatchEditForm.CheckBoxModifyUseCairo.value Then
        GLEForm.CheckBoxUseCairo.value = True
    End If
    If BatchEditForm.CheckBoxModifyPNGTransparent.value Then
        GLEForm.checkboxPNGTransparent.value = True
    End If
    If BatchEditForm.CheckBoxReplaceText.value Then
        If BatchEditForm.TextBoxFindText.Text <> "" Then
            GLEForm.TextBoxGLECode.Text = Replace(GLEForm.TextBoxGLECode.Text, BatchEditForm.TextBoxFindText.Text, BatchEditForm.TextBoxReplacementText.Text)
        End If
    End If
End Sub


Function CountDisplaysInShape(vSh As Shape) As Integer
    DisplayCount = 0
    If vSh.Type = msoGroup Then ' grouped displays
        Dim s As Shape
        For Each s In vSh.GroupItems
            DisplayCount = DisplayCount + CountDisplaysInShape(s)
        Next
    Else ' single display
        If IsPowerGLEShape(vSh) Then
            DisplayCount = 1
        End If
    End If
    CountDisplaysInShape = DisplayCount
End Function

Function CountDisplaysInSelection(Sel As Selection) As Integer
    Dim vSh As Shape
    
    DisplayCount = 0
    If Sel.HasChildShapeRange Then ' displays within a group
        For Each vSh In Sel.ChildShapeRange
            DisplayCount = DisplayCount + CountDisplaysInShape(vSh)
        Next vSh
    Else
        For Each vSh In Sel.ShapeRange
            DisplayCount = DisplayCount + CountDisplaysInShape(vSh)
        Next vSh
    End If
    CountDisplaysInSelection = DisplayCount
End Function

Function CountDisplaysInSlide(vSl As slide) As Integer
    Dim vSh As Shape
    DisplayCount = 0
    For Each vSh In vSl.Shapes
        DisplayCount = DisplayCount + CountDisplaysInShape(vSh)
    Next vSh
    CountDisplaysInSlide = DisplayCount
End Function


Sub LoadSettingsForm()
    Load SettingsForm
    SettingsForm.Show
End Sub



