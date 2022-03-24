VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GLEForm 
   Caption         =   "PowerGLE"
   ClientHeight    =   8360.001
   ClientLeft      =   11
   ClientTop       =   330
   ClientWidth     =   7898
   OleObjectBlob   =   "GLEForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GLEForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' -- GLEForm.frm
'
' Form that handles creation and editing of GLE code
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
Dim TemplateSortedListString As String
Dim TemplateSortedList() As String
Dim TemplateNameSortedListString As String
Dim FormHeightWidthSet As Boolean
Dim DoneWithActivation As Boolean

Private Sub UserForm_Initialize()
    ' called when form is loaded
    LoadSettings ' loads the defaults
    Dim OutputPath As String
    OutputPath = GetTempPath()
    If OutputPath = "" Then
        ' error most likely due to not having saved the file at least once
        Exit Sub
    End If
    CreateFolder (OutputPath) ' make sure it exists
    If GlobalOldShape Is Nothing Then
        ButtonGenerate.Caption = "Generate"
        ButtonGenerateAndClose.Caption = "Generate & Close"
        TextBoxFigureName = GetNextFigureName(OutputPath)
    Else
        RetrieveOldShapeInfo GlobalOldShape
        ButtonGenerate.Caption = "Regenerate"
        ButtonGenerateAndClose.Caption = "Regenerate & Close"
        TextBoxFigureName = GlobalOldShape.Tags(GetShapeTagName(TAG_FIGURE_NAME))
    End If
    ButtonGenerate.Accelerator = "G"
    ButtonGenerateAndClose.Accelerator = "C"
    ' With multiple monitors, the "CenterOwner" option to open the UserForm in the center of the parent window
    ' does not seem to work, at least in Office 2010.
    ' The following code to manually place the UserForm somehow makes the "CenterOwner" option work.
    ' Remark: if used with the Manual placement option, it would place the window to the left, under the ribbon.
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
End Sub

Private Function isFormModeless() As Boolean
    On Error GoTo EH
    Me.Show vbModeless
    isFormModeless = True
    Exit Function
EH:
    isFormModeless = False
End Function

Private Sub UserForm_Activate()
    DoneWithActivation = False
    ' We have to be careful of the case where the edit window gets activated in vbModeless mode
    If Not isFormModeless Then
        'Execute macro to enable resizeability
        MakeFormResizable
        If Not FormHeightWidthSet Then
            GLEForm.Height = GetValue(GLE_FORM_HEIGHT_VALUE_NAME)
            GLEForm.Width = GetValue(GLE_FORM_WIDTH_VALUE_NAME)
        End If
        ResizeForm
        DoneWithActivation = True
    End If
End Sub

Private Sub SaveSettings()
    ' take values from gui and save in the registry
    SetValue INITIAL_SOURCECODE_VALUE_NAME, CStr(TextBoxGLECode.Text)
    SetValue SOURCECODE_CURSOR_POSITION_VALUE_NAME, CLng(TextBoxGLECode.SelStart)
    SetValue USE_CAIRO_VALUE_NAME, BoolToInt(CheckBoxUseCairo.value)
    SetValue BITMAP_DPI_VALUE_NAME, CStr(TextBoxLocalDPI.Text)
    SetValue PNG_TRANSPARENT_VALUE_NAME, BoolToInt(checkboxPNGTransparent.value)
    SetValue DEBUG_VALUE_NAME, BoolToInt(checkboxDebug.value)
    SetValue EDITOR_FONT_SIZE_VALUE_NAME, CStr(TextBoxGLECode.Font.size)
    SetValue EDITOR_WORD_WRAP_VALUE_NAME, BoolToInt(TextBoxGLECode.WordWrap)
    SetValue GLE_FORM_HEIGHT_VALUE_NAME, GLEForm.Height
    SetValue GLE_FORM_WIDTH_VALUE_NAME, GLEForm.Width
End Sub

Private Sub LoadSettings()
    ' populates GUI  user elements from what is in the registry or defaults
    TextBoxGLECode.Text = GetValue(INITIAL_SOURCECODE_VALUE_NAME)
    TextBoxGLECode.SelStart = GetValue(SOURCECODE_CURSOR_POSITION_VALUE_NAME)
    CheckBoxUseCairo.value = GetValue(USE_CAIRO_VALUE_NAME)
    TextBoxLocalDPI.Text = GetValue(BITMAP_DPI_VALUE_NAME)
    checkboxPNGTransparent.value = GetValue(PNG_TRANSPARENT_VALUE_NAME)
    checkboxDebug.value = GetValue(DEBUG_VALUE_NAME)
    TextBoxGLECode.Font.size = GetValue(EDITOR_FONT_SIZE_VALUE_NAME)
    TextBoxGLECode.WordWrap = GetValue(EDITOR_WORD_WRAP_VALUE_NAME)
    ToggleButtonWrap.value = TextBoxGLECode.WordWrap
    TextBoxTempFolder.Text = GetTempPath(False)
    ComboBoxOutputFormat.List = ArrayFromCSVList(OUTPUT_FORMATS)
    ComboBoxOutputFormat.ListIndex = GetArrayIndex(OUTPUT_FORMATS, GetValue(OUTPUT_FORMAT_VALUE_NAME))
End Sub

Private Sub AddTagsToShape(vSh As Shape)
    ' takes gui elements and stores them in the shape
    Dim Index As Integer
    With vSh.Tags
        .Add GetShapeTagName(TAG_FIGURE), POWER_GLE_UUID
        .Add GetShapeTagName(TAG_VERSION), POWER_GLE_VERSION_NUMBER
        .Add GetShapeTagName(TAG_SOURCE_CODE), TextBoxGLECode.Text
        .Add GetShapeTagName(EDITOR_FONT_SIZE_VALUE_NAME), val(TextBoxGLECode.Font.size)
        .Add GetShapeTagName(SOURCECODE_CURSOR_POSITION_VALUE_NAME), TextBoxGLECode.SelStart
        .Add GetShapeTagName(TAG_FIGURE_NAME), TextBoxFigureName.Text
        .Add GetShapeTagName(TAG_TEMP_FOLDER), TextBoxTempFolder.Text
        .Add GetShapeTagName(GLE_FORM_HEIGHT_VALUE_NAME), GLEForm.Height
        .Add GetShapeTagName(GLE_FORM_WIDTH_VALUE_NAME), GLEForm.Width
        .Add GetShapeTagName(EDITOR_WORD_WRAP_VALUE_NAME), TextBoxGLECode.WordWrap
        .Add GetShapeTagName(USE_CAIRO_VALUE_NAME), CheckBoxUseCairo.value
        .Add GetShapeTagName(PNG_TRANSPARENT_VALUE_NAME), checkboxPNGTransparent.value
    End With
    Index = 1
    For Each i In GlobalDataFiles.Keys
        vSh.Tags.Add GetShapeTagName(TAG_DATA_FILENAME) + "_" + CStr(Index), CStr(i)
        vSh.Tags.Add GetShapeTagName(TAG_DATA_FILE_CONTENT) + "_" + CStr(Index), GlobalDataFiles(i)
        Index = Index + 1
    Next i
End Sub

Sub RetrieveOldShapeInfo(oldshape As Shape)
    ' get shape info and populate the gui
    Dim FormHeightSet As Boolean
    Dim FormWidthSet As Boolean
    Dim CursorPosition As Integer
    CursorPosition = 0
    FormHeightSet = False
    FormWidthSet = False
    
    With oldshape.Tags
        If .Item(GetShapeTagName(TAG_SOURCE_CODE)) <> "" Then
            TextBoxGLECode.Text = .Item(GetShapeTagName(TAG_SOURCE_CODE))
            CursorPosition = Len(TextBoxGLECode.Text)
        End If
        If .Item(GetShapeTagName(EDITOR_FONT_SIZE_VALUE_NAME)) <> "" Then
            TextBoxGLECode.Font.size = CInt(.Item(GetShapeTagName(EDITOR_FONT_SIZE_VALUE_NAME)))
        End If
        If .Item(GetShapeTagName(TAG_OUTPUT_DPI)) <> "" Then
            TextBoxLocalDPI.Text = .Item(GetShapeTagName(TAG_OUTPUT_DPI))
        End If
        If .Item(GetShapeTagName(TAG_FIGURE_NAME)) <> "" Then
            TextBoxFigureName.Text = .Item(GetShapeTagName(TAG_FIGURE_NAME))
        End If
        If .Item(GetShapeTagName(TAG_TEMP_FOLDER)) <> "" Then
            TextBoxTempFolder.Text = .Item(GetShapeTagName(TAG_TEMP_FOLDER))
        End If
        If .Item(GetShapeTagName(PNGTRANSPARENT_VALUE_NAME)) <> "" Then
            checkboxPNGTransparent.value = SanitizeBoolean(.Item(GetShapeTagName(PNGTRANSPARENT_VALUE_NAME)), True)
        End If
        If .Item(GetShapeTagName(SOURCECODE_CURSOR_POSITION_VALUE_NAME)) <> "" Then
            CursorPosition = CInt(.Item(GetShapeTagName(SOURCECODE_CURSOR_POSITION_VALUE_NAME)))
        End If
        If .Item(GetShapeTagName(USE_CAIRO_VALUE_NAME)) <> "" Then
            CheckBoxUseCairo = SanitizeBoolean(.Item(GetShapeTagName(USE_CAIRO_VALUE_NAME)), True)
        End If
        If .Item(GetShapeTagName(GLE_FORM_HEIGHT_VALUE_NAME)) <> "" Then
            GLEForm.Height = .Item(GetShapeTagName(GLE_FORM_HEIGHT_VALUE_NAME))
            FormHeightSet = True
        End If
        If .Item(GetShapeTagName(GLE_FORM_WIDTH_VALUE_NAME)) <> "" Then
            GLEForm.Width = .Item(GetShapeTagName(GLE_FORM_WIDTH_VALUE_NAME))
            FormWidthSet = True
        End If
        If .Item(GetShapeTagName(EDITOR_WORD_WRAP_VALUE_NAME)) <> "" Then
            TextBoxGLECode.WordWrap = SanitizeBoolean(.Item(GetShapeTagName(EDITOR_WORD_WRAP_VALUE_NAME)), True)
            ToggleButtonWrap.value = TextBoxGLECode.WordWrap
        End If
    End With
    
    Dim filenames As New Scripting.Dictionary
    Dim Contents As New Scripting.Dictionary
    Dim allMatches As Object
    Dim RXFilename As Object
    Dim RXContent As Object
    
    Set RXFilename = CreateObject("VBScript.RegExp")
    RXFilename.Global = True
    RXFilename.IgnoreCase = True
    With RXFilename
        .Pattern = "(^" + GetShapeTagName(TAG_DATA_FILENAME) + "_" + "([0-9]+)$)"
    End With
    Set RXContent = CreateObject("VBScript.RegExp")
    RXContent.Global = True
    RXContent.IgnoreCase = True
    With RXContent
        .Pattern = "(^" + GetShapeTagName(TAG_DATA_FILE_CONTENT) + "_" + "([0-9]+)$)"
    End With
    With oldshape.Tags
        For j = 1 To .count
            Debug.Print .name(j) & vbTab & .value(j)
            Set allMatches = RXFilename.Execute(.name(j))
            If allMatches.count <> 0 Then
                filenames.Add allMatches.Item(0).submatches.Item(1), .value(j)
            End If
            Set allMatches = RXContent.Execute(.name(j))
            If allMatches.count <> 0 Then
                Contents.Add allMatches.Item(0).submatches.Item(1), .value(j)
            End If
        Next j
    End With
    For Each i In filenames
        GlobalDataFiles.Add filenames(i), Contents(i)
        ListBoxDataFiles.AddItem filenames(i)
    Next i
    FormHeightWidthSet = FormHeightSet And FormWidthSet
    TextBoxGLECode.SelStart = CursorPosition
End Sub

Private Function SanitizeBoolean(Str As String, Def As Boolean) As Boolean
    On Error GoTo ErrWrongBoolean:
    SanitizeBoolean = CBool(Str)
    Exit Function
ErrWrongBoolean:
    SanitizeBoolean = Def
    Resume Next
End Function


Private Sub UserForm_Resize()
    ' Minimal size
    If GLEForm.Height < GLE_FORM_MIN_HEIGHT Then
        GLEForm.Height = GLE_FORM_MIN_HEIGHT
    End If
    If GLEForm.Width < GLE_FORM_MIN_WIDTH Then
        GLEForm.Width = GLE_FORM_MIN_WIDTH
    End If
    ResizeForm
End Sub

Private Sub ResizeForm()
    Dim bordersize As Integer
    bordersize = 6
    MultiPage1.Left = bordersize
    MultiPage1.Top = bordersize
    MultiPage1.Width = GLEForm.Width - bordersize * 2
    MultiPage1.Height = GLEForm.Height - MultiPage1.Top
    TextBoxGLECode.Width = MultiPage1.Width - bordersize * 2
    TextBoxGLECode.Height = MultiPage1.Height - TextBoxGLECode.Top - FrameControls.Height - 7 * bordersize
    FrameControls.Top = TextBoxGLECode.Top + TextBoxGLECode.Height
End Sub


Private Sub ButtonCancel_Click()
    Unload GLEForm
    ' GLEForm.Hide
End Sub

Sub ButtonGenerate_Click()
    RunGLE
End Sub

Sub ButtonGenerateAndClose_Click()
    RunGLE
    Set GlobalOldShape = Nothing
    Unload GLEForm
End Sub

Sub RunGLE()
    Dim OutputPath As String
    OutputPath = GetTempPath()
    CreateFolder (OutputPath) ' make sure it exists
    If OutputPath = "" Then
        Exit Sub
    End If
    ' store this UUID in the shape info for future reference
    Dim FigureUUID As String
    Dim FigureName As String
    ' this get populated upon init or user changes it
    FigureName = TextBoxFigureName.value
    If GlobalOldShape Is Nothing Then
        FigureUUID = GetUUID()
    Else
        FigureUUID = GlobalOldShape.Tags(GetShapeTagName(TAG_FIGURE_UUID))
        If FigureName <> GlobalOldShape.Tags(GetShapeTagName(TAG_FIGURE_NAME)) Then
            ' user is changing name - ok so rename the folder
            ' but make sure new folder does not exist
            ' make new unique name ie boxes_1 or cancel so user cna change
        End If
    End If
    ' store each file in its own unique folder named TEMP_FILENAME_##
    ' eg figure_1 figure_2 etc.
    ' which matches gle filename
    OutputPath = AddSlash(AddSlash(OutputPath) + FigureName)
    ' does folder exist? need to warn user for new figures only
    ' probably combine this with above folder seletion
    CreateFolder (OutputPath)
        
    Dim debugMode As Boolean
    debugMode = checkboxDebug.value
        
    Dim OutputFormat, OutputFileExt, OutputFilename As String
    Dim OutputFormatIndex As Integer
    OutputFormatIndex = ComboBoxOutputFormat.ListIndex
    OutputFormat = GetArrayValue(OUTPUT_FORMATS, OutputFormatIndex)
    OutputFileExt = GetArrayValue(OUTPUT_FORMAT_FILE_EXT, OutputFormatIndex)
    OutputFilename = FigureName & "." & OutputFileExt
     
    Dim TimeOutTimeSeconds As Long
    TimeOutTimeSeconds = GetValue(TIMEOUT_VALUE_NAME)
    TimeOutTime = val(TimeOutTimeString) * 1000
    
    Dim OutputDpiString As String
    OutputDpiString = TextBoxLocalDPI.Text
    Dim OutputDpi As Long
    OutputDpi = val(OutputDpiString)
      
    ' Test if path writable
    If Not IsPathWritable(OutputPath) Then
        MsgBox "The temporary folder " & OutputPath & " is not writable."
        Exit Sub
    End If
    SaveTextFile OutputPath, FigureName + "." + GLE_EXT, TextBoxGLECode.Text, GetValue(USE_UTF8_VALUE_NAME), True
    For Each i In GlobalDataFiles.Keys
        SaveTextFile OutputPath, CStr(i), GlobalDataFiles(i), GetValue(USE_UTF8_VALUE_NAME), True
    Next i
    ' Run GLE  -- create Bitmap File for Inclusion
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim LogFile As Object
    Dim Cairo As String
    Cairo = ""
    If (CheckBoxUseCairo.value) Then
        Cairo = " /cairo "
    End If
    Dim Transparent As String
    Transparent = ""
    If (checkboxPNGTransparent And OutputFormat = "PNG") Then
        Transparent = " /transparent "
    End If
    ' FrameProcess.Visible = True
    LabelProcess.Caption = "GLE to PNG ..."
    ' FrameProcess.Repaint
    Dim GLEExecutable As String
    GLEExecutable = GetValue(GLE_EXECUTABLE_VALUE_NAME)
    Dim cmd As String
    cmd = Quote(GLEExecutable) + " /output " + Quote(OutputPath + OutputFilename) + " /device " + OutputFormat + " /resolution " + OutputDpiString + Cairo + Transparent + Quote(OutputPath + FigureName + "." + GLE_EXT)
    Debug.Print cmd
    'TextBoxGLECode.Text = cmd
    RetVal& = Execute(cmd, OutputPath, debugMode, TimeOutTimeSeconds * 1000)
    If (Not fs.FileExists(OutputPath + OutputFilename)) Then
        ' Error in  GLE
            MsgBox "GLE did not return in " & CStr(TimeOutTimeSeconds) & " seconds and may have hung." _
            & vbNewLine & "Please make sure your code compiles outside of PowerGLE." _
            & vbNewLine & "You may also try generating in Debug mode, as it will let you know if any font/package is missing."
        'FrameProcess.Visible = False
        Exit Sub
    End If
    ' GLE run successful.
    ' Insert the Image into the slide
    LabelProcess.Caption = "Insert image..."
    ' FrameProcess.Repaint
    Dim PosX As Single
    Dim PosY As Single
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    Dim s As Shape
    IsInGroup = False
    If Not GlobalOldShape Is Nothing Then
        ' If Edit mode, store parameters of old image
        If Sel.ShapeRange.Type = msoGroup And Sel.HasChildShapeRange Then
            ' Old image is part of a group
            ' Set oldshape = Sel.ChildShapeRange(1)
            IsInGroup = True
            Dim arr() As Variant ' gather all shapes to be regrouped later on
            j = 0
            For Each s In Sel.ShapeRange.GroupItems
                If s.name <> oldshape.name Then
                    j = j + 1
                    ReDim Preserve arr(1 To j)
                    arr(j) = s.name
                End If
            Next
            ' Store the group's animation and Zorder info in a dummy object tmpGroup
            Dim oldGroup As Shape
            Set oldGroup = Sel.ShapeRange(1)
            Dim tmpGroup As Shape
            Set tmpGroup = ActiveWindow.Selection.SlideRange.Shapes.AddShape(msoShapeDiamond, 1, 1, 1, 1)
            MoveAnimation oldGroup, tmpGroup
            MatchZOrder oldGroup, tmpGroup
            ' Tag all elements in the group with their hierarchy level and their name or group name
            Dim MaxGroupLevel As Long
            MaxGroupLevel = TagGroupHierarchy(arr, GlobalOldShape.name)
        Else
            ' Set oldshape = Sel.ShapeRange(1)
        End If
        PosX = GlobalOldShape.Left
        PosY = GlobalOldShape.Top
    Else
        ' default position
        PosX = DEFAULT_SLIDE_POSTIION_X
        PosY = DEFAULT_SLIDE_POSTIION_Y
        If Sel.Type = ppSelectionShapes Then ' if something is selected on a slide, use its position for the new display
            'If Sel.ShapeRange.Type = msoGroup And Sel.HasChildShapeRange Then
            '    Set oldshape = Sel.ChildShapeRange(1)
            'Else
            '    Set oldshape = Sel.ShapeRange(1)
            'End If
            PosX = Sel.ShapeRange(1).Left
            PosY = Sel.ShapeRange(1).Top
        End If
    End If
    ' Get scaling factors
    Dim tScaleWidth As Single, tScaleHeight As Single, ScalingGain As Double
    ScalingGain = GetValue(SCALING_GAIN_VALUE_NAME)
    If ScalingGain = 0 Then
        ScalingGain = DEFAULT_SCALING_GAIN
        SetValue SCALING_GAIN_VALUE_NAME, ScalingGain
    End If
    MagicScalingFactor = lDotsPerInch() / OutputDpi * ScalingGain
    tScaleWidth = MagicScalingFactor
    tScaleHeight = tScaleWidth
    ' Insert image and rescale it
    Dim newShape As Shape
    Set newShape = AddDisplayShape(OutputPath + OutputFilename, PosX, PosY)
    ' Resize to the true size of the png file and adjust using the manual scaling factors set in Main Settings
    With newShape
        .ScaleHeight 1#, msoTrue
        .ScaleWidth 1#, msoTrue
        .LockAspectRatio = msoFalse
        ' not sure why this is needed since its set below
        .ScaleHeight ScalingGain, msoFalse
        .ScaleWidth ScalingGain, msoFalse
        .Tags.Add GetShapeTagName(TAG_OUTPUT_DPI), OutputDpi ' Stores this display's resolution
        ' Add tags storing the original height and width, used next time to keep resizing ratio.
        .Tags.Add GetShapeTagName(TAG_ORIGINAL_HEIGHT), newShape.Height
        .Tags.Add GetShapeTagName(TAG_ORIGINAL_WIDTH), newShape.Width
        .Tags.Add GetShapeTagName(TAG_FIGURE_UUID), FigureUUID
        .Tags.Add GetShapeTagName(TAG_FIGURE_NAME), FigureName
        .Tags.Add GetShapeTagName(TAG_OUTPUT_FORMAT), OutputFormat
        .Tags.Add GetShapeTagName(TAG_SLIDE_INDEX), ActiveWindow.View.slide.SlideIndex
        ' Apply scaling factors
        .ScaleHeight tScaleHeight, msoFalse
        .ScaleWidth tScaleWidth, msoFalse
        .LockAspectRatio = msoTrue
    End With
    If Not GlobalOldShape Is Nothing Then
        ' Force the new shape to have the same size as the old shape
        With newShape
            .LockAspectRatio = msoFalse
            .Height = GlobalOldShape.Height
            .Width = GlobalOldShape.Width
            .LockAspectRatio = msoTrue
        End With
        ' preserve old shape rotation
        newShape.Rotation = GlobalOldShape.Rotation
        newShape.LockAspectRatio = GlobalOldShape.LockAspectRatio ' Unlock aspect ratio if old display had it unlocked
    End If
    ' Add tags
    Call AddTagsToShape(newShape)
    ' Copy animation settings and formatting from old image, then delete it
    If Not GlobalOldShape Is Nothing Then
        Dim TransferDesign As Boolean
        TransferDesign = True
        If IsInGroup Then
            ' Transfer format to new shape
            MatchZOrder GlobalOldShape, newShape
            If TransferDesign Then
                GlobalOldShape.PickUp
                newShape.Apply
            End If
            ' Handle the case of shape within EMF group.
            Dim DeleteLowestLayer As Boolean
            DeleteLowestLayer = False
            If GlobalOldShape.Tags.Item("EMFchild") <> "" Then
                DeleteLowestLayer = True
            End If
            GlobalOldShape.Delete
            
            Dim newGroup As Shape
            ' Get current slide, it will be used to group ranges
            Dim Sld As slide
            Dim SlideIndex As Long
            SlideIndex = ActiveWindow.View.slide.SlideIndex
            Set Sld = ActivePresentation.Slides(SlideIndex)

            ' Group all non-modified elements from old group, plus modified element
            j = j + 1
            ReDim Preserve arr(1 To j)
            arr(j) = newShape.name
            If DeleteLowestLayer Then
                Dim arr_remain() As Variant
                j_remain = 0
                For Each n In arr
                    Set s = ActiveWindow.Selection.SlideRange.Shapes(n)
                    ThisShapeLevel = 0
                    For i_tag = 1 To s.Tags.count
                        If (s.Tags.name(i_tag) = "LAYER") Then
                            ThisShapeLevel = val(s.Tags.value(i_tag))
                        End If
                    Next
                    If ThisShapeLevel = 1 Then
                        s.Delete
                    Else
                        j_remain = j_remain + 1
                        ReDim Preserve arr_remain(1 To j_remain)
                        arr_remain(j_remain) = s.name
                    End If
                Next
                newShape.Tags.Add "LAYER", 2
                arr = arr_remain
            Else
                newShape.Tags.Add "LAYER", 1
            End If
            newShape.Tags.Add "SELECTIONNAME", newShape.name
            
            ' Hierarchically re-group elements
            For Level = 1 To MaxGroupLevel
                Dim CurrentLevelArr() As Variant
                j_current = 0
                For Each n In arr
                    ThisShapeLevel = 0
                    Dim ThisShapeSelectionName As String
                    ThisShapeSelectionName = ""
                    On Error Resume Next
                    With ActiveWindow.Selection.SlideRange.Shapes(n).Tags
                        For i_tag = 1 To .count
                            If (.name(i_tag) = "LAYER") Then
                                ThisShapeLevel = val(.value(i_tag))
                            End If
                            If (.name(i_tag) = "SELECTIONNAME") Then
                                ThisShapeSelectionName = .value(i_tag)
                            End If
                        Next
                    End With
                    
                    
                    If ThisShapeLevel = Level Then
                        If j_current > 0 Then
                            If Not IsInArray(CurrentLevelArr, ThisShapeSelectionName) Then
                                j_current = j_current + 1
                                ReDim Preserve CurrentLevelArr(1 To j_current)
                                CurrentLevelArr(j_current) = ThisShapeSelectionName
                            End If
                        Else
                            j_current = j_current + 1
                            ReDim Preserve CurrentLevelArr(1 To j_current)
                            CurrentLevelArr(j_current) = ThisShapeSelectionName
                        End If
                    End If
                Next
                
                If j_current > 1 Then
                    Set newGroup = Sld.Shapes.Range(CurrentLevelArr).Group
                    j = j + 1
                    ReDim Preserve arr(1 To j)
                    arr(j) = newGroup.name
                    newGroup.Tags.Add "SELECTIONNAME", newGroup.name
                    newGroup.Tags.Add "LAYER", Level + 1
                End If
                
            Next
            
            ' Delete the tags to avoid conflict with future runs
            For Each n In arr
                On Error Resume Next
                    ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Delete ("SELECTIONNAME")
                    ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Delete ("LAYER")
            Next
            
            ' Use temporary group to retrieve the group's original animation and Zorder
            MoveAnimation tmpGroup, newGroup
            MatchZOrder tmpGroup, newGroup
            tmpGroup.Delete
        Else
            ' not in group
            MoveAnimation GlobalOldShape, newShape
            MatchZOrder GlobalOldShape, newShape
            If TransferDesign Then
                GlobalOldShape.PickUp
                newShape.Apply
                GlobalOldShape.Delete
            Else
                GlobalOldShape.Delete
            End If
        End If
    End If
    
    ' Add Alternative Text
    FormAltTextandTitle newShape
    
    ' Select the new shape
    newShape.Select
    Set GlobalOldShape = newShape ' save in case user is not closing
    ' Delete temp files if not in debug mode only if user wants - make this an option
    ' If (Not debugMode) And (Not UseExternalEditor) Then fs.DeleteFile OutputPath + FilePrefix + "*.*"
    ' and remove temp UUID folder too'
       
    ' FrameProcess.Visible = False
    If GetValue(PRESERVE_TEMP_FILES_VALUE_NAME) <> True Then
        ' dont implement'
        'Dim FSO As New FileSystemObject
        'Set FSO = CreateObject("Scripting.FileSystemObject")
        'FSO.DeleteFolder "C:\TestFolder", False
    End If
End Sub



Private Function LineToFreeform(s As Shape) As Shape
    t = s.Line.Weight
    Dim ApplyTransform As Boolean
    ApplyTransform = True
    
    Dim bHflip As Boolean
    Dim bVflip As Boolean
    Dim nBegin As Long
    Dim nEnd As Long
    Dim aC(1 To 4, 1 To 2) As Double
    
    With s
        aC(1, 1) = .Left:           aC(1, 2) = .Top
        aC(2, 1) = .Left + .Width:  aC(2, 2) = .Top
        aC(3, 1) = .Left:           aC(3, 2) = .Top + .Height
        aC(4, 1) = .Left + .Width:  aC(4, 2) = .Top + .Height
    
        bHflip = .HorizontalFlip
        bVflip = .VerticalFlip
    End With
    
    If bHflip = bVflip Then
        If bVflip = False Then
            ' down to right -- South-East
            nBegin = 1: nEnd = 4
        Else
            ' up to left -- North-West
            nBegin = 4: nEnd = 1
        End If
    ElseIf bHflip = False Then
        ' up to right -- North-East
        nBegin = 3: nEnd = 2
    Else
        ' down to left -- South-West
        nBegin = 2: nEnd = 3
    End If
    xs = aC(nBegin, 1)
    ys = aC(nBegin, 2)
    xe = aC(nEnd, 1)
    ye = aC(nEnd, 2)
    
    ' Get unit vector in orthogonal direction
    xd = xe - xs
    yd = ye - ys
    
    s_length = Sqr(xd * xd + yd * yd)
    If s_length > 0 Then
    n_x = -yd / s_length
    n_y = xd / s_length
    Else
    n_x = 0
    n_y = 0
    End If
    
    x1 = xs + n_x * t / 2
    y1 = ys + n_y * t / 2
    x2 = xe + n_x * t / 2
    y2 = ye + n_y * t / 2
    x3 = xe - n_x * t / 2
    y3 = ye - n_y * t / 2
    x4 = xs - n_x * t / 2
    y4 = ys - n_y * t / 2
        
    'End If
        
    If ApplyTransform Then
        Dim builder As FreeformBuilder
        Set builder = ActiveWindow.Selection.SlideRange(1).Shapes.BuildFreeform(msoEditingCorner, x1, y1)
        builder.AddNodes msoSegmentLine, msoEditingAuto, x2, y2
        builder.AddNodes msoSegmentLine, msoEditingAuto, x3, y3
        builder.AddNodes msoSegmentLine, msoEditingAuto, x4, y4
        builder.AddNodes msoSegmentLine, msoEditingAuto, x1, y1
        Dim oSh As Shape
        Set oSh = builder.ConvertToShape
        oSh.Fill.ForeColor = s.Line.ForeColor
        oSh.Fill.Visible = msoTrue
        oSh.Line.Visible = msoFalse
        oSh.Rotation = s.Rotation
        Set LineToFreeform = oSh
    Else
        Set LineToFreeform = s
    End If
End Function

Private Function TagGroupHierarchy(arr As Variant, TargetName As String) As Long
    ' Arr is the list of names of (leaf) elements in this group
    ' TargetName is the display which is being modified. We're going down the branch containing it.
    Dim Sel As Selection
    ActiveWindow.Selection.SlideRange.Shapes(TargetName).Select
    Set Sel = Application.ActiveWindow.Selection
    
    ' This function expects to receive a grouped ShapeRange
    ' We ungroup to reveal the structure at the layer below
    Sel.ShapeRange.Ungroup
    ActiveWindow.Selection.SlideRange.Shapes(TargetName).Select
           
    If Sel.ShapeRange.Type = msoGroup Then
        ' We need to go further down, the element being edited is still within a group
        ' Get the name of the Target group in which it is
        TargetGroupName = Sel.ShapeRange(1).name
        
        Dim Arr_In() As Variant ' shapes in the same group
        Dim Arr_Out() As Variant ' shapes not in the same group
        
        ' Split range according to whether elements are in the same group or not
        j_in = 0
        j_out = 0
        For Each n In arr
            ActiveWindow.Selection.SlideRange.Shapes(n).Select
            If Sel.ShapeRange.Type = msoGroup Then
                ' object is in group
                If Sel.ShapeRange(1).name = TargetGroupName Then
                    j_in = j_in + 1
                    ReDim Preserve Arr_In(1 To j_in)
                    Arr_In(j_in) = n
                Else
                    j_out = j_out + 1
                    ReDim Preserve Arr_Out(1 To j_out)
                    Arr_Out(j_out) = n
                End If
            Else ' object not in group, so it can't be in the same group as Target
                j_out = j_out + 1
                ReDim Preserve Arr_Out(1 To j_out)
                Arr_Out(j_out) = n
            End If
        Next
        
        ' Build shape range with all elements in that group, go one level down
        Tmp = TagGroupHierarchy(Arr_In, TargetName)
        TagGroupHierarchy = Tmp + 1
        
        ' For all elements not in that group, tag them
        For Each n In Arr_Out
            ActiveWindow.Selection.SlideRange.Shapes(n).Select
            ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "LAYER", TagGroupHierarchy
            If Sel.ShapeRange.Type = msoGroup Then
                ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "SELECTIONNAME", Sel.ShapeRange(1).name
            Else
                ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "SELECTIONNAME", n
            End If
        Next
        
    Else ' we reached the final layer: the element being edited is by itself,
         ' all other elements will need to be handled either through their group
         ' name if in a group, or their name if not
        TagGroupHierarchy = 1
        For Each n In arr
            ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "LAYER", TagGroupHierarchy
            ActiveWindow.Selection.SlideRange.Shapes(n).Tags.Add "SELECTIONNAME", n
        Next
    End If


End Function

' Add picture as shape taking care of not inserting it in empty placeholder
Private Function AddDisplayShape(path As String, PosX As Single, PosY As Single) As Shape
' from http://www.vbaexpress.com/forum/showthread.php?47687-Addpicture-adds-the-picture-to-a-placeholder-rather-as-a-new-shape
' modified based on http://www.vbaexpress.com/forum/showthread.php?37561-Delete-empty-placeholders
    Dim oshp As Shape
    Dim osld As slide
    On Error Resume Next
    Set osld = ActiveWindow.Selection.SlideRange(1)
    If Err <> 0 Then Exit Function
    On Error GoTo 0
    For Each oshp In osld.Shapes
        If oshp.Type = msoPlaceholder Then
            If oshp.PlaceholderFormat.ContainedType = msoAutoShape Then
                If oshp.HasTextFrame Then
                    If Not oshp.TextFrame.HasText Then oshp.TextFrame.TextRange = "DUMMY"
                End If
            End If
        End If
    Next oshp
    Set AddDisplayShape = osld.Shapes.AddPicture(path, msoFalse, msoTrue, PosX, PosY, -1, -1)
    For Each oshp In osld.Shapes
        If oshp.Type = msoPlaceholder Then
            If oshp.PlaceholderFormat.ContainedType = msoAutoShape Then
                If oshp.HasTextFrame Then
                    If oshp.TextFrame.TextRange = "DUMMY" Then oshp.TextFrame.DeleteText
                End If
            End If
        End If
    Next oshp
End Function

Private Sub ButtonGLEFileImport_Click()
    MultiPage1.value = 0
    ' import data file into form
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = True
    fd.InitialFileName = "*." + GLE_EXT
    fd.Filters.Clear
    fd.ButtonName = "Import"
    'fd.Filters.Add NameFilterDesc
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            'import the file here'
            Debug.Print vrtSelectedItem
            Dim fs
            Dim DataFile As Object
            Const ForReading = 1, ForWriting = 2, ForAppending = 3
            Set fs = CreateObject("Scripting.FileSystemObject")
            If fs.FileExists(vrtSelectedItem) Then
                Set DataFile = fs.OpenTextFile(vrtSelectedItem, ForReading)
                TextBoxGLECode.Text = DataFile.ReadAll
                DataFile.Close
            End If
        Next vrtSelectedItem
    End If
    Set fd = Nothing
    Call ToggleInputMode
End Sub


Private Sub ButtonMakeDefault_Click()
    SaveSettings
    Select Case MultiPage1.value
        Case 0 ' Direct input
          ''  TextBoxGLECode.SetFocus
        Case 1 ' Read from file
          ''  TextBoxFile.SetFocus
        Case Else ' Templates
         ''   TextBoxTemplateCode.SetFocus
    End Select
End Sub

Sub CmdButtonExternalEditor_Click()
    Dim OutputPath As String
    Dim Filename As String
    
    OutputPath = GetTempPath()
    CreateFolder (OutputPath) ' make sure it exists
    If OutputPath = "" Then
        Exit Sub
    End If
    Dim FigureName As String
    ' this get populated upon init or user changes it
    FigureName = TextBoxFigureName.value
    OutputPath = AddSlash(AddSlash(OutputPath) + FigureName)
    ' does folder exist? need to warn user for new figures only
    ' probably combine this with above folder seletion
    CreateFolder (OutputPath)
    Filename = FigureName + "." + GLE_EXT
    If Not DoneWithActivation Then
        UserForm_Activate
    End If
    ' Write gle code to file and call external editor
    SaveTextFile OutputPath, Filename, TextBoxGLECode.Text, GetValue(USE_UTF8_VALUE_NAME), True
    ' Launch external editor
    On Error GoTo ShellError
    Debug.Print Quote(GetValue(EXTERNAL_EDITOR_EXECUTABLE_VALUE_NAME)) + " " + Quote(OutputPath + Filename)
    Shell Quote(GetValue(EXTERNAL_EDITOR_EXECUTABLE_VALUE_NAME)) + " " + Quote(OutputPath + Filename), vbNormalFocus
    ' Show dialog form to reload from file or cancel
    ExternalEditorForm.Show
    Exit Sub
    
ShellError:
    MsgBox "Error Launching External Editor." & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
    Exit Sub
End Sub

Private Sub ButtonDataImport_Click()
    ' import data file into form
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    Dim vrtSelectedItem As Variant
    fd.AllowMultiSelect = True
    fd.InitialFileName = "*.*"
    fd.Filters.Clear
    fd.ButtonName = "Import"
    'fd.Filters.Add NameFilterDesc
    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            'import the file here'
            Debug.Print vrtSelectedItem
            Dim fs
            Dim DataFile As Object
            Const ForReading = 1, ForWriting = 2, ForAppending = 3
            Set fs = CreateObject("Scripting.FileSystemObject")
            If fs.FileExists(vrtSelectedItem) Then
                ListBoxDataFiles.AddItem fs.GetFileName(vrtSelectedItem)
                Set DataFile = fs.OpenTextFile(vrtSelectedItem, ForReading)
                GlobalDataFiles.Add fs.GetFileName(vrtSelectedItem), DataFile.ReadAll
                DataFile.Close
            End If
        Next vrtSelectedItem
    End If
    Set fd = Nothing

End Sub

Private Sub ButtonDataRename_Click()
    ' rename existsing datafile
    Dim sel_index As Integer
    Dim old_name As String
    Dim new_name As String
    sel_index = ListBoxDataFiles.ListIndex
    If sel_index <> -1 Then
        old_name = ListBoxDataFiles.Text
        new_name = InputBox("Enter new name for the data file.", "Rename data file", old_name)
        If new_name = old_name Or new_name = "" Then
            Exit Sub
        End If
        GlobalDataFiles.Add new_name, GlobalDataFiles.Item(old_name)
        GlobalDataFiles.Remove old_name
        ListBoxDataFiles.RemoveItem sel_index
        ListBoxDataFiles.AddItem new_name
    End If
End Sub

Private Sub ButtonDataRemove_Click()
    ' remove existing datafile
    Dim sel_index As Integer
    sel_index = ListBoxDataFiles.ListIndex
    If sel_index <> -1 Then
        GlobalDataFiles.Remove ListBoxDataFiles.Text
        ListBoxDataFiles.RemoveItem Index
    End If
End Sub


Private Sub CmdButtonEditorFontDown_Click()
    If TextBoxGLECode.Font.size > 4 Then
        TextBoxGLECode.Font.size = TextBoxGLECode.Font.size - 1
    End If
    SetValue EDITOR_FONT_SIZE_VALUE_NAME, TextBoxGLECode.Font.size
End Sub

Private Sub CmdButtonEditorFontUp_Click()
    If TextBoxGLECode.Font.size < 72 Then
        TextBoxGLECode.Font.size = TextBoxGLECode.Font.size + 1
    End If
    SetValue EDITOR_FONT_SIZE_VALUE_NAME, TextBoxGLECode.Font.size
End Sub

Private Sub ToggleButtonWrap_Click()
    If ToggleButtonWrap.value = True Then
        TextBoxGLECode.WordWrap = True
    Else
        TextBoxGLECode.WordWrap = False
    End If
End Sub


Private Sub MoveAnimation(oldshape As Shape, newShape As Shape)
    ' Move the animation settings of oldShape to newShape
    With ActiveWindow.Selection.SlideRange(1).TimeLine
        Dim eff As Effect
        For Each eff In .MainSequence
            If eff.Shape.name = oldshape.name Then eff.Shape = newShape
        Next
    End With
End Sub

Private Sub MatchZOrder(oldshape As Shape, newShape As Shape)
    ' Make the Z order of newShape equal to 1 higher than that of oldShape
    newShape.ZOrder msoBringToFront
    While (newShape.ZOrderPosition > oldshape.ZOrderPosition + 1)
        newShape.ZOrder msoSendBackward
    Wend
End Sub

Private Sub DeleteAnimation(oldshape As Shape)
    ' Delete the animation settings of oldShape
    With ActiveWindow.Selection.SlideRange(1).TimeLine
        For i = .MainSequence.count To 1 Step -1
            Dim eff As Effect
            Set eff = .MainSequence(i)
            If eff.Shape.name = oldshape.name Then eff.Delete
        Next
    End With
End Sub

Private Sub TransferGroupFormat(oldshape As Shape, newShape As Shape)
    On Error Resume Next
    ' Transfer group formatting
    If oldshape.Glow.Radius > 0 Then
        newShape.Glow.Color = oldshape.Glow.Color
        newShape.Glow.Radius = oldshape.Glow.Radius
        newShape.Glow.Transparency = oldshape.Glow.Transparency
    End If
    If oldshape.Reflection.Type <> msoReflectionTypeNone Then
        newShape.Reflection.Blur = oldshape.Reflection.Blur
        newShape.Reflection.Offset = oldshape.Reflection.Offset
        newShape.Reflection.size = oldshape.Reflection.size
        newShape.Reflection.Transparency = oldshape.Reflection.Transparency
        newShape.Reflection.Type = oldshape.Reflection.Type
    End If
    
    If oldshape.SoftEdge.Type <> msoSoftEdgeTypeNone Then
        newShape.SoftEdge.Radius = oldshape.SoftEdge.Radius
    End If
    
    If oldshape.Shadow.Visible Then
        newShape.Shadow.Visible = oldshape.Shadow.Visible
        newShape.Shadow.Blur = oldshape.Shadow.Blur
        newShape.Shadow.ForeColor = oldshape.Shadow.ForeColor
        newShape.Shadow.OffsetX = oldshape.Shadow.OffsetX
        newShape.Shadow.OffsetY = oldshape.Shadow.OffsetY
        newShape.Shadow.RotateWithShape = oldshape.Shadow.RotateWithShape
        newShape.Shadow.size = oldshape.Shadow.size
        newShape.Shadow.Style = oldshape.Shadow.Style
        newShape.Shadow.Transparency = oldshape.Shadow.Transparency
        newShape.Shadow.Type = oldshape.Shadow.Type
    End If
    
    If oldshape.ThreeD.Visible Then
        'newShape.ThreeD.BevelBottomDepth = oldshape.ThreeD.BevelBottomDepth
        'newShape.ThreeD.BevelBottomInset = oldshape.ThreeD.BevelBottomInset
        'newShape.ThreeD.BevelBottomType = oldshape.ThreeD.BevelBottomType
        'newShape.ThreeD.BevelTopDepth = oldshape.ThreeD.BevelTopDepth
        'newShape.ThreeD.BevelTopInset = oldshape.ThreeD.BevelTopInset
        'newShape.ThreeD.BevelTopType = oldshape.ThreeD.BevelTopType
        'newShape.ThreeD.ContourColor = oldshape.ThreeD.ContourColor
        'newShape.ThreeD.ContourWidth = oldshape.ThreeD.ContourWidth
        'newShape.ThreeD.Depth = oldshape.ThreeD.Depth
        'newShape.ThreeD.ExtrusionColor = oldshape.ThreeD.ExtrusionColor
        'newShape.ThreeD.ExtrusionColorType = oldshape.ThreeD.ExtrusionColorType
        newShape.ThreeD.Visible = oldshape.ThreeD.Visible
        newShape.ThreeD.Perspective = oldshape.ThreeD.Perspective
        newShape.ThreeD.FieldOfView = oldshape.ThreeD.FieldOfView
        newShape.ThreeD.LightAngle = oldshape.ThreeD.LightAngle
        'newShape.ThreeD.ProjectText = oldshape.ThreeD.ProjectText
        'If oldshape.ThreeD.PresetExtrusionDirection <> msoPresetExtrusionDirectionMixed Then
        '    newShape.ThreeD.SetExtrusionDirection oldshape.ThreeD.PresetExtrusionDirection
        'End If
        newShape.ThreeD.PresetLighting = oldshape.ThreeD.PresetLighting
        If oldshape.ThreeD.PresetLightingDirection <> msoPresetLightingDirectionMixed Then
            newShape.ThreeD.PresetLightingDirection = oldshape.ThreeD.PresetLightingDirection
        End If
        If oldshape.ThreeD.PresetLightingSoftness <> msoPresetLightingSoftnessMixed Then
            newShape.ThreeD.PresetLightingSoftness = oldshape.ThreeD.PresetLightingSoftness
        End If
        If oldshape.ThreeD.PresetMaterial <> msoPresetMaterialMixed Then
            newShape.ThreeD.PresetMaterial = oldshape.ThreeD.PresetMaterial
        End If
        If oldshape.ThreeD.PresetCamera <> msoPresetCameraMixed Then
            newShape.ThreeD.SetPresetCamera oldshape.ThreeD.PresetCamera
        End If
        newShape.ThreeD.RotationX = oldshape.ThreeD.RotationX
        newShape.ThreeD.RotationY = oldshape.ThreeD.RotationY
        newShape.ThreeD.RotationZ = oldshape.ThreeD.RotationZ
        'newShape.ThreeD.Z = oldshape.ThreeD.Z
    End If
End Sub

'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    ' If CloseMode = vbFormControlMenu Then
'        ' Cancel = True
'        ' ButtonCancel_Click
'    ' End If
'End Sub

Private Sub MultiPage1_Change()
    Call ToggleInputMode
End Sub


Private Sub ToggleInputMode()
    Set fs = CreateObject("Scripting.FileSystemObject")
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Select Case MultiPage1.value
        Case 0 ' Direct input
            TextBoxGLECode.SetFocus
        Case 1 ' Read from file
            ListBoxDataFiles.SetFocus
            ' ButtonLoadFile.Enabled = fs.FileExists(TextBoxFile.Text)
    End Select
    Call UserForm_Resize
End Sub

' Mousewheel functions

Private Sub TextBoxGLECode_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                        ByVal X As Single, ByVal Y As Single)
    If Not Me Is Nothing Then
        HookListBoxScroll Me, Me.TextBoxGLECode
    End If
End Sub




' Attempt at getting DPI of previous display, but I cannot find a way to retrieve
' that info for an embedded display, as there does not seem to be a way to load
' the display as an ImageFile Object
' Requires Microsoft Windows Image Acquisition Library
'Private Function GetImageFileDPI(fileNm As String) As Long
'    Dim imgFile As Object
'    Set imgFile = CreateObject("WIA.ImageFile")
'    imgFile.LoadFile (fileNm)
'    GetImageFileDPI = 96
'    If imgFile.HorizontalResolution <> "" Then
'        GetImageFileDPI = Round(imgFile.HorizontalResolution)
'    End If
'End Function

'Private Sub TextBoxTemplateCode_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
 '                       ByVal X As Single, ByVal Y As Single)
'    If Not Me Is Nothing Then
 '       HookListBoxScroll Me, Me.TextBoxTemplateCode
'    End If
'End Sub



' v1.58: I'm removing this because the support is not great, and I don't think scrolling is very useful
' for this combobox. The issue is that the combobox is within a frame, and once it gets the hook, we cannot
' unhook until we leave the whole frame, not just the combobox.
'Private Sub ComboBoxLaTexEngine_MouseMove( _
'                        ByVal Button As Integer, ByVal Shift As Integer, _
'                        ByVal X As Single, ByVal Y As Single)
'    If Not Me Is Nothing Then
'         HookListBoxScroll Me, Me.ComboBoxLaTexEngine
'    End If
'End Sub

' It seems difficult to get good mouse wheel support simultaneously for the Bitmap/Vector combobox
' and the LatexEngine combobox, because they are in the same frame, and whoever gets the hook first holds to it.
'Private Sub ComboBoxBitmapVector_MouseMove( _
'                        ByVal Button As Integer, ByVal Shift As Integer, _
'                        ByVal X As Single, ByVal Y As Single)
'    If Not Me Is Nothing Then
'         HookListBoxScroll Me, Me.ComboBoxBitmapVector
'    End If
'End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        UnhookListBoxScroll
End Sub



' Sub AddMenuItem(itemText As String, itemCommand As String, itemFaceId As Long)
'     ' Check if we have already added the menu item
'     Dim initialized As Boolean
'     Dim bef As Integer
'     initialized = False
'     bef = 1
'     Dim Menu As CommandBars
'     Set Menu = Application.CommandBars
'     For i = 1 To Menu("Insert").Controls.count
'         With Menu("Insert").Controls(i)
'             If .Caption = itemText Then
'                 initialized = True
'                 Exit For
'             ElseIf InStr(.Caption, "Dia&gram") Then
'                 bef = i
'             End If
'         End With
'     Next
    
'     ' Create the menu choice.
'     If Not initialized Then
'         Dim NewControl As CommandBarControl
'         Set NewControl = Menu("Insert").Controls.Add _
'                               (Type:=msoControlButton, _
'                                before:=bef, _
'                                Id:=itemFaceId)
'         NewControl.Caption = itemText
'         NewControl.OnAction = itemCommand
'         NewControl.Style = msoButton
'     End If
' End Sub

 Sub UnInitializeApp()

 ' End Sub
'     'RemoveMenuItem "New GLE display..."
'     'RemoveMenuItem "Edit GLE display..."
'     'RemoveMenuItem "Regenerate selection..."
'     'RemoveMenuItem "Vectorize selection..."
'     'RemoveMenuItem "Rasterize selection..."
'     'RemoveMenuItem "Settings..."
'     'RemoveMenuItem "Insert vector file..."
'     ' Clean up older versions
'     'RemoveMenuItem "Regenerate selected displays..."
'     'RemoveMenuItem "Convert to EMF..."
'     'RemoveMenuItem "Convert to PNG..."
 End Sub

' Sub RemoveMenuItem(itemText As String)
'     Dim Menu As CommandBars
'     Set Menu = Application.CommandBars
'     For i = 1 To Menu("Insert").Controls.count
'         If Menu("Insert").Controls(i).Caption = itemText Then
'             Menu("Insert").Controls(i).Delete
'             Exit For
'         End If
'     Next
' End Sub
