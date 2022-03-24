Attribute VB_Name = "CommonRoutines"
'
' -- CommonRoutines.bas  -  collection of handy VBA routines
'
' Author: Vincent LaBella
' vlabella@sunypoly.edu
'
' Some routines based on code from IguanaTeX  www.jonathanleroux.org/software/iguanatex/
'
Public Function GetTempPath(Optional PrependPresentation As Boolean = True) As String
    ' Return an absolute path to temp location
    ' if relative path is selected the active presentation path will be prepended if PrependPresentation is True (default)
    Dim res As String
    Dim UseAbsolutev As Variant
    UseAbsolutev = GetValue(USE_ABSOLUTE_TEMP_DIR_VALUE_NAME)
    Dim UseAbsolute As Boolean
    UseAbsolute = CBool(UseAbsolutev)
    res = CStr(GetValue(ABSOLUTE_TEMP_DIR_VALUE_NAME))
    If UseAbsolute = False Then
        res = CStr(GetValue(RELATIVE_TEMP_DIR_VALUE_NAME))
        If PrependPresentation Then
            ' relative prepend the current presentation path to it
            Dim sPath As String
            sPath = ActivePresentation.path
            If Len(sPath) > 0 Then
                res = AddSlash(sPath) + res
            Else
                MsgBox "The current presentation must be saved once prior to adding a GLE figure."
                GetTempPath = ""
                Exit Function
            End If
        End If
    End If
    ' add presentation name and change . to \_
    res = AddSlash(res) + Replace(ActivePresentation.name, ".", "_")
    GetTempPath = AddSlash(res)
End Function

Public Function GetFigureName(Optional Index As Long = 0) As String
    ' returns the default figure name from the index provided
    If Index = 0 Then
        GetFigureName = TEMP_FILENAME + "_"
    Else
        GetFigureName = TEMP_FILENAME + "_" + CStr(Index)
    End If
End Function

Public Function GetNextFigureName(strPath As String) As String
    ' gets the next available figure name from the path provided
    ' each figure name is the folder name as well'
    ' FIGURE_NAME_XXX  where XXX is a number e.g. figure_1 figure_2
    GetNextFigureName = GetFigureName(GetNextFigureIndexInFolder(strPath))
End Function

Public Function GetNextFigureIndexInFolder(strPath As String) As Long
    ' gets the next available figure index in folder
    ' FIGURE_NAME_XXX  where XXX is a number e.g. figure_1 figure_2
    Dim max_index As Integer
    Dim new_index As Integer
    Dim RegEx As Object
    Dim allMatches As Object
    max_index = 1
    Set xFSO = CreateObject("Scripting.FileSystemObject")
    Set xFolder = xFSO.GetFolder(strPath)
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.Global = True
    RegEx.IgnoreCase = True
    With RegEx
        .Pattern = "(^" + GetFigureName() + "([0-9]+)$)"
    End With
    For Each fold In xFolder.SubFolders
        ' get number part of folder if it matches
        Set allMatches = RegEx.Execute(fold.name)
        If allMatches.count <> 0 Then
            new_index = CInt(allMatches.Item(0).submatches.Item(1))
            If max_index <= new_index Then
                max_index = new_index + 1
            End If
        End If
    Next fold
    GetNextFigureIndexInFolder = max_index
End Function

Public Function GetNextFigureIndexInShapes(Shapes() As Variant) As Long
    ' gets the next available figure index in folder
    ' FIGURE_NAME_XXX  where XXX is a number e.g. figure_1 figure_2
    Dim max_index As Integer
    Dim new_index As Integer
    Dim RegEx As Object
    Dim allMatches As Object
    max_index = 1
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.Global = True
    RegEx.IgnoreCase = True
    With RegEx
        .Pattern = "(^" + GetFigureName() + "([0-9]+)$)"
    End With
    For Each s In Shapes
        ' get number part of figure name if it matches
        Set allMatches = RegEx.Execute(s.Tags(GetShapeTagName(TAG_FIGURE_NAME)))
        If allMatches.count <> 0 Then
            new_index = CInt(allMatches.Item(0).submatches.Item(1))
            If max_index <= new_index Then
                max_index = new_index + 1
            End If
        End If
    Next
    GetNextFigureIndexInShapes = max_index
End Function


Public Sub CreateFolder(strPath As String)
    ' recursively creates folders even if parents do not exist
    Dim elm As Variant
    Dim strCheckPath As String
    strCheckPath = ""
    For Each elm In Split(strPath, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
End Sub

Public Sub SaveTextFile(OutputPath As String, Filename As String, Content As String, Optional UseUTF8 As Boolean = True, Optional Overwrite As Boolean = True)
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(OutputPath + Filename) Then
        If Overwrite = True Then
            fs.DeleteFile OutputPath + Filename
        Else
            Exit Sub
        End If
    End If
    
    If UseUTF8 = False Then
        Set f = fs.CreateTextFile(OutputPath + Filename, True)
        f.Write Content
        f.Close
    Else
        Dim BinaryStream As Object
        Set BinaryStream = CreateObject("ADODB.stream")
        BinaryStream.Type = 1
        BinaryStream.Open
        Dim adodbStream  As Object
        Set adodbStream = CreateObject("ADODB.Stream")
        With adodbStream
            .Type = 2 'Stream type
            .Charset = "utf-8"
            .Open
            .WriteText Content
            '.SaveToFile OutputPath & FilePrefix & ".gle", 2 'Save binary data To disk; problem: this includes a BOM
            ' Workaround to avoid BOM in file:
            .Position = 3 'skip BOM
            .CopyTo BinaryStream
            .Flush
            .Close
        End With
        BinaryStream.SaveToFile OutputPath + Filename, 2 'Save binary data To disk
        BinaryStream.Flush
        BinaryStream.Close
    End If
    Set fs = Nothing
End Sub

Public Function ArrayFromCSVList(value As String) As String()
    ArrayFromCSVList = Split(value, ",")
End Function

Public Function GetArrayIndex(a As String, value As String) As Integer
    GetArrayIndex = 0
    Dim la() As String
    la = ArrayFromCSVList(a)
    Dim i As Integer
    For i = LBound(la) To UBound(la)
        If la(i) = value Then
            GetArrayIndex = i
            Exit For
        End If
    Next
End Function

Public Function GetArrayValue(a As String, Index As Integer) As String
    GetArrayValue = ""
    Dim la() As String
    la = ArrayFromCSVList(a)
    If Index >= LBound(la) And Index <= UBound(la) Then
        GetArrayValue = la(Index)
    End If
End Function

' Public Function GetUUID(Optional lowercase As Boolean, Optional parens As Boolean) As String
' ' not good may produce non unique ids
'     Dim k As Integer
'     Dim h As String
'     GetUUID = Space(36)
'     For k = 1 To Len(GetUUID)
'         Randomize
'         Select Case k
'             Case 9, 14, 19, 24: h = "-"
'             Case 15:            h = "4"
'             Case 20:            h = Hex(Rnd * 3 + 8)
'             Case Else:          h = Hex(Rnd * 15)
'         End Select
'         Mid(GetUUID, k, 1) = h
'     Next
'     If lowercase Then GetUUID = LCase(GetUUID)
'     If parens Then GetUUID = "{" & GetUUID & "}"
' End Function

Public Function IsPathWritable(TempPath As String) As Boolean
    Dim FName As String
    Dim FHdl As Integer
    FName = TempPath & GetUUID()
    On Error GoTo TempFolderNotWritable
    FHdl = FreeFile()
    Open FName For Output Access Write As FHdl
    Print #FHdl, "TESTWRITE"
    Close FHdl
    IsPathWritable = True
    Kill FName
    On Error GoTo 0
    Exit Function

TempFolderNotWritable:
    IsPathWritable = False
End Function

Public Function IsInArray(arr As Variant, valueToCheck As String) As Boolean
    IsInArray = False
    For Each n In arr
        If n = valueToCheck Then
            IsInArray = True
            Exit For
        End If
    Next
End Function

Public Function max(X, Y As Variant) As Variant
  max = IIf(X > Y, X, Y)
End Function

Public Function min(X, Y As Variant) As Variant
   min = IIf(X < Y, X, Y)
End Function

Public Function Quote(val As String) As String
    ' returns quoted string
    Quote = """" + val + """"
End Function

Public Function UnQuote(val As String) As String
    ' returns string with quotes removed
    If Left(val, 1) = """" Then val = Mid(val, 2, Len(val) - 1)
    If Right(val, 1) = """" Then val = Left(val, Len(val) - 1)
    UnQuote = val
End Function

Public Function AddSlash(val As String) As String
    ' add slash to end of string if needed
    AddSlash = val
    If Right(val, 1) <> "\" Then
        AddSlash = val + "\"
    End If
End Function

Public Function PackArrayToString(vArray As Variant) As String
    Dim strDelimiter As String
    strDelimiter = "|"
    PackArrayToString = Join(vArray, strDelimiter)
End Function

Public Function UnpackStringToArray(Str As String) As Variant
    Dim strDelimiter As String
    strDelimiter = "|"
    UnpackStringToArray = Split(Str, strDelimiter, , vbTextCompare)
End Function

Function mergeArrays(arr1() As Variant, arr2() As Variant) As Variant
    ' merge two arrays into one see https://stackoverflow.com/questions/1588913/how-do-i-merge-two-arrays-in-vba?answertab=active#tab-top
    ' with some modifications if arrays nore not itialized (ie. empty)
    Dim returnThis() As Variant
    Dim len1 As Long, len2 As Long, lenRe As Long, counter As Long
    If Not IsArray(arr1) And Not IsArray(arr2) Then Exit Function
    len1 = 0
    len2 = 0
    'In VB, for whatever reason, Not myArray returns the SafeArray pointer. For uninitialized arrays, this returns -1.
    If (Not arr1) <> -1 Then
        ' array initialzed
        len1 = UBound(arr1)
    End If
    If (Not arr2) <> -1 Then
        ' array initialzed
        len2 = UBound(arr2)
    End If
    lenRe = len1 + len2
    If lenRe > 0 Then
        ReDim returnThis(1 To lenRe)
        counter = 1
        Do While counter <= len1 'copy first array into returnThis
            Set returnThis(counter) = arr1(counter)
            counter = counter + 1
        Loop
        If len2 > 0 Then
            Do While counter <= lenRe 'copy the second array in returnThis
                Set returnThis(counter) = arr2(counter - len1)
                counter = counter + 1
            Loop
        End If
    End If
    mergeArrays = returnThis
End Function

Public Sub FormAltTextandTitle(ByRef lshape As Variant)
    ' forms based on what is in tags
    lshape.AlternativeText = "PowerGLE Figure: " + lshape.Tags(GetShapeTagName(TAG_FIGURE_NAME)) + "." + GLE_EXT + " format: " + lshape.Tags(GetShapeTagName(TAG_OUTPUT_FORMAT)) + " dpi: " + lshape.Tags(GetShapeTagName(TAG_OUTPUT_DPI))
    lshape.Title = "PowerGLE Figure: " + lshape.Tags(GetShapeTagName(TAG_FIGURE_NAME)) + "." + GLE_EXT
End Sub

Public Sub FullyUngroupShape(newShape As Shape)
    Dim Shr As ShapeRange
    Dim s As Shape
    If newShape.Type = msoGroup Then
        Set Shr = newShape.Ungroup
        For i = 1 To Shr.count
            Set s = Shr.Item(i)
            If s.Type = msoGroup Then
                Call FullyUngroupShape(s)
            End If
        Next
    End If
End Sub

Public Function GetAllShapesInGroup(newShape As Shape, Optional PowerGLEShapesOnly As Boolean = True) As Variant
    ' returns array of shapes if single shape it will be a one element array
    Dim arr() As Variant
    Dim s As Shape
    If newShape.Type = msoGroup Then
        For Each s In newShape.GroupItems
            Dim b() As Variant
            b = GetAllShapesInGroup(s, PowerGLEShapesOnly)
            arr = mergeArrays(arr, b)
        Next
    Else
        If (PowerGLEShapesOnly = True And IsPowerGLEShape(newShape) = True) Or PowerGLEShapesOnly = False Then
            ReDim arr(1 To 1)
            Set arr(1) = newShape
        End If
    End If
    GetAllShapesInGroup = arr
End Function

Public Function GetAllShapesInShapeRange(newShapeRange As ShapeRange, Optional PowerGLEShapesOnly As Boolean = True) As Variant
    ' returns array of shapes if single shape it will be a one element array
    Dim arr() As Variant
    Dim j As Long
    Dim s As Shape
    For Each s In newShapeRange
        Dim b() As Variant
        b = GetAllShapesInGroup(s, PowerGLEShapesOnly)
        arr = mergeArrays(arr, b)
    Next
    GetAllShapesInShapeRange = arr
End Function


Public Function GetAllShapesOnSlide(newSlide As slide, Optional PowerGLEShapesOnly As Boolean = True) As Variant
    ' returns all shapes on slide in an array
    Dim arr() As Variant
    Dim s As Shape
    For Each s In newSlide.Shapes
        Dim b() As Variant
        b = GetAllShapesInGroup(s, PowerGLEShapesOnly)
        arr = mergeArrays(arr, b)
    Next
    GetAllShapesOnSlide = arr
End Function

Public Function GetAllShapesInPresentation(Pres As Presentation, Optional PowerGLEShapesOnly As Boolean = True) As Variant
    Dim arr() As Variant
    Dim s As slide
    For Each s In Pres.Slides
        Dim b() As Variant
        b = GetAllShapesOnSlide(s, PowerGLEShapesOnly)
        arr = mergeArrays(arr, b)
    Next
    GetAllShapesInPresentation = arr
End Function


' Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (guid As UUID_TYPE) As LongPtr
' Private Declare PtrSafe Function StringFromUUID2 Lib "ole32.dll" (guid As UUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr

' Private Type UUID_TYPE
'                 Data1 As Long
'                 Data2 As Integer
'                 Data3 As Integer
'                 Data4(7) As Byte
' End Type

' Public Function GetUUID() As String
'     Dim guid As UUID_TYPE
'     Dim strGuid As String
'     Dim retValue As LongPtr
'     Const guidLength As Long = 39 'registry UUID format with null terminator {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}
    
'     retValue = CoCreateGuid(guid)
'     If retValue = 0 Then
'       strGuid = String$(guidLength, vbNullChar)
'       retValue = StringFromUUID2(guid, StrPtr(strGuid), guidLength)
'       If retValue = guidLength Then
'          ' valid UUID as a string
'          GetUUID = Mid$(strGuid, 2, 36)
'       End If
'     End If
' End Function




