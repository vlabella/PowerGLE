Attribute VB_Name = "Config"
'
' -- Config.bas
'
' Global Config file. Public constants and default settings
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
'
' -- Constants
'
Public Const POWER_GLE_VERSION_NUMBER As String = "1.0.0"
Public Const POWER_GLE_UUID = "c0a8eff3-9149-4dd4-85e5-4cb87433ebf2" ' all shapes must have this UUID to be identified as a PowerGLE figure
Public Const POWER_GLE_URL As String = "https://github.com/vlabella/PowerGLE"
Public Const POWER_GLE_REGISTRY_PATH As String = "Software\PowerGLE"
Public Const TEMP_FILENAME As String = "figure"
Public Const GLE_EXT As String = "gle"
Public Const GLE_URL As String = "http://glx.sourceforge.io"
Public Const EXTERNAL_EDITOR_URL As String = "http://www.sublimetext.com"
' (arrays cannot be constants so CSV lists are utilized and split using ArrayFromCSVString
Public Const OUTPUT_FORMATS As String = "PNG,JPEG,TIFF"
Public Const OUTPUT_FORMAT_FILE_EXT As String = "png,jpeg,tiff"
Public Const GLE_FORM_MIN_HEIGHT As Integer = 350
Public Const GLE_FORM_MIN_WIDTH As Integer = 250
Public Const ENDL As String = vbCrLf
'
' -- User changeable settings defaults
'
Public Const DEFAULT_GLE_EXECUTABLE As String = "C:\Program Files\gle\bin\gle.exe"
Public Const DEFAULT_BITMAP_DPI As Integer = 250
Public Const DEFAULT_OUTPUT_FORMAT As String = "PNG"
Public Const DEFAULT_USE_CAIRO As Boolean = True
Public Const DEFAULT_PRESERVE_TEMP_FILES As Boolean = True
Public Const DEFAULT_PNG_TRANSPARENT As Boolean = False
Public Const DEFAULT_USE_UTF8 As Boolean = True
Public Const DEFAULT_EDITOR_FONT_SIZE As Integer = 10
Public Const DEFAULT_TIMEOUT As Integer = 30
Public Const DEFAULT_EDITOR_WORD_WRAP As Boolean = False
Public Const DEFAULT_ABSOLUTE_TEMP_DIR As String = "C:\temp\PowerGLE"
Public Const DEFAULT_RELATIVE_TEMP_DIR As String = ".\PowerGLE"
Public Const DEFAULT_USE_ABSOLUTE_TEMP_DIR As Boolean = False
Public Const DEFAULT_EXTERNAL_EDITOR_EXECUTABLE As String = "C:\Program Files\Sublime Text\sublime_text.exe"
Public Const DEFAULT_USE_EXTERNAL_EDITOR As Boolean = False
Public Const DEFAULT_SCALING_GAIN As Double = 1#
Public Const DEFAULT_GLE_FORM_HEIGHT As Integer = 312
Public Const DEFAULT_GLE_FORM_WIDTH As Integer = 400
Public Const DEFAULT_DEBUG As Boolean = False
Public Const DEFAULT_SLIDE_POSTIION_X As Integer = 100
Public Const DEFAULT_SLIDE_POSTIION_Y As Integer = 100
Public Const DEFAULT_GLE_CODE As String = "size 10 10" + ENDL + "set font texcmss" + ENDL + "set hei 0.5" + ENDL + "amove 0 0" + ENDL
'
' -- Registry names for user changeable settings
'
Public Const GLE_EXECUTABLE_VALUE_NAME = "GLE_EXECUTABLE"
Public Const INITIAL_SOURCECODE_VALUE_NAME = "INITIAL_SOURCECODE"
Public Const SOURCECODE_CURSOR_POSITION_VALUE_NAME = "CURSOR_POSITION"
Public Const BITMAP_DPI_VALUE_NAME = "BITMAP_DPI"
Public Const OUTPUT_FORMAT_VALUE_NAME = "BITMAP_FORMAT"
Public Const USE_CAIRO_VALUE_NAME = "USE_CAIRO"
Public Const PRESERVE_TEMP_FILES_VALUE_NAME = "PRESERVE_TEMP_FILES"
Public Const PNG_TRANSPARENT_VALUE_NAME = "PNG_TRANSPARENT"
Public Const ABSOLUTE_TEMP_DIR_VALUE_NAME = "ABSOLUTE_TEMP_DIR"
Public Const USE_ABSOLUTE_TEMP_DIR_VALUE_NAME = "USE_ABSOLUTE_TEMP_DIR"
Public Const RELATIVE_TEMP_DIR_VALUE_NAME = "RELATIVE_TEMP_DIR"
Public Const PNGTRANSPARENT_VALUE_NAME = "PNG_TRANSPARENT" ' cannot be PNG_TRANSPARENT for some reason
Public Const EDITOR_FONT_SIZE_VALUE_NAME = "EDITOR_FONT_SIZE"
Public Const EDITOR_WORD_WRAP_VALUE_NAME = "EDITOR_WORD_WRAP"
Public Const EXTERNAL_EDITOR_EXECUTABLE_VALUE_NAME = "EXTERNAL_EDITOR_EXECUTABLE"
Public Const USE_UTF8_VALUE_NAME = "USE_UTF8"
Public Const TIMEOUT_VALUE_NAME = "TIMEOUT"
Public Const USE_EXTERNAL_EDITOR_VALUE_NAME = "USE_EXTERNAL_EDITOR"
Public Const SCALING_GAIN_VALUE_NAME = "SCALING_GAIN"
Public Const GLE_FORM_HEIGHT_VALUE_NAME = "GLE_FORM_HEIGHT"
Public Const GLE_FORM_WIDTH_VALUE_NAME = "GLE_FORM_WIDTH"
Public Const DEFAULT_SLIDE_POSTIION_X_VALUE_NAME = "DEFAULT_SLIDE_POSTIION_X"
Public Const DEFAULT_SLIDE_POSTIION_Y_VALUE_NAME = "DEFAULT_SLIDE_POSTIION_Y"
Public Const DEBUG_VALUE_NAME = "DEBUG"
'Public Const VALUE_NAME = ""
'
' - Shape Tag names for saving in PowerPoint file
'
Public Const SHAPE_TAG_PREFIX = "POWER_GLE" ' all tags are prefixed with this
Public Const TAG_FIGURE = "FIGURE"
Public Const TAG_FIGURE_UUID = "FIGURE_UUID"
Public Const TAG_OUTPUT_DPI = "OUTPUT_DPI"
Public Const TAG_OUTPUT_FORMAT = "OUTPUT_FORMAT"
Public Const TAG_VERSION = "VERSION"
Public Const TAG_SOURCE_CODE = "SOURCE_CODE"
Public Const TAG_FIGURE_NAME = "FIGURE_NAME"
Public Const TAG_TEMP_FOLDER = "TEMP_FOLDER"
Public Const TAG_DATA_FILENAME = "DATA_FILENAME"
Public Const TAG_DATA_FILE_CONTENT = "DATA_FILE_CONTENT"
Public Const TAG_ORIGINAL_HEIGHT = "ORIGINAL_HEIGHT"
Public Const TAG_ORIGINAL_WIDTH = "ORIGINAL_WIDTH"
Public Const TAG_SLIDE_INDEX = "SLIDE_INDEX"

Public Function GetShapeTagName(name As String) As String
    ' returns shape tag name with prefix
    GetShapeTagName = SHAPE_TAG_PREFIX + "_" + name
End Function
'
' Get and Set interface for user settings -> read/write to registry
' GetValue reads from registry or default if registry value does not exist
' SetValue saves to registry
' GetDefault is a look up table since hashmaps dont exists in VBA - or dictionaries cannot be const
' calling code must know value name of setting defined above as constants
' value type in registry is determined by VBA type
'
Public Function GetDefaultValue(name As String) As Variant
    ' look up table
    GetDefaultValue = ""
    If name = GLE_EXECUTABLE_VALUE_NAME Then
        GetDefaultValue = DEFAULT_GLE_EXECUTABLE
    ElseIf name = INITIAL_SOURCECODE_VALUE_NAME Then
        GetDefaultValue = DEFAULT_GLE_CODE
    ElseIf name = SOURCECODE_CURSOR_POSITION_VALUE_NAME Then
        GetDefaultValue = Len(DEFAULT_GLE_CODE)
    ElseIf name = BITMAP_DPI_VALUE_NAME Then
        GetDefaultValue = DEFAULT_BITMAP_DPI
    ElseIf name = OUTPUT_FORMAT_VALUE_NAME Then
        GetDefaultValue = DEFAULT_OUTPUT_FORMAT
    ElseIf name = USE_CAIRO_VALUE_NAME Then
        GetDefaultValue = DEFAULT_USE_CAIRO
    ElseIf name = PRESERVE_TEMP_FILES_VALUE_NAME Then
        GetDefaultValue = DEFAULT_PRESERVE_TEMP_FILES
    ElseIf name = USE_ABSOLUTE_TEMP_DIR_VALUE_NAME Then
        GetDefaultValue = DEFAULT_USE_ABSOLUTE_TEMP_DIR
    ElseIf name = ABSOLUTE_TEMP_DIR_VALUE_NAME Then
        GetDefaultValue = DEFAULT_ABSOLUTE_TEMP_DIR
    ElseIf name = RELATIVE_TEMP_DIR_VALUE_NAME Then
        GetDefaultValue = DEFAULT_RELATIVE_TEMP_DIR
    ElseIf name = PNGTRANSPARENT_VALUE_NAME Then
        GetDefaultValue = DEFAULT_PNG_TRANSPARENT
    ElseIf name = EDITOR_FONT_SIZE_VALUE_NAME Then
        GetDefaultValue = DEFAULT_EDITOR_FONT_SIZE
    ElseIf name = EDITOR_WORD_WRAP_VALUE_NAME Then
        GetDefaultValue = DEFAULT_EDITOR_WORD_WRAP
    ElseIf name = EXTERNAL_EDITOR_EXECUTABLE_VALUE_NAME Then
        GetDefaultValue = DEFAULT_EXTERNAL_EDITOR_EXECUTABLE
    ElseIf name = USE_UTF8_VALUE_NAME Then
        GetDefaultValue = DEFAULT_USE_UTF8
    ElseIf name = USE_EXTERNAL_EDITOR_VALUE_NAME Then
        GetDefaultValue = DEFAULT_USE_EXTERNAL_EDITOR
    ElseIf name = SCALING_GAIN_VALUE_NAME Then
        GetDefaultValue = DEFAULT_SCALING_GAIN
    ElseIf name = TIMEOUT_VALUE_NAME Then
        GetDefaultValue = DEFAULT_TIMEOUT
    ElseIf name = GLE_FORM_HEIGHT_VALUE_NAME Then
        GetDefaultValue = DEFAULT_GLE_FORM_HEIGHT
    ElseIf name = GLE_FORM_WIDTH_VALUE_NAME Then
        GetDefaultValue = DEFAULT_GLE_FORM_WIDTH
    ElseIf name = DEBUG_VALUE_NAME Then
        GetDefaultValue = DEFAULT_DEBUG
    ElseIf name = SLIDE_POSITION_X_VALUE_NAME Then
        GetDefaultValue = DEFAULT_SLIDE_POSITION_X
    ElseIf name = SLIDE_POSITION_Y_VALUE_NAME Then
        GetDefaultValue = DEFAULT_SLIDE_POSITION_Y
    End If
End Function

Public Function GetValue(name As String) As Variant
    GetValue = ""
    Dim DefaultValue As Variant
    DefaultValue = GetDefaultValue(name)
    If DefaultValue <> "" Then
        GetValue = GetRegistryValue(HKEY_CURRENT_USER, POWER_GLE_REGISTRY_PATH, name, DefaultValue)
    End If
End Function

Public Sub SetValue(name As String, value As Variant)
    Dim ValueType As Long
    ValueType = 0
    Select Case VarType(value)
    Case vbInteger, vbLong
        ValueType = REG_DWORD
    Case vbBoolean
        ValueType = REG_DWORD
        value = BoolToInt(CBool(value))
    Case vbSingle, vbDouble:
        ValueType = REG_SZ
        value = CStr(value)
    Case vbString:
        ValueType = REG_SZ
    End Select
    If ValueType <> 0 Then
        SetRegistryValue HKEY_CURRENT_USER, POWER_GLE_REGISTRY_PATH, name, ValueType, value
    End If
End Sub

Public Function BoolToInt(value As Boolean) As Long
    ' so the bool gets stored as a 1 or 0 in the registry instead of 0XFFFFFFFF and 0
    BoolToInt = 0&
    If value Then
        BoolToInt = 1&
    End If
End Function

Public Function BoolToString(val) As String
    BoolToString = "0"
    If val Then
        BoolToString = "1"
    End If
End Function

Public Function StringToBool(val) As Boolean
    StringToBool = True
    If val = "0" Then
        StringToBool = False
    End If
End Function





