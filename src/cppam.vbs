'
' cppam.vbs  Create PowerPoint.pptm and ppam files for PowerGLE from a collection of vb scripts and forms
' author: Vincent LaBella vlabella@sunypoly.edu
' usage cscript cppam.vbs filename
'
Option Explicit

Sub WriteLine ( strLine )
    WScript.Stdout.WriteLine strLine
End Sub

' define  constants since vbscript cannot have library references
'
' http://msdn.microsoft.com/en-us/library/office/aa432714(v=office.12).aspx
Const msoFalse = 0   ' False.
Const msoTrue = -1   ' True.

' http://msdn.microsoft.com/en-us/library/office/bb265636(v=office.12).aspx
Const ppFixedFormatIntentScreen = 1 ' Intent is to view exported file on screen.
Const ppFixedFormatIntentPrint = 2  ' Intent is to print exported file.

' http://msdn.microsoft.com/en-us/library/office/ff746754.aspx
Const ppFixedFormatTypeXPS = 1  ' XPS format
Const ppFixedFormatTypePDF = 2  ' PDF format

' http://msdn.microsoft.com/en-us/library/office/ff744564.aspx
Const ppPrintHandoutVerticalFirst = 1   ' Slides are ordered vertically, with the first slide in the upper-left corner and the second slide below it.
Const ppPrintHandoutHorizontalFirst = 2 ' Slides are ordered horizontally, with the first slide in the upper-left corner and the second slide to the right of it.

' http://msdn.microsoft.com/en-us/library/office/ff744185.aspx
Const ppPrintOutputSlides = 1               ' Slides
Const ppPrintOutputTwoSlideHandouts = 2     ' Two Slide Handouts
Const ppPrintOutputThreeSlideHandouts = 3   ' Three Slide Handouts
Const ppPrintOutputSixSlideHandouts = 4     ' Six Slide Handouts
Const ppPrintOutputNotesPages = 5           ' Notes Pages
Const ppPrintOutputOutline = 6              ' Outline
Const ppPrintOutputBuildSlides = 7          ' Build Slides
Const ppPrintOutputFourSlideHandouts = 8    ' Four Slide Handouts
Const ppPrintOutputNineSlideHandouts = 9    ' Nine Slide Handouts
Const ppPrintOutputOneSlideHandouts = 10    ' Single Slide Handouts

' http://msdn.microsoft.com/en-us/library/office/ff745585.aspx
Const ppPrintAll = 1            ' Print all slides in the presentation.
Const ppPrintSelection = 2      ' Print a selection of slides.
Const ppPrintCurrent = 3        ' Print the current slide from the presentation.
Const ppPrintSlideRange = 4     ' Print a range of slides.
Const ppPrintNamedSlideShow = 5 ' Print a named slideshow.

' http://msdn.microsoft.com/en-us/library/office/ff744228.aspx
Const ppShowAll = 1             ' Show all.
Const ppShowNamedSlideShow = 3  ' Show named slideshow.
Const ppShowSlideRange = 2      ' Show slide range.

Const ppSaveAsOpenXMLAddin  =  30
Const ppSaveAsOpenXMLPicturePresentation = 36
Const ppSaveAsOpenXMLPresentation =24
Const ppSaveAsOpenXMLPresentationMacroEnabled =25
Const ppLayoutBlank = 12

Const vbext_ct_StdModule =  1'   Standard module
Const vbext_ct_ClassModule =   2'   Class module
Const vbext_ct_MSForm= 3  ' Microsoft Form
Const vbext_ct_ActiveXDesigner=    11'  ActiveX Designer
Const vbext_ct_Document  = 100' Document Module

Dim inputFile
Dim outputFile
Dim ppamFile
Dim pptmFile
Dim objPPT
Dim objPresentation
Dim objPrintOptions
Dim objFso
Dim oShell
Dim ofso
Dim CurrentDirectory

Set oShell = CreateObject("WScript.Shell")

If WScript.Arguments.Count <> 1 Then
    WriteLine "You need to specify input and output files."
    WScript.Quit
End If

ppamFile = WScript.Arguments(0)+".ppam"
pptmFile = WScript.Arguments(0)+".pptm"

WriteLine "Creating [" + pptmFile  + "] and [" + ppamFile + "]"

ppamFile = oShell.CurrentDirectory+"\"+ppamFile
pptmFile = oShell.CurrentDirectory+"\"+pptmFile

Set objFso = CreateObject("Scripting.FileSystemObject")

'If objFso.FileExists( outputFile ) Then
'    WriteLine "Your output file (' & outputFile & ') already exists!"
'    WScript.Quit
'End If

'WriteLine "Input File:  " & inputFile
'WriteLine "Output File: " & outputFile

' build presenation'
if 1 = 1 then
Set objPPT = CreateObject( "PowerPoint.Application" )
Dim NewPres
Dim Slide
Dim TextBox
Set NewPres = objPPT.Presentations.Add
Set Slide = NewPres.Slides.Add(1, 16)
Set TextBox = Slide.Shapes.Item(1)
TextBox.TextFrame.TextRange.Text = "PowerGLE - PowerPoint Add-in for GLE (glx.sourceforge.net)"
Set TextBox = Slide.Shapes.Item(2)
TextBox.Delete

Dim sh 
Set sh = Slide.Shapes.AddPicture( oShell.CurrentDirectory+"\"+"logo.png", msoFalse, msoTrue, 0, 0)
sh.Top = (NewPres.PageSetup.SlideHeight - sh.Height) / 2
sh.Left = (NewPres.PageSetup.SlideWidth - sh.Width) / 2
' import macros and save file'
Dim files(13)
files(0) = "Config.bas"
files(1) = "CommonRoutines.bas"
files(2) = "PowerGLE.bas"
files(3) = "WinLibRoutines.bas"
files(4) = "SettingsForm.frm"
files(5) = "ExportVBA.bas"
files(6) = "AboutBox.frm"
files(7) = "BatchEditForm.frm"
files(8) = "ExternalEditorForm.frm"
files(9) = "GLEForm.frm"
files(10) = "LogFileViewer.frm"
files(11) = "RegenerateForm.frm"
files(12) = "AppEventHandler.cls"


Dim file
For Each file In files
    WriteLine file
    if file <> "" then 
        NewPres.VBProject.VBComponents.Import oShell.CurrentDirectory+"\"+file
    end if
Next


NewPres.VBProject.Name = "PowerGLEAddin"
NewPres.VBProject.Description = "PowerPoint Add-in for GLE. glx.sourceforge.net"
' add reference to the Microsoft Scripting Runtime version 1.0
NewPres.VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0

'
' add the custom ribbon
'
' objPPT.IRibbonExtensibility.GetCustomUI  oShell.CurrentDirectory+"\"+customui.xml
'NewPres.SaveAs ppamFile , ppSaveAsOpenXMLAddin

NewPres.SaveAs pptmFile , ppSaveAsOpenXMLPresentationMacroEnabled
NewPres.Close
Set Slide = Nothing
Set TextBox = Nothing
Set NewPres = Nothing
objPPT.Quit
Set objPPT = Nothing

end if
'
' -- add the customUI colder and two .xml files for the ribbon bar
'
if objFso.FolderExists(oShell.CurrentDirectory+"\"+"customUI") then
    objFso.DeleteFolder oShell.CurrentDirectory+"\"+"customUI"
end if
objFso.CreateFolder oShell.CurrentDirectory+"\"+"customUI"
objFso.CopyFile oShell.CurrentDirectory+"\"+"customUI.xml", oShell.CurrentDirectory+"\"+"customUI\customUI.xml"
objFso.CopyFile oShell.CurrentDirectory+"\"+"customUI.xml", oShell.CurrentDirectory+"\"+"customUI\customUI14.xml"
Dim Shell, objExec, cmd
Set Shell = CreateObject("WScript.Shell")
cmd = "zip.exe -r "+pptmFile+" customUI"
WriteLine cmd
Set objExec = Shell.Exec(cmd)
WriteLine objExec.Status
Do Until objExec.Status = 0
    WScript.Sleep 10
Loop
WriteLine objExec.StdOut.ReadAll()
Set objExec = Nothing
objFso.DeleteFolder oShell.CurrentDirectory+"\customUI"
'
' -- modify the _resl\.rels XML file to include the relationships to the new ribbon bar
'
if objFso.FolderExists(oShell.CurrentDirectory+"\"+"_rels") then
    objFso.DeleteFolder oShell.CurrentDirectory+"\"+"_rels"
end if
cmd = "unzip.exe "+pptmFile+" _rels\.rels"
WriteLine cmd
Set objExec = Shell.Exec(cmd)
WriteLine objExec.Status
Do Until objExec.Status = 0
    WScript.Sleep 10
Loop
WriteLine objExec.StdOut.ReadAll()
Set objExec = Nothing
Dim inFile , outFile
cmd = oShell.CurrentDirectory+"\_rels\.rels"
WriteLine cmd
Set inFile = objFso.OpenTextFile( cmd , 1 , 1 )
Dim text
Do Until inFile.AtEndOfStream
    text = text + infile.ReadLine
Loop
'
Dim xml , xmllasttag
xmllasttag = "</Relationships>"
xml = "<Relationship Target=""customUI/customUI14.xml"" Type=""http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"" Id=""Rcf95296139a74fce""/>"
xml = xml + "<Relationship Target=""customUI/customUI.xml"" Type=""http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"" Id=""Rfc8d3c28a5a34944""/>"
xml = xml + xmllasttag
text = Replace( text, xmllasttag, xml)
Set outFile = objFso.CreateTextFile( oShell.CurrentDirectory+"\_rels\.relsnew" , 1)
outfile.WriteLine text
outFile.Close
inFile.Close
objFso.DeleteFile oShell.CurrentDirectory+"\_rels\.rels"
objFso.MoveFile oShell.CurrentDirectory+"\_rels\.relsnew" , oShell.CurrentDirectory+"\_rels\.rels"
cmd = "zip.exe -r "+pptmFile+" _rels"
WriteLine cmd
Set objExec = Shell.Exec(cmd)
WriteLine objExec.Status
Do Until objExec.Status = 0
    WScript.Sleep 1
Loop
WriteLine objExec.StdOut.ReadAll()
Set objExec = Nothing
objFso.DeleteFolder oShell.CurrentDirectory+"\_rels"
'
' -- ppmt file creation donw - now open it and save it as a ppam file
'
Set objPPT = CreateObject( "PowerPoint.Application" )
objPPT.Visible = True
set NewPres = objPPT.Presentations.Open(pptmFile)
NewPres.SaveAs ppamFile, ppSaveAsOpenXMLAddin
NewPres.Close
Set NewPres = Nothing
objPPT.Quit
Set objPPT = Nothing


