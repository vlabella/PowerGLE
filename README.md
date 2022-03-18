# PowerGLE
PowerGLE is a MS-Office PowerPoint Add-in for GLE.  GLE (Graphics Layout Engine) is a graphics scripting language designed for creating publication quality graphs, plots, diagrams, figures and slides (see https://glx.sourceforge.io or https://github.com/vlabella/GLE).

PowerGLE generates and inserts bitmap images (PNG, JPEG, or TIFF) onto a slide that can be manipulated like any other PowerPoint graphic. It is inspired by IguanaTeX and works in a similar manner.  The GLE code and data to draw the figures is saved within the PowerPoint presentation, making managing and editing of the GLE code for numerous figures simpler than externaly storing the code and copy and pasting from the GLE previewer.

PowerGLE is written in Visual Basic For Applications (VBA) and is a collection of macros that is "compiled" into a .ppam file that is added into PowerPoint.

## Requirements

* GLE and its supporting applications such as LaTeX and ghostscript.
* MS office 2016 or later (may work on earlier versions but is untested)

## Installation

Download the PowerGLE.ppam file.  This is a PowerPoint Add-in file that must be manually added into the PowerPoint application.  Open PowerPoint and select Options->Add-ins->Manage PowerPoint Add-ins->Go.  Then select "Add New..." and choose the PowerGLE.ppam file.  In addition, macros must be enabled by going to Options->Trust Center->Trust Center Settings and selecting "Enable all macros ...".

## Usage

Select the PowerGLE tab and choose "New GLE Figure", enter the GLE code and either click "Generate" or "Generate & Close" to see the GLE figure on the slide.  To edit the figure select the figure and then choose "Edit GLE Figure" from the ribbon bar.  Choos "Options" to see the variaous options.

By default PowerGLE writes all the GLE code to a temporary folder where the application file resides.  Each figure gets its own folder named "figure_#" to contian the output and secondary files in.  By default these folder are
 
* .\PowerGLE\\<presentation file name\>\figure_1, 
* .\PowerGLE\\<presentation file name\>\figure_2, 
* .\PowerGLE\\<presentation file name\>\figure_3 , 
* etc.  
 
 The gle file will be named figure_#.gle within the folder.  This name can be changed by entering a new name in "Figure Name" field.  In this way all the GLE code can be quickly retrieved outside of the PowerPoint application.  PowerGLE leaves these files on the computer after the application closes.
 
 Data files can be imported for each figure under the "Data File(s)" tab.  The name of the file can be used in the GLE code data statement.

## Building

The VBA code is contained in this repo along with a vbscript that generates both the .ppmt and .ppam files.  Open the .ppmt file and hit Alt-F11 to open the VBA editing console to twiddle the code.
