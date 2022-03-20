# PowerGLE

PowerGLE is an MS-Office PowerPoint Add-in for Graphics Layout Engine ([GLE](https://glx.sourceforge.io)), which is a graphics scripting language designed for creating publication quality graphs, plots, diagrams. PowerGLE generates and inserts bitmap images (PNG, JPEG, or TIFF) from GLE source code onto a slide that can be manipulated as a PowerPoint image. It is inspired by [IguanaTeX](https://www.jonathanleroux.org/software/iguanatex)  and works in a similar manner.  The GLE code and data to draw the figures is saved within the PowerPoint presentation, making managing and editing GLE code for numerous figures simpler than externally storing it and copy & pasting from the GLE previewer.

## Download & Installation

PowerGLE is distributed as a .ppam (PowerPoint Add-in File) and is a collection of VBA macros and forms. Download the PowerGLE Add-in file from: 

* GitHub [PowerGLE.ppam](https://github.com/vlabella/PowerGLE/releases/download/1.0.0/PowerGLE-1.0.0.ppam) 
* SourceForge [PowerGLE.ppam](https://sourceforge.net/projects/glx/files/PowerGLE/1.0.0/PowerGLE-1.0.0.ppam/download)

Install it by opening PowerPoint and selecting Options->Add-ins->Manage PowerPoint Add-ins->Go.  Then select "Add New..." and choose the PowerGLE.ppam file.  Macros must also be enabled in Options->Trust Center->Trust Center Settings by selecting "Enable all macros ..."

## Screen Shot

![PowerGLE Screen Shot](https://glx.sourceforge.io/images/PowerGLEScreenShot.PNG "PowerGLE Screen Shot")

## Requirements

* [GLE](https://glx.sourceforge.io) and its supporting applications such as LaTeX and ghostscript.
* MS office 2016 or later (may work on earlier versions but is untested)

## Usage

Select the PowerGLE tab and choose "New GLE Figure", enter the GLE code and either click "Generate" or "Generate & Close" to see the GLE figure on the slide. If a new presentation is created it must be saved at least once prior to creating a new GLE figure. To edit the figure select the figure and then choose "Edit GLE Figure" from the ribbon bar.  Choose "Options" to see the various options.

By default PowerGLE writes all the GLE code to a temporary subfolder within the folder that contains the current presentation.  The code for each GLE figure is contained within its own subfolder named "figure_#" along with any secondary and output files.  By default these folder are

    .\PowerGLE\<presentation file name>\figure_1
    .\PowerGLE\<presentation file name>\figure_2
    .\PowerGLE\<presentation file name>\figure_3 
    ...

The GLE file is named `figure_#.gle` within the folder.  This name can be changed by entering a new name in "Figure Name" field.  In this way all the GLE code can be quickly retrieved outside of the PowerPoint application.  PowerGLE leaves these files on the computer after the application closes.  These files are not needed by PowerGLE and can be manually deleted if desired since all the GLE code is stored within the PowerPoint presentation.

Data files can be imported for each figure under the "Data File(s)" tab.  The name of the file can be used in the GLE code `data` statement.

## Building

The visual basic for applications (VBA) code is contained within this repository along with a vbscript `cppam.vbs` that generates both the `.pptm` and `.ppam` files.  To create both `PowerGLE.pptm` (macro enable presentation) and `PowerGLE.ppam` (PowerPoint add-in) type the following at the command prompt.

    cscript cppam.vbs PowerGLE

Open the PowerGLE.pptm file and hit Alt-F11 to open the VBA editing console to twiddle the code.  The `cppam.vbs` script relies on both the zip and unzip applications, which should be installed on the machine. (see http://infozip.sourceforge.net/)

Any changes that are made to the code in the `PowerGLE.pptm` file must be exported by running the macro `ExportVisualBasicCode` within PowerPoint.  This will place all the VBA code in a subfolder `.\PowerGLE_VBA`.  This will need to be copied over the original code for storing under version control or regenerating the PowerPoint files using `cppam.vbs`.
