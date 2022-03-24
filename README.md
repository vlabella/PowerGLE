# PowerGLE

PowerGLE is an MS-Office PowerPoint Add-in for Graphics Layout Engine ([GLE](https://glx.sourceforge.io)), which is a graphics scripting language designed for creating publication quality graphs, plots, and diagrams. PowerGLE generates and inserts figures created by GLE onto a slide that can be manipulated as a PowerPoint image. The GLE code and data utilized to draw each figure is stored within the PowerPoint presentation, simplifying the management and editing of large presentations with multiple GLE figures. PowerGLE is inspired by [IguanaTeX](https://www.jonathanleroux.org/software/iguanatex).

## Download & Installation

PowerGLE is distributed as a .ppam (PowerPoint Add-in) and is a collection of VBA macros and forms. Download the PowerGLE Add-in file from: 

* GitHub [PowerGLE.ppam](https://github.com/vlabella/PowerGLE/releases/download/1.0.1/PowerGLE_v1_0_1.ppam) 
* SourceForge [PowerGLE.ppam](https://sourceforge.net/projects/glx/files/PowerGLE/1.0.1/PowerGLE_v1_0_1.ppam/download)

Install it by opening PowerPoint and selecting Options->Add-ins->Manage PowerPoint Add-ins->Go.  Then select "Add New..." and choose the PowerGLE.ppam file.  Macros must also be enabled in Options->Trust Center->Trust Center Settings by selecting "Enable all macros ..."

## Screen Shot

![PowerGLE Screen Shot](https://glx.sourceforge.io/images/PowerGLEScreenShot.PNG "PowerGLE Screen Shot")

## Requirements

* [GLE](https://glx.sourceforge.io) and its supporting applications such as [LaTeX](https://www.latex-project.org/) and [Ghostscript](https://www.ghostscript.com/).
* MS office 2016 or later (may work on earlier versions but is untested)

## Usage & Features

### Figure creation & editing

Select the PowerGLE tab and choose "New GLE Figure", enter the GLE code and either click "Generate" or "Generate & Close" to see the GLE figure on the slide. If a new presentation is created it must be saved at least once prior to creating a new GLE figure. To edit the figure select the figure and then choose "Edit GLE Figure" from the ribbon bar.  Choose "Options" to see the various options.

### Importing data files

Data files can be imported for each figure under the "Data File(s)" tab.  The name of the file can be used in the GLE code `data` statement.

### Options

The options that control the creation of the GLE figure are:

* DPI: controls the resolution or dots per inch of the bitmap image.
* Cairo: GLE will use the [cairo](https://www.cairographics.org/) engine when rendering the figure.
* Output format: PNG, JPEG, and TIFF are supported (bitmap types only).
* Transparent: make the background transparent, only for PNG format.
* Figure name: controls the filename and temporary folder name for this image (see below).
* Scaling gain: controls the initial size of the figure on the slide.  The scale of a newly created PowerPoint image is calculated as \(screen_dpi/output_dpi x scaling_gain\).  Increasing or decreasing the gain will make the initial size large or smaller, respectively.


### Temporary Files & Figure Name

PowerGLE creates many temporary files while creating each figure that are left on the computer after the application closes.   This allows for quick retrieval of the GLE code outside of the PowerPoint application.  *These temporary files are not needed by PowerGLE since all the GLE code is stored within the PowerPoint presentation.*

A relative temporary root folder is utilized by default as a sub-folder within the folder that contains the current presentation.  The default name for this sub-folder is `PowerGLE\<presentation_name>`, where `<presentation_name>` is the filename (with extension) of the active presentation, where the `.` is replaced by `_`.  Within this sub-folder, each GLE figure's code and output files are contained within its own sub-folder named the figure name.  The default name for each figure is `figure_#`, where `#` is a number.  For example, for a PowerPoint presentation named `My Presentation.pptm` these folders would look like

    .\PowerGLE\My Presentation_pptm\figure_1
    .\PowerGLE\My Presentation_pptm\figure_2
    .\PowerGLE\My Presentation_pptm\figure_3 
    ...

The GLE file is named `figure_#.gle` within the folder.  This name can be changed for any figure by entering a new name in "Figure Name" field, which will also change its folder name as well.   

There is an option to use an absolute temporary root folder as well, but use this with caution as two presentations with the same name from different folders will overwrite each others files.  For example if `c:\temp` is chosen for a presentation named `My Presentation.pptm` these folders would look like

    c:\temp\PowerGLE\My Presentation_pptm\figure_1
    c:\temp\PowerGLE\My Presentation_pptm\figure_2
    c:\temp\PowerGLE\My Presentation_pptm\figure_3 
    ...

When a GLE figure is copied within PowerPoint the new figure is given the next available default figure name `figure_#`, even if the source figure has a non-default figure name.


## Building

The visual basic for applications (VBA) code is contained within this repository along with a vbscript `cppam.vbs` that generates both the `.pptm` and `.ppam` files.  To create both `PowerGLE.pptm` (macro enable presentation) and `PowerGLE.ppam` (PowerPoint add-in) type the following at the command prompt.

    cscript cppam.vbs PowerGLE [version]

Open the PowerGLE.pptm file and hit Alt-F11 to open the VBA editing console to twiddle the code.  The `cppam.vbs` script relies on both the zip and unzip applications, which should be installed on the machine. (see http://infozip.sourceforge.net/).  The version number is optional and will be added to the filename. '.' in the version number will be changed to '\_' since PowerPoint cannot load .ppam files with more than one '.' in the filename.

Any changes that are made to the code in the `PowerGLE.pptm` file must be exported by running the macro `ExportVisualBasicCode` within PowerPoint.  This will place all the VBA code in a sub-folder `.\PowerGLE_VBA`.  This will need to be copied over the original code for storing under version control or regenerating the PowerPoint files using `cppam.vbs`.
