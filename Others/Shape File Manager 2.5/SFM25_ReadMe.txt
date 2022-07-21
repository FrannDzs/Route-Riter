Shape File Manager v2.5 (SFM25) is a revised version of Paul Gaudsen's SFM v2.4a.

CHANGES:
=======

     - Rotate an object 90 degrees clockwise (CW) or counterclockwise (CCW).
     - Adjust MIP map levels for textures.
     - Configure and use an alternate Unicode text editor.
     - Separate dialog box for settings.
     - Algorithm to recalculate normals when scaling shape files.
     - Capability to adjust complex bounding boxes.
     - Confirmation dialog boxes for compress, uncompress, reverse and rotate.
     - Numerous tweaks and fixes to the program and user interface.

     - Removed option to adjust WAG and ENG files

       Due to the non uniform formatting used in WAG and ENG files, successfully adjusting
       all necessary parameters is problematical at best. Further, it is not a trivial
       operation to Reverse, Rotate, Scale or Shift the shape files for wagons or engines as
       this will usually affect the animation. If the user understands all of the additional
       editing that must be done to the shape file to accomplish this, he or she is better
       advised to manually edit the WAG or ENG file also.

     - Dropped compatibility with Windows 9x and ME
     
       Windows 9x and ME require additional downloads to work correctly which are no longer
       available or supported by Microsoft. 


INSTALLATION:
============

     Copy the file SFM25.HTA into any folder you like, and double-click on it to run
     the program.  If you have a working copy of SFM2.4a (or earlier), just copy
     SFM25.HTA into the same folder.

     Super Simple Installation:  SFM25 will look for FFEDITC_UNICODE.EXE in the same
     folder that contains SMF25.
     
     1) Find FFEDITC_UNICODE.EXE (see below for default location).
     2) Copy SFM25.HTA and NEWSHAPE.BNF into that folder
     3) Double-click on SFM25.HTA to run the program.
     4) (Optional) Create a shortcut to SFM25.HTA in your start menu or on the desktop. 

     FFEDITC_UNICODE.EXE: SFM25 calls FFEDITC_UNICODE.EXE to compress and uncompress
     shape files.  The full path to FFEDITC_UNICODE.EXE must be manually entered in the
     settings dialog if SFM25 is unable to locate it. The default search order is:
     
     1) The saved SFM25 setting (if it exist).
     2) The SFM25 folder.
     3) The original installation folder of MSTS (read from the registry).
     4) C:\Program Files\Microsoft Games\Train Simulator\UTILS\FFEDIT\  

     NEWSHAPE.BNF: Some shape files, particularly locomotive models, that utilize
     "slerp_rot" in the ANIMATIONS section require a patched version of the NEWSHAPE.BNF
     file to compress/uncompress correctly.  Copy the enclosed patched NEWSHAPE.BNF to the
     folder that contains FFEDITC_UNICODE.EXE after making a backup of the original.
     (A copy of the original NEWSHAPE.BNF is also included in this distribution.)

     Default installation folder for FFEDITC_UNICODE.EXE and NEWSHAPE.BNF:
     
          C:\Program Files\Microsoft Games\Train Simulator\UTILS\FFEDIT\
     

REQUIREMENTS:
============

SFM25 should run on Windows XP (or higher).  SFM25 will run on Windows 2000 but may require
the installation of Internet Explorer 6 and Windows Script v5.7 (currently still available
from Microsoft). As of version 2.5 all support from Windows 98 and ME has been removed
from SFM.

If an object fails to compress or uncompress with an [Object Error] message.

     1) Check the security level in your internet options under the heading of "Intranet".
        This should be Medium or Medium-Low for SFM25 to work.

     2) Put the SFM25.HTA file into the same folder as the FFEDIT program and create a
        shortcut to it on the desktop or menu. This may problem may occur if SFM25 is copied
        directly onto the desktop.

Some users have reported that FFEDITC_UNICODE.EXE must be configured to run in compatibility
mode under Vista and Windows 7. 


CREDITS:
=======

     Paul Gausden (Decapod) for the original Shape File Manager and documentation.


DISCLAIMER, LICENSE and COPYRIGHT:
=================================

     This software is provided free of charge for any non-commercial purpose.
     
     No warranty is offered with this software.
     
     The authors take no responsibility for any problems arising directly or
     indirectly from its use.

     I have not knowingly misused or misappropriated any protected works in the
     creation of this software. 

     Huecuvoe
     Huecuvoe@AOL.com


INSTRUCTIONS:
============

MSTS Shape File Manager Version 2.5 - Help

** This utility is intended to help MSTS model builders manage shape files. Use
of this tool by someone unfamiliar with the file management requirements of MSTS may
result in routes being unable to load. **

The main form behaves much like the Windows Explorer - the left side shows the
current folder with a button to navigate up one level to the parent.

Below the current folders is a list of sub-folders. Selecting one of these makes
it the current folder.  The Buttons across the top allow navigating to another drive.

On the right side is a list of shape files contained in the folder including size and
compression information.  The default is to display a maximum of 600 shape file names.
The file list is limited because on some systems and under some conditions SFM25 will
display an error message if there are too many shape files in a folder. This limit can
be disabled in the SFM25 "Settings" dialog. In any case, it is better to run SFM25 on
shape files in a working folder and not directly on a route installation. 

Click on the "Options" button to the right of the shape file name to display the menu of
options (actions) that are available for the highlighted target file.


Options for COMPRESSED files:

Uncompress - Call FFEDITC_UNICODE.EXE to uncompress the shape file - may not work on
locomotives with animations unless you have the patched NEWSHAPE.BNF file on your system.

Edit .SD File - Edit the .SD file with the configured Unicode editor. If the .SD file does
not exist it will be created.


Options for UNCOMPRESSED files:

Compress - Call FFEDITC_UNICODE.EXE to compress the shape file - may not work on
locomotives with animations unless you have the patched NEWSHAPE.BSF file on your system.

Distance Levels - Allows changes to the shapes distance levels of detail. Reducing values
here help to improve frame rates by not loading the shape at distances over the values
entered (where there is only one level). Basically the maximum viewable distance should be
proportional to the object size. A backup of the shape file is made with a ".PreDistance"
file extension.

MIP Map Levels - Allows changes to the shapes MIP Map levels for textures. Reducing values
here may help to improve the appearance of textures and decrease "blurriness" at the
expense of increased aliasing and moire. A backup of the shape file is made with a
".PreTexture" file extension.

Reverse - Reverse an object (rotate 180 degrees about the Y axis) by altering the
vol_sphere, points, vectors, sort_vectors, matrices and animations sections of the file.
Backups of the .S and .SD files are made with a ".PreReverse" file extension.

Rotate CCW - Rotate an object 90 degrees counterclockwise about the Y axis (looking down)
by altering the vol_sphere, points, vectors, sort_vectors, matrices and animations
sections of the file. Backups of the .S and .SD files are made with a ".PreRotate" file
extension.

Rotate CW - Rotate an object 90 degrees clockwise about the Y axis(looking down). Backups
of the .S and .SD files are made with a ".PreRotate" file extension.

Scale - Resize an object by altering the vol_sphere, points, vectors, sort_vectors,
matrices and animations sections of the file. Backups of the .S and .SD files are made with
a ".PreScale" file extension.

Shift - Adjust an objects position relative to its origin (pivot point). The 3 prompts
are for the distance moved in metres i.e. 0.05 = 5cm  - positive Y values are up. Backups
of the .S and .SD files are made with a  ".PreShift" file extension.

Texture Mode - Allows the user to change the texture mode of the matrices (groups) of
objects in a shape file. This option also applies the specular highlight fix for shiny
textures. Unless the matrices have been well named, this process can be a bit hit and
miss. Backups of the .S and .SD files are made with a ".PreTexture" file extension.

Edit .S File - Edit the .S file with the configured Unicode editor.

Edit .SD File - Edit the .SD file with the configured Unicode editor. If the .SD file does
not exist it will be created.

Note: If a shape data (.SD) file exists, it will be automatically adjusted as part of the
Reverse, Rotate, Scale and Shift options.


SETTINGS:

The following SFM25 options can be configured in the "Settings" dialog:

     FFEDITC_UNICODE.EXE: Enter the full path to FFEDITC_UNICODE.EXE (with or without
                          the trailing backslash). If no path for FFEDITC_UNICODE.EXE
                          has been configured, SFM25 will try and locate it in the
                          SFM25 folder or the installation folder for MSTS.  

          Unicode Editor: By default, SFM25 will use WORDPAD.EXE to edit Unicode files.
                          The user can configure an alternate Unicode editor by entering
                          the fully qualified pathname. The path name is not required if
                          the alternate Unicode editor is on the user's PATH.
                          
  Confirm ALL Operations: By default, SFM25 will ask for confirmation before performing
                          any operation. If this option is "unchecked", SFM25 will
                          immediately COMPRESS, UNCOMPRESS, REVERSE or ROTATE a shape
                          file without confirmation. Disabling confirmation may speed up
                          multiple operations but increases the risk of a mistake. 

         Limit File List: By default, SFM25 will limit the file list to a maximum of 600
                          names.  If this option is "unchecked", SFM25 will not limit the
                          file list. Disabling "Limit File List" may result in very slow
                          execution and warning or error messages. Your computer may
                          become unresponsive or crash completely. It is better to run
                          SFM25 on shape files in a working folder and not directly on a
                          route installation.


CAUTIONS:
========

Shape File Manager is a simple program designed to make relatively simple changes to MSTS
shape files.  It is NOT a substitute for dedicated 3D modeling software.

Shape files are very complicated entities and may be corrupted and rendered unusable by
SFM25. Although SFM25 will normally function properly on "simple" shape files; complicated
shape files, especially those involving animation, may cause it to fail.  Some shape files
have defective or incomplete animation specifications and will not ROTATE or REVERSE
correctly.

SFM25 must recalculate the surface normals when scaling a shape file using different scale
factors for X, Y and Z.  This may introduce errors into the shape file that will cause it
to display incorrectly.  Shape files with animations, particularly rolling stock, are
especially susceptible to this problem.

The user is cautioned to ALWAYS make secure backups. 


HISTORY:
=======

Version 2.5, August 2012 - Rotate shape 90 degrees CCW and CW.
                           Adjust MIP map levels for textures.
                           Configure alternate Unicode text editor.
                           Separate dialog box for settings.
                           Algorithm to recalculate normals when scaling shape files.
                           Capability to adjust complex bounding boxes.
                           Numerous tweaks and fixes to the program and user interface.
                           Removed option to adjust WAG and ENG files.
                           Dropped compatibility with Windows 9x and ME.

Version 2.4a, July 2005 - Some compressed files appeared as unknown.

Version 2.4, Jun 2003 - Attempt to fix Win 98/ME problems and shifting matrix bug

Version 2.3x, Nov 2002 - Correction of object reverse and scaling with negative scale
                         factors by Okrasa Ghia

Version 2.3a, Oct 2002 - Sorted file and directory lists, object reverse (rotate 180 degrees).
                         A few display bug fixes and tried to detect ffeditc_unicode.exe

Version 2.2, May 2002 - Added Distance Level adjuster for altering the distance at which
                        shapes change to lower levels or the maximum visible distance.
        2.2a/b        - Fixed translucency priority bug
        2.2c          - Fixed missing bracket in wag/eng scale update
        2.2d          - Fix problem if read only file is uncompressed.

Version 2.1, May 2002 - Shift option added for adjusting objects positions relative to
                        their pivot point.

Version 2, March 2002 - Corrected scaling for animations, scaling applied to WAG/ENG files
                        (some parameters) and Texture mode for specular highlights.
