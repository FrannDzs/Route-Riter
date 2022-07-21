SHAPE FILE MANAGER Version 2.4a
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

6th July 2005
################

Some small fixes.

1st December 2004
################

Some small fixes.

13th August 2003
################

This version has several fixes for previous versions, and supercedes 2.3x and 2.3xx

2.4 now detects windows NT/2000/XP systems and runs the compress/uncompress
as an invisible process, should solve the problem with sorting of files in
European locales and stop the lock-up when examining route shape folders
by only returning the first 600 files.

Installation
############

Copy the file SFM.HTA into any folder you like, and double-click
on it to run the program.


Problems?
#########

This text file contains possible solutions to problems with running this program:

Firstly, this program requires IE6 or higher to work.

*	Win 98 and ME users especially - Version 5.6 of Windows Scripting is required.
	This can be downloaded from MS and is only 650Kb
	http://msdn.microsoft.com/library/default.asp?url=/downloads/list/webdev.asp

*	If SFM shows no menus when the mouse moves over the Option button next to the S file name
	 - one solution has been to install all the current MS patches for IE

*	If an object fails to compress/uncompress with an [Object Error] message.
	1) Check the security level in your internet options under the
	 heading of "Intranet" - This should be Medium or Medium-Low for SFM to work.
	2) Put the SFM.HTA file into the same folder as the FFEDIT
	 program and create a shortcut to it on the desktop or menu.
	 Sometimes if SFM is copied directly onto the desktop this error can occur.

Hope this helps....

Paul

====================================================================================
For those of you for whom the Instructions button does not work, here is a
duplicate of the text:

MSTS Shape File Manager Version 2.4 - Help
** This utility is designed to help MSTS model builders manage shape files.
Use of this tool by someone unfamiliar with the file management requirements of
MSTS may result in routes being unable to load. **

Version 2.4, Jun 2003 - Attempt to fix Win 98/ME problems

Version 2.3x, Nov 2002 - Correction of object reverse and scaling with
negative scale factors by Okrasa Ghia

Version 2.3a, Oct 2002 - Sorted file and directory lists, object reverse
(rotate 180 degrees), a few display bug fixes and tried to detect ffeditc_unicode.exe

Version 2.2, May 2002 - Added Distance Level adjuster for altering the
distance at which shapes change to lower levels or the maximum visible distance.

2.2a/b - fixed translucency priority bug, 2.2c fixed missing bracket in
wag/eng scale update, 2.2d fix problem if read only file is uncompressed.

Version 2.1, May 2002 - Shift option added for adjusting objects positions
relative to their pivot point.

Version 2, March 2002 - corrected scaling for animations, scaling applied to
WAG/ENG files (some parameters) and Texture mode for specular highlights.

The main form behaves much like the Windows Explorer - the left side shows
the current folder with a button to navigate up one level to the parent. Below
the current folders is a list of sub-folders. Selecting one of these makes it
the current folder.
The Buttons across the top allow swapping to another drive.

On the right side is a list of shape files contained in the folder. Size and
compression information is shown.  Moving the mouse over the Options area 
shows a menu that is available for the current file.

Compressed files may be Uncompressed or their associated SD files can 
be edited, in this version.

Uncompressed files have the following options:

Compress - this runs FFEDITC_UNICODE.EXE to compress the file - may not 
work on locomotives with animations unless you have the patched newshape.bnf 
file on your system (included in this ZIP archive)

Scale - resizes an object by altering the points, matrices and vol_sphere 
sections of the file - will also update .sd bounding box information if it 
can. In both cases, a backup of the previous state of the file is kept with 
a ".PreScale" file extension.

Reverse - reverses (rotate 180 degrees) an object by altering the points, 
matrices and vol_sphere sections of the file - will also update .sd bounding 
box information if it can. In both cases, a backup of the previous state of 
the file is kept with a ".PreScale" file extension

Shift - adjusts an objects position relative to its origin (pivot point). 
The 3 prompts are for the distance moved in metres i.e. 0.05 = 5cm - positive Y 
values are up. Useful for adjusting models that sink into the rails slightly.

Distance Levels - allows changes to the shapes distance levels of detail. Reducing 
values here help to improve frame rates by not loading the shape at distances 
over the values entered (where there is only one level). Basically the maximum 
viewable distance should be proportional to the object size

Texture Mode - Allows the user to change the texture mode of the groups of 
objects in a shape. This option also applies the specular highlight fix for 
shiny textures. Unless the groups have been well named, this process can be 
a bit hit and miss. A backup of the previous state of the file is kept with 
a ".PreTexture" file extension

Wordpad Edit - runs Wordpad.exe with the .S file
Wordpad Edit .SD - if the .sd file does not exist it is created then opened in Wordpad. 

The WAG and ENG file parameters scaled are :
POSITION (lights), 
RADIUS (lights), 
HEADOUT, 
PASSENGERCABINHEADPOS, 
INTAKEPOINT, 
SIZE, 
CENTREOFGRAVITY, 
INERTIATENSOR, 
WHEELRADIUS and all parameters ending with FX (Steam and Diesel)

A certain amount of manual editing of these files will also be needed to 
scale the power and weight correctly. 

A small word of warning on WAG/ENG files - These are manually changed by authors 
and sometimes the scaling does not work properly e.g. on entries such as "24in/2" 
and can cause effects like wheels disappearing. Always a good idea to manually 
check the WAG/ENG file afterwards. 
