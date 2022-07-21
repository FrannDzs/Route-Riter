 Archibald, Horace & Zipper
================================
By Okrasa Ghia (okrasa@xtracks.tk, okrasaghia@yahoo.com), July 2005.
All the utilities have been extensively tested but as with any pice of
software there are no guarantees, use at your own risk.
These utilities are free to use and share, you must not sell them.
Please inquire about uploading at other sites than www.train-sim.com.
Also inquire if you want to use the TK.MSTS.Tokens.dll for reading and
writing MSTS files, I'm open to licensing it for use in other utilities.

 Installation
---------------
Installation of these utilities is simple and does not even require 
you have MSTS installed on the computer.
Only requirement is that you have 'Microsoft .Net Framework v1.1' or
later, the name of the installer is 'dotnetfx.exe'.
If you don't have it already it can be freely downloaded here:
http://
or you can get it from 'Windows Update'.
If you have a fairly new computer with Windows regularly updated chanse
is you already have the framework installed.

The files included with this package can be placed anywhere you like.
Only requirement is that the exe files and dll are located in the same 
folder or the applications will fail to locate the dll.

 Archibald
------------
This is a tool aimed at content makers. 
Archibald can open and display the content of different MSTS files in 
a tree view display. Does syntax checking of file content and can both 
read and write compressed MSTS files. One example of use for Archibald
is to change the names of textures in shape files without the need to
go through the process of decompressing, open in Wordpad and the again
compressing the file. Another use is to take a peek into terrain files.
Supported files are: tsection.dat, s, sd, t, tdb, trk, w, eng & wag 
(the files Horace needs with a few additions). 
I will add other files as needed.

Archibald started out as a testing platform for the reading and writing
of files needed by Horace but is usefull enough to be released.
Please note that you can drop files om Archibald's icon/shortcut or window
to open the file thus no need to associate any files with Archibald.

WARNING - while using Archibald makes editing MSTS files easier you can
make much damage to the files if you don't know what you are doing.

 Horace 2.0
-------------
This tool is only aimed at route builders.
A much needed update of Horace that can handle the now quite large 
'standardized tsection.dat'. This is actually a complete rewrite of the
program giving much improved speed and handling of compressed w-files.
Now can also update the tdb-file if you only update the dynamic track
sections (tested on Marias Pass). Alterations to static track sections
will still force you to make a regular tdb rebuild in the Route Editor.
This version of Horace does not read the registry but lets you freely
select tsection.dat and route, you don't even need MSTS installed.

The purpouse of Horace is to adjust ids of dynamic and static track
sections when you do changes to the global tsection.dat file. One such 
situation is when you install the 'standardized tsection.dat' for using
XTracks, New Roads, Scale Rail or the like. If you don't adjust the ids
your route will loose all dynamic track sections as soon as you add a 
another dynamic track. Updating from one build of the 'standardized 
tsection.dat' to a later build you do not need to perform this update.

WARNING - using Horace without understanding it's purpouse will likely
cause harm to your route and could DESTROY it beyond repair.
Always backup your route before using Horace.

 Mapper
--------
Mapper is a tool for everyone who's been missing a map of that newly 
downloaded route. This tool can generate a map or profile automatically 
from the routes track database. The finished map/profile will be showed 
in Mappers window and can then be saved as a gif file.

Options in Mappers window lets you include station names, mileposts 
and/or fuelpoints. There is also the option to add a grid in km or 
miles. The profile shows altitude in meters or feet when grid is 
selected. Input boxes for width and height lets you set the size of 
the generated image. When doing a map only the width value is used, 
height is then calculated not to distort the map.

The profiling part is a prototypical work and only realy usefull with 
routes that have a long main line. Routes with many branches and 
alternative paths like the PO&N will not get a meaningfull profile.

 Zipper
---------
This is probably the utility that will get the most users.
TrainSim Zipper is for anyone that wants to (de)compress shape and world
files. It's an user friendly replacement for the MSTS ffeditc utility.
Can process individual files and/or recurse through folders processing 
all files it can find. Starting from your main Train Simulator folder it
could process all s & w files in your entire installation.

A number of options lets you configure how Zipper should operate.
These options can either be changed in the Zipper window or be added as
switches to shortcuts or bat files. Supports drag-n-drop in both window 
and shortcuts for easy use. Without any switches dropping a file onto 
Zipper will toggle it between compressed and Unicode text-file format.
If compressed before it will become readable text and vice versa.
Using switches you can have icons to compress or decompress files without
worrying about the files state before dropped on Zipper.

Switches:
 /b - change files to uncompressed binary format
 /c - change files to compressed binary format
 /u - change files to Unicode text format
 /r - process folders recursively searching through all subfolders
 /f=<pattern> - only process files matching pattern (*=wildcard),
      pattern is only used whe searching folders
 /s - enforces stricter check of files before processing
      (detects duplicate UiDs in w-files)
Switches /b, /c & /u are mutualy exclusive but can be freely combined
with the other switches.

Examples:
'Zipper.exe /c /r' makes for a shortcut compressing all files dropped 
on it, folders dropped on it will be processed recursively looking for
all s- & w-files. Files already compressed will be left alone.
'Zipper.exe /u /f=*.w' makes for a shortcut decompressing all w-files
in folders dropped on it. Already decompressed w-files will stay text.

 Development support
---------------------
For the benefit of other developers I have included documentation of 
the TK.MSTS.Tokens.dll and a type library (tlb) for the use from VB6.
Developers of MSTS utilities may use the TK.MSTS.Tokens.dll in their 
utilities if the resulting utility is free and full credit is given for 
the use of the TK.MSTS.Tokens.dll. I place no restriction of the form 
of distribution or where these utilities are uploaded as long as there 
is no other charges than for handling and possibly the cost of media.

 Credits
----------
These utilities were made in C# using #develop and #ZipLib from
http://www.icsharpcode.net.
I have to thank Martyn T Griffin for sharing his source code for
compressing and decompressing s- & w-files. Without it I would not
have been able to make these utilities even though my code is much 
different than Martyn's original Visual Basic code.
Thanks also to Mike Simpson (Route Riter) for valuable comments 
trying out the beta releases.

