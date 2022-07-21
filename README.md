# Route-Riter
 Route Riter is a set of tools for Microsoft Train Simulator created by Mike Simpson.

> *If you find a bug or find an improvement for the repository/project, please feel free to inquire/request a change. Your collaboration is greatly appreciated.

> *There are Forms and Modules that are unused if you check the folder containing the source, but due to my limited knowledge in MSTS I could not decipher if these are needed or obsolete. Please, if you know something about it I hope for your collaboration.

## *This repository is under reconstruction.*

# 1. Pre-requisites:
- Download [Visual Basic 6.0](https://winworldpc.com/product/microsoft-visual-bas/60) or [Visual Studio 6.0](https://winworldpc.com/product/microsoft-visual-stu/60)

 - Download [MSDN Library](https://winworldpc.com/product/msdn/vs-60) (Optional)
   
    ***I assume you are running Windows 10 64 bit***
 
- Install Visual Studio 6 and MSDN Library.

*There is a special procedure to install Visual Studio 6 in Windows 10. Use the* [Visual Studio 6 Installer Wizard v5.0](https://github.com/FrannDzs/Route-Riter/blob/main-(7.6.26)/VS6InstallerSetup.exe)

**How to proceed**
- [(Video-tutorial)](https://www.youtube.com/watch?v=1tkTb6AYlAg)
- [Manual installation tutorial](https://www.codeproject.com/Articles/1191047/Install-Visual-Studio-on-Windows)

**After downloading and install the MSDN library you can install the following update for it:
- **MSDN Latest Version Update**: 
   - **Disc1**: https://archive.org/details/MSDN_Library_October_2001_Disc_1 
   - **Disc2**: https://archive.org/details/MSDN_Library_October_2001_Disc_2 
   - **Disc3**: https://archive.org/details/MSDN_Library_October_2001_Disc_3
 - Install [Service Pack 6](https://web.archive.org/web/20120707200906/http://download.microsoft.com/download/1/9/f/19fe4660-5792-4683-99e0-8d48c22eed74/Vs6sp6.exe)
 
 - Install [Microsoft Visual Basic 6.0 Service Pack 6 Cumulative Update](https://www.microsoft.com/en-us/download/details.aspx?id=7030)
 
 - Install [Microsoft Visual Basic 6.0 Service Pack 6 Security Rollup Update](https://www.microsoft.com/en-us/download/details.aspx?id=50722)
   
 - Install C1 Controls manually with [cabs files](https://github.com/FrannDzs/Route-Riter/tree/main-(7.6.26)/Source/Dependancies/ComponentOne%20Installers) or with [.ocx files](https://github.com/FrannDzs/Route-Riter/tree/main-(7.6.26)/Source/Dependancies/ComponentOne%20.ocx), or better make an installation from the [installers.](https://github.com/FrannDzs/Route-Riter/tree/main-(7.6.26)/Source/Dependancies/ComponentOne%20Installers)
 
If you have problems loading the vsflex8l.ocx vsprint.ocx and c1sizer.ocx controls in visual basic try copying and registering them in both system directories: /system32 and /SysWOW64 

# 2. Recommended/Interest addins/tools for VB6 IDE:

 - [Code Advisor](https://www.microsoft.com/en-US/download/details.aspx?id=1222) (Free) (Optional)

 - [Code Help 3.0](https://github.com/clayreimann/CodeHelp) (Free) (Tabs in VB IDE) (Recommended)

 - [Visual Basic 6 Mouse Wheel Fix](https://github.com/FrannDzs/Route-Riter/blob/main-(7.6.27)/Others/vb6mousewheelfix.exe) (Free) (Fix Scroll Whel in IDE)

 - [MZ-Tools 8.0](https://www.mztools.com/v8/mztools8.aspx) (Paid) (Tools Suite) (Very Recommended)

 - [CodeSMART](https://www.axtools.com/products-codesmart-vb6.php) (Paid) (Tools Suite) (Very Recommended)

 - [Codejock Suite Pro for ActiveX](https://codejock.com/products/suitepro/?2yn6s14z=p1z) (Paid) (Controls Suite) (Recommended)

 - [ModernVB](https://github.com/VykosX/ModernVB) (Free) (Modernize your VB6 IDE) (Very Recommended)

 - [OLEEXP : Modern Shell Interfaces](https://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb) (Free)

 - [VB Common Controls Replacement Library](https://github.com/Kr00l/VBCCR) (Free) (Common Controls Remake) (Recommended)

 - [DX9VB](https://github.com/thetrik/DX9VB) (Free) (Direct3D9 for Visual Basic 6)

 - [VbTrickMultiThreading](https://github.com/thetrik/VbTrickThreading) (Free) (Module for working with multithreading in VB6)

 - [VbTrickTimer](https://github.com/thetrik/VbTrickTimer) (Free) (Timer class for VB6/VBA compatible with 64 bit office)
 
 - [VB64bitDLLusage](https://github.com/thetrik/Vb64BitDllUsage) (Free) (Using 64-bit dll in VB6 (in WOW64)

 - [VBPNG](https://github.com/thetrik/VbPng) (Free) (provide the ability to work with PNG/ICO/CUR/ANI images using the standard controls)

 - [DeleteVbwFiles](https://github.com/EduardoVB/VB6-AddIn-Delete-vbw-Files) (Free) (VB6 Add-In to delete vbw files automatically (when opening or closing project)

 - [VB6NAMESPACES](https://github.com/WindowStations/VB6NameSpaces) (Free) (A single VB.NET assembly makes it possible to interoperate with VBA/VB6 including VB.NET Forms, Controls, Properties, Events, and NameSpaces, instanced as nested class buckets)

 - [VB6 PORTER](https://github.com/VBForumsCommunity/VB6Porter) (Free) (Supports a use of the language that accommodates porting VB code, both forwards and backwards)

 - [VBSQlite](https://github.com/Kr00l/VBSQLite) (Free) (VB SQLite Library)

# 3. Run compiling
 - Run Visual Basic 6.0, choose the option to load existing project and choose Route_Riter7.vbp to load the project.
 
 - Make the necessary changes and run with F5 or ctrl + F5 to run with a complete compilation.

 - From Project>Properties>Compile you can choose the compilation mode.

 - From the file menu you can generate an .exe

 - Drop the compiled executable to the Release folder

# 4. After compiling the project 
 - Extract all compressed files from [./Release/Dependancies](https://github.com/FrannDzs/Route-Riter/tree/main-(7.6.26)/Release/Dependancies) directory to the root of the Release folder. 
 Install [mwgfxdll.exe.](https://github.com/FrannDzs/Route-Riter/blob/main-(7.6.26)/Release/mwgfxdll.exe)

# Development plans
 - Improve the current code following the original direction.

# Future development plans
- Prepare the Visual Basic 6.0 code for migration to Visual Basic .NET.

- Separate application and data tiers into a DLL separate from presentation.

- Change the user interface to an inductive user interface, https://msdn.microsoft.com/en-us/library/ms997506.aspx

- Fully parse the Microsoft Train Simulator files by adapting an XML parser, http://pugixml.org/

# Credits/Acknowledgement:
© Mike Simpson for writing this initially, the kindness and giving the necessary permissions.
virtualtrains@tpg.com

© Jeffrey Kraus for donating his code and allowing this to be revived.
http://www.digital-rails.com
support@digital-rails.com

© Carl-Heinz Rave
TsUtils
http://www.carloshr.de
mail@carloshr.de

© Scott Miller
AceIt
aceit@ameritech.net

© Martin Wright
TGATools2A
http://www.mwgfx.co.uk/index.htm
martin@mwgfx.co.uk

© Paul Gausden
Shape Viewer
https://decapod99.wordpress.com/

© Edward Grubb
PicFormat32
https://github.com/Planet-Source-Code/edward-grubb-ed0-picformat32__1-13267

© Franky Braem
SAWZipNG
http://www.codeproject.com/Articles/875/SAWZip-zip-file-manipulation-control

© Jean-loup Gailly
zlib.dll
http://www.zlib.net/

© ComponentOne
c1sizer.ocx, vsflex8l.ocx and vsprint8.ocx
http://www.componentone.com/

© Jordan Rusell Software
Inno Setup Installer
https://jrsoftware.org/

© Uwe Herklotz
UHARC.EXE
Uwe.Herklotz@gmx.de

© Jerry Sulivan
tester
jhsulliv@comcast.net

© Giorgio Brausi
VS6Installer 
http://nuke.vbcorner.net/VS6Installer/tabid/125/language/en-US/Default.aspx

© UPX Packer
upx.exe
https://upx.github.io/

© Okrasa Ghia
FCalc
okrasaghia@yahoo.com

And much more.

# License:

GNU GPLv3

# Disclaimer
The Route-Riter source code and all software in this repository is provided for educational purposes ONLY. This repository is not affiliated with or endorsed by their respective copyright holders.
