# Route-Riter

Route Riter is a set of tools for Microsoft Train Simulator created by Mike Simpson.

> **Note:** If you find a bug or have an improvement for the repository/project, please feel free to submit an inquiry or request a change. Your collaboration is greatly appreciated.

> **Note:** There are unused Forms and Modules in the source folder. If you have knowledge of MSTS, please help determine if these are needed or obsolete. Your collaboration is welcome.

## ⚠️ This repository is under reconstruction.

# 1. Prerequisites

- Download [Visual Basic 6.0](https://winworldpc.com/product/microsoft-visual-bas/60) or [Visual Studio 6.0](https://winworldpc.com/product/microsoft-visual-stu/60)
- Download [MSDN Library](https://winworldpc.com/product/msdn/vs-60) (Optional)
- Install Visual Studio 6 and MSDN Library.
- For Windows 10 64-bit users, use the [Visual Studio 6 Installer Wizard v5.0](https://github.com/FrannDzs/Route-Riter/blob/main-(7.6.26)/VS6InstallerSetup.exe) for installation. Refer to the [(Video Tutorial)](https://www.youtube.com/watch?v=1tkTb6AYlAg) or [Manual installation tutorial](https://www.codeproject.com/Articles/1191047/Install-Visual-Studio-on-Windows) for guidance.
- After installing the MSDN library, you can apply the following updates:
  - **MSDN Latest Version Update**:
    - **Disc1**: [Download](https://archive.org/details/MSDN_Library_October_2001_Disc_1)
    - **Disc2**: [Download](https://archive.org/details/MSDN_Library_October_2001_Disc_2)
    - **Disc3**: [Download](https://archive.org/details/MSDN_Library_October_2001_Disc_3)
- Install [Service Pack 6](https://web.archive.org/web/20120707200906/http://download.microsoft.com/download/1/9/f/19fe4660-5792-4683-99e0-8d48c22eed74/Vs6sp6.exe)
- Install [Microsoft Visual Basic 6.0 Service Pack 6 Cumulative Update](https://www.microsoft.com/en-us/download/details.aspx?id=7030)
- Install [Microsoft Visual Basic 6.0 Service Pack 6 Security Rollup Update](https://www.microsoft.com/en-us/download/details.aspx?id=50722)
- Install the C1 Controls manually using the [cabs files](https://github.com/FrannDzs/Route-Riter/tree/main-(7.6.26)/Source/Dependancies/ComponentOne%20Installers) or the [.ocx files](https://github.com/FrannDzs/Route-Riter/tree/main-(7.6.26)/Source/Dependancies/ComponentOne%20.ocx). Alternatively, you can use the [installers](https://github.com/FrannDzs/Route-Riter/tree/main-(7.6.26)/Source/Dependancies/ComponentOne%20Installers).
- If you encounter issues loading the `vsflex8l.ocx`, `vsprint.ocx`, and `c1sizer.ocx` controls in Visual Basic, try copying and registering them in both the `/system32` and `/SysWOW64` directories.

# 2. Recommended Add-ins/Tools for VB6 IDE

- [Code Advisor](https://www.microsoft.com/en-US/download/details.aspx?id=1222) (Free) - Optional
- [Code Help 3.0](https://github.com/clayreimann/CodeHelp) (Free) - Recommended for Tabs in VB IDE
- [Visual Basic 6 Mouse Wheel Fix](https://github.com/FrannDzs/Route-Riter/blob/main-(7.6.27)/Others/vb6mousewheelfix.exe) (Free) - Fixes Scroll Wheel in IDE
- [MZ-Tools 8.0](https://www.mztools.com/v8/mztools8.aspx) (Paid) - A Very Recommended Tools Suite
- [CodeSMART](https://www.axtools.com/products-codesmart-vb6.php) (Paid) - A Very Recommended Tools Suite
- [Codejock Suite Pro for ActiveX](https://codejock.com/products/suitepro/?2yn6s14z=p1z) (Paid) - Recommended Controls Suite
- [ModernVB](https://github.com/VykosX/ModernVB) (Free) - Modernizes your VB6 IDE (Very Recommended)
- [OLEEXP: Modern Shell Interfaces](https://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb) (Free)
- [VB Common Controls Replacement Library](https://github.com/Kr00l/VBCCR) (Free) - Remake of Common Controls (Recommended)
- [DX9VB](https://github.com/thetrik/DX9VB) (Free) - Direct3D9 for Visual Basic 6
- [VbTrickMultiThreading](https://github.com/thetrik/VbTrickThreading) (Free) - Module for working with multithreading in VB6
- [VbTrickTimer](https://github.com/thetrik/VbTrickTimer) (Free) - Timer class for VB6/VBA compatible with 64-bit Office
- [VB64bitDLLusage](https://github.com/thetrik/Vb64BitDllUsage) (Free) - Using 64-bit DLL in VB6 (in WOW64)
- [VBPNG](https://github.com/thetrik/VbPng) (Free) - Provides the ability to work with PNG/ICO/CUR/ANI images using the standard controls
- [DeleteVbwFiles](https://github.com/EduardoVB/VB6-AddIn-Delete-vbw-Files) (Free) - VB6 Add-In to delete .vbw files automatically when opening or closing a project
- [VB6NAMESPACES](https://github.com/WindowStations/VB6NameSpaces) (Free) - Enables interoperability with VBA/VB6, including VB.NET Forms, Controls, Properties, Events, and NameSpaces, instanced as nested class buckets
- [VB6 PORTER](https://github.com/VBForumsCommunity/VB6Porter) (Free) - Supports porting VB code forwards and backwards
- [VBSQlite](https://github.com/Kr00l/VBSQLite) (Free) - VB SQLite Library

# 3. Running the Compilation

1. Run Visual Basic 6.0 and choose the option to load an existing project.
2. Select `Route_Riter7.vbp` to load the project.
3. Make the necessary changes and press **F5** or **Ctrl + F5** to compile and run.
4. From the **Project** menu, choose **Properties** > **Compile** to select the compilation mode.
5. From the **File** menu, generate an .exe file.
6. Drop the compiled executable into the **Release** folder.

# 4. After Compiling the Project

1. Extract all compressed files from the [./Release/Dependancies](https://github.com/FrannDzs/Route-Riter/tree/main-(7.6.26)/Release/Dependancies) directory to the root of the **Release** folder.
2. Install [mwgfxdll.exe](https://github.com/FrannDzs/Route-Riter/blob/main-(7.6.26)/Release/mwgfxdll.exe).

# Development Plans

- Improve the current code following the original direction.

# Future Development Plans

- Prepare the Visual Basic 6.0 code for migration to Visual Basic .NET.
- Separate the application and data tiers into a separate DLL from the presentation.
- Change the user interface to an inductive user interface using [MSDN's Inductive User Interface Guidelines](https://msdn.microsoft.com/en-us/library/ms997506.aspx).
- Fully parse the Microsoft Train Simulator files by adapting an XML parser such as [pugixml](http://pugixml.org/).

# Credits/Acknowledgement

This project would not have been possible without the contributions and permissions from the following individuals:

- Mike Simpson - for writing the initial code and granting necessary permissions. Contact: virtualtrains@tpg.com
- Jeffrey Kraus - for donating his code and allowing this project to be revived. Contact: support@digital-rails.com
- Carl-Heinz Rave - TsUtils. Contact: mail@carloshr.de
- Scott Miller - AceIt. Contact: aceit@ameritech.net
- Martin Wright - TGATools2A. Contact: martin@mwgfx.co.uk
- Paul Gausden - Shape Viewer. Contact: [Decapod99 Blog](https://decapod99.wordpress.com/)
- Edward Grubb - PicFormat32. Contact: [Planet Source Code](https://github.com/Planet-Source-Code/edward-grubb-ed0-picformat32__1-13267)
- Franky Braem - SAWZipNG. Contact: [CodeProject](http://www.codeproject.com/Articles/875/SAWZip-zip-file-manipulation-control)
- Jean-loup Gailly - zlib.dll. Contact: [zlib.net](http://www.zlib.net/)
- ComponentOne - c1sizer.ocx, vsflex8l.ocx, and vsprint8.ocx. Contact: [ComponentOne](http://www.componentone.com/)
- Jordan Rusell Software - Inno Setup Installer. Contact: [JR Software](https://jrsoftware.org/)
- Uwe Herklotz - UHARC.EXE. Contact: Uwe.Herklotz@gmx.de
- Jerry Sulivan - tester. Contact: jhsulliv@comcast.net
- Giorgio Brausi - VS6Installer. Contact: [VB Corner](http://nuke.vbcorner.net/VS6Installer/tabid/125/language/en-US/Default.aspx)
- UPX Packer - upx.exe. Contact: [UPX](https://upx.github.io/)
- Okrasa Ghia - FCalc. Contact: okrasaghia@yahoo.com
- And many more.

# License

This project is licensed under GNU GPLv3.

# Disclaimer

The Route-Riter source code and all software in this repository are provided for educational purposes ONLY. This repository is not affiliated with or endorsed by their respective copyright holders.
