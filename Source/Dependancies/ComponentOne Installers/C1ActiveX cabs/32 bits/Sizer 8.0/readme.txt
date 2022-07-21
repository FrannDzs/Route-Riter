////////////////////////////////////////////////////////////////////////
//
// C1Sizer 8 Maintenance History
//
////////////////////////////////////////////////////////////////////////


=========================================================================================       
C1Sizer.ocx Build Number 8.0.20173.155   Build Date: October 25, 2017
========================================================================================= 

 ** AboutBox info has been updated.

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20173.154   Build Date: October 23, 2017
========================================================================================= 

 ** AboutBox info has been updated.

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20173.153   Build Date: October 7, 2017
========================================================================================= 

 ** AboutBox info has been updated.

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20173.152   Build Date: October 3, 2017
========================================================================================= 

 ** AboutBox info has been updated.

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20163.151   Build Date: December 27, 2016
========================================================================================= 

 ** AboutBox design has been updated.

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20133.150   Build Date: December 12, 2013
========================================================================================= 

Corrections 
----------- 
 ** C1Elastic could interfere with ComboBox text selection in some cases (TFS 4277)

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20113.148   Build Date: December 15, 2011
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** Increased maximum number of caption labels from 80 to 1024 (TFS 16660)

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20113.147   Build Date: October 20, 2011
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V3/2011 build
 
Corrections 
----------- 
 ** Splitter bar left black trail on some system running Aero themes (TFS 14498)

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20111.146   Build Date: March 8, 2011
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V1/2011 build
 ** Built using VS2010 and the latest ATL implementation (with security adjustments for Vista/Win7)

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20101.145   Build Date: August 21, 2008
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V1/2010 build
 ** Built using VS2008 and the latest ATL implementation (with security adjustments for Vista/Win7)

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20093.143   Build Date: August 21, 2008
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V3/2009 drop

Corrections 
----------- 
- VSAwk did not detect design time correctly when used in MFC applications (5820)

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20081.142   Build Date: November 25, 2008
========================================================================================= 

Corrections 
----------- 
- Improved bounds-checking when setting tab captions (malicious code could cause IE to crash).
- Limited number of tabs to 30,000 (malicious code could cause IE to crash).

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20081.141   Build Date: September 17, 2008
========================================================================================= 

Corrections 
----------- 
- C1Awk was looking for help in C1Sizer.chm; changed that to Sizer8.chm to match 
  C1Elastic and C1Tab.

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20081.140   Build Date: May 13, 2008
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V1/2008 drop

- Rebuilt for "NX" compliance

  *** IMPORTANT: THIS UPGRADE IS REQUIRED FOR VISUAL STUDIO 2008 ***

  This new build leverages improvements made to the ATL library, including changes that
  make the control "NX" compatible.
  
  If you use previous versions of the control in Visual Studio 2008 projects, your 
  application will generate an access violation, due to Data Execution Prevention (DEP)
  (http://msdn2.microsoft.com/en-us/library/aa366553.aspx). This is because the previous
  versions of the control were based on ATL 3.0, which injected executable code in its
  data area ('thunks').

  If you try to add an older version of the control to a form designer in Visual Studio 2008, 
  you may see a misleading error like "Unable to get window handle for the AxMyATLCtrl
  control, windowless ActiveX controls are not supported".  The inner exception, if viewed, 
  would be more revealing: "Attempted to read or write protected memory. This is often an 
  indication that other memory is corrupt."  However, in this case, corruption is most 
  likely not the problem, but rather the attempt to execute code in NX memory.

  For more details, please see
  http://kbalertz.com/948468/Applications-Using-Older-Components-Experience-Conflicts.aspx  



=========================================================================================       
C1Sizer.ocx Build Number 8.0.20073.135   Build Date: December 18, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Improved licensing 


=========================================================================================       
C1Sizer.ocx Build Number 8.0.20072.41   Build Date: March 7, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V2/2007 drop

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20071.40   Build Date: January 14, 2007
========================================================================================= 

Corrections 
----------- 
 * 'ResizeFonts' property not functional on Labels when label's "AutoSize" property is set to "True"
    and there are other invisible controls present  (AXSZR000048) 

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20071.39   Build Date: October 2, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V1/2007 drop
 * Ampersand sign (&&) was not displayed when the C1Tab's Position property set to "tpLeft", and "tpRight" (AXSZR000050)
 * Application not responding when GridRows and GridColumns values set to very large values (> 20,000). (AXSZR000046)
 * 'Company Name' not included in C1SizerPpg.dll. (AXSZR000040)

=========================================================================================       
C1Sizer.ocx Build Number 8.0.20063.38   Build Date: June 15, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V3/2006 drop
 * changed help file name to Sizer8.chm (to avoid conflict with .NET version)

Corrections 
----------- 
 * none


=========================================================================================       
C1Sizer.ocx Build Number 8.0.20062.36   Build Date: February 16, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none

Corrections 
----------- 
 * TagLabel color did not reflect the Enabled state of controls without windows (e.g. Label)


=========================================================================================       
C1Sizer.ocx Build Number 8.0.20062.35   Build Date: January 21, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V2/2006 drop

Corrections 
----------- 
 * none


=========================================================================================       
C1Sizer.ocx Build Number 8.0.20061.35   Build Date: October 18, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V1/2006 drop

Corrections 
----------- 
 * None


=========================================================================================       
C1Sizer.ocx Build Number 8.0.20053.35   Build Date: June 7, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V3/2005 drop

Corrections 
----------- 
 * None


=========================================================================================       
C1Sizer.ocx Build Number 8.0.20052.35   Build Date: January 28, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V2/2005 drop
 * improved accessibility support (didn't work well with MS Accessibility Explorer)
 * improved behavior of off-screen paint buffer when used within IE

Corrections 
----------- 
 * None


=========================================================================================       
C1Sizer.ocx Build Number 8.0.20051.34   Build Date: October 28, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q1/2005 drop
 * Added hot mouse tracking to themed tabs
 
Corrections 
----------- 
 * Improved rendering of SizerOne with XP Themes. (AXSZR000019)
 * Improved sizing of TreeView and ListView controls when grid cell with (or height are set to zero. (AXSZR000023) 
   (there controls cannot be resized to zero width/height, they always remain at least one pixel in size).
 * Improved grid precision to account for round-off (AXSZR000024) 
 * Disable AutoScroll when MultiRow is set to true (AXSZR000025)



=========================================================================================       
C1Sizer.ocx Build Number 8.0.20044.33   Build Date: July 23, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q4/2004 drop
 
Corrections 
----------- 
 * none


=========================================================================================       
C1Sizer.ocx Build Number 8.0.20043.33   Build Date: April 29, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q3/2004 drop
 * use bold font for measurements only when BoldCurrent is set to true

Corrections 
----------- 
 * paint tab captions as disabled when the control is disabled


=========================================================================================       
C1Sizer.ocx Build Number 8.0.20042.32   Build Date: January 2, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q2/2004 drop

Corrections 
----------- 
 * GetTab returned zero if the control was invisible


=========================================================================================       
C1Sizer.ocx Build Number 8.0.20041.31   Build Date: October 22, 2003 
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
	New Settings for Appearance property <<DOC>>

		C1Elastic: Appearance = apXPThemes
			The new setting will make the control use XP themes if they 
			are available and enabled (see below).
			In the C1Elastic, XP themes affect the way in which progress
			bars and control borders are painted.
			If you select this style and XP themes are not available,
			then the control will be painted with a flat appearance.

		C1Tab: Appearance = tapXPThemes
			The new setting will make the control use XP themes if they 
			are available and enabled (see below).
			In the C1Tab, XP themes affect the way in which tabs are
			painted.
			If you select this style and XP themes are not available,
			then the control will be painted with a flat appearance.

		** NOTES ON XP THEMES:

		A visual style is included in the Windows XP release. In addition, 
		other themes or visual styles are available in the Windows XP Plus Pack. 
		You can use helper libraries and application programming interfaces 
		(APIs) to incorporate a Windows XP visual style into an application 
		with few code changes.

		Windows XP applies a visual style to the non-client (frame and caption) 
		area by default. To apply a visual style to common controls in the client
		area, you must use version 6 or later of the ComCtl32.dll file. ComCtl32.dll 
		version 6 is not a redistributable system component. ComCtl32.dll version 6 
		contains both the user controls and the common controls. By default, 
		applications use the controls that are defined in the User32.dll file. 
		In addition, applications use the common controls that are defined in 
		ComCtl32.dll version 5 by default.

		To use the Windows XP visual styles from an application, you must add an 
		application manifest file. This application manifest file should specify that
		ComCtl32.dll version 6 be used if it is available. One of the features that
		is included with this component is support for changing the appearance of 
		controls in a window.

		** Therefore, you must follow two steps to enable the Windows XP theme or visual
		style in Visual Basic 6.0: 

		1) Call the InitCommonControls functionAdd an application manifest file
		2) Add an application manifest file


		** EXAMPLE:

		1) Call the InitCommonControls Function

		You must call the InitCommonControls function in the Form_Initialize event:

		Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
		Private Sub Form_Initialize()
		    InitCommonControls
		End Sub
				
		NOTE: Do not call InitCommonControls in the Form_Load event. When you call 
		InitCommonControls from the Form_Load event, the form cannot load. 

		2) Add a manifest file to your application:

		You must add a file named YourApp.exe.manifest to the same folder as your 
		executable file. For example, if your application is named Generic.exe, 
		include a manifest file that is named Generic.exe.manifest. The application 
		manifest file has Extensible Markup Language (XML) format similar to the following:

		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
			<assemblyIdentity
			    version="1.0.0.0"
			    processorArchitecture="X86"
			    name="CompanyName.ProductName.YourApp"
			    type="win32"
			/>
		<description>Your application description here.</description>
		<dependency>
		    <dependentAssembly>
		        <assemblyIdentity
		            type="win32"
		            name="Microsoft.Windows.Common-Controls"
		            version="6.0.0.0"
		            processorArchitecture="X86"
		            publicKeyToken="6595b64144ccf1df"
		            language="*"
		        />
		    </dependentAssembly>
		</dependency>
		</assembly>
				
		After you place the application manifest file in the same folder as the executable 
		file, you can run the compiled executable file to display the Windows XP visual 
		style in the application.

		NOTE: You cannot view visual styles when you run the compiled executable from the 
		Visual Basic 6.0 Integrated Development Environment (IDE). 

		Although you can enable a Windows XP theme or visual style in Visual Basic 6.0 by 
		calling InitCommonControls and by using an application manifest file, Microsoft does 
		not officially support this feature.

		If you enable a Windows XP theme in Visual Basic 6.0, you may encounter unexpected 
		behavior. For example, if you place option buttons on top of a Frame control and then 
		enable a Windows XP theme or visual style, the option buttons on the Frame control
		appear as black blocks when you run the executable file. 

		You can also embed the manifest into the executable file. In this case, you won't 
		need the separate manifest file.

		The manifest file should be embedded into the executable using a resource editor. 
		The manifest should be embedded as a resource of type RT_RESOURCE and ID 1.

		For details on this procedure, please refer to MSDN (Using Themes with Windows XP).

Corrections 
----------- 
 ** Improved grid editor dialog to update dimension text box after toolbar clicks



8.0.20041.31 - October 22
    Q1 2004 drop
8.0.20034.31 - Aug 7
    Q4 2003 drop
8.0.20033.31 - June 17
    Fixed licensing expiration problem
8.0.20033.28 - May 6
    Q3 2003 drop
8.0.20032.28 - April 22
    Improved MDI fix in build 27
8.0.20032.27 - April 4
    Improved MDI fix in build 26
    Added support for ctl-Tab and shift-ctl-Tab keys to switch tabs
    Form_KeyUp was firing twice
8.0.20032.26 - Jan 20
    Q2 2003 drop
8.0.20031.26 - Dec 22
    Fixed ATL bug when unloading MDI form from event handler
    Fixed accessibility to register under Win95/NT
8.0.20031.25 - Dec 16
    Fixed design-time problem (tabs disabled after deleting contained Elastic)
    Increased the max number of child controls from 256 to 1024
    Synchronized last build # with version 7 for easier tracking
8.0.20031.5 - Nov 7
    Q1/2003 drop




////////////////////////////////////////////////////////////////////////
//
// What's new in C1Sizer version 8.0
//
////////////////////////////////////////////////////////////////////////


All controls
------------

Subscription licensing scheme, new About box, all incremental updates and fixes 
applied to previous versions. If you haven't been dowloading the latest patches
from our web site periodically, this is an easy way to get everything in one step.



C1Sizer, C1Tab
--------------

New properties:

string  AccessibleName           Gets or sets the name of the control used by accessibility client applications.
string  AccessibleDescription    Gets or sets the description of the control used by accessibility client applications.
string  AccessibleValue          Gets or sets the value of the control used by accessibility client applications.
Variant AccessibleRole           Gets or sets the role of the control used by accessibility client applications.

These new properties support Microsoft's Active Accessibility effort. Use them to make your 
applications more friendly to people with physical impairments, and to comply with US 
regulations.
