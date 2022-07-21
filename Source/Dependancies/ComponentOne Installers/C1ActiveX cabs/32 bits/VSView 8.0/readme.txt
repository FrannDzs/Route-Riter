////////////////////////////////////////////////////////////////////////////////
//
// VSView 8.0 Maintenance log
//
////////////////////////////////////////////////////////////////////////////////

=========================================================================================       
VSView8 Build Number 8.0.20183.178   Build Date: February 15, 2018
========================================================================================= 

- Address negative heights in WMF/EMF embedded images for PDF conversions.

=========================================================================================       
VSView8 Build Number 8.0.20173.177   Build Date: November 30, 2017
========================================================================================= 

- New AboutBox information.

=========================================================================================       
VSView8 Build Number 8.0.20173.176   Build Date: October 25, 2017
========================================================================================= 

- New AboutBox information.

=========================================================================================       
VSView8 Build Number 8.0.20173.175   Build Date: October 23, 2017
========================================================================================= 

- New AboutBox information.

=========================================================================================       
VSView8 Build Number 8.0.20173.174   Build Date: October 7, 2017
========================================================================================= 

- New AboutBox information.

=========================================================================================       
VSView8 Build Number 8.0.20173.173   Build Date: October 3, 2017
========================================================================================= 

- New AboutBox information.

=========================================================================================       
VSView8 Build Number 8.0.20163.172   Build Date: December 27, 2016
========================================================================================= 

- New AboutBox design.

=========================================================================================       
VSView8 Build Number 8.0.20132.171   Build Date: March 31, 2014
========================================================================================= 

Corrections 
------------------------------------------- 
- Improved PDF font rendering (TFS 51822)

=========================================================================================       
VSView8 Build Number 8.0.20132.170   Build Date: July 11, 2013
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- V2/2013 build

Corrections 
------------------------------------------- 
- Improved RTF export logic for table borders (TFS 39637)

=========================================================================================       
VSView8 Build Number 8.0.20121.166   Build Date: February 14, 2011
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- V1/2012 build
- Ignore link targets with invalid lengths (zero or greater than 1024 chars)

=========================================================================================       
VSView8 Build Number 8.0.20113.165   Build Date: October 13, 2011
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- V3/2011 build
- Create temp files in the doc folder instead of the temp folder (TFS 16277)

=========================================================================================       
VSView8 Build Number 8.0.20112.164   Build Date: August 26, 2011
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- V2/2011 build

=========================================================================================       
VSView8 Build Number 8.0.20111.162   Build Date: March 21, 2011
========================================================================================= 

Corrections 
------------------------------------------- 
- Fixed licensing issue with VB6's Licenses collection and the VSPDF control.

=========================================================================================       
VSView8 Build Number 8.0.20111.161   Build Date: January 26, 2011
========================================================================================= 

Corrections 
------------------------------------------- 
 - Static link to VS2008 runtime library (instead of VS2010)
   This is required to support Win2000 and other older operating systems
   see http://support.microsoft.com/kb/2005279


=========================================================================================       
VSView8 Build Number 8.0.20111.160   Build Date: January 6, 2011
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- V1/2011 build
- upgraded to VS2010 and the latest ATL libraries
- Improved VSPdf rendering of metafiles (especially VSFlex images with cell borders) [TFS 13561]

Corrections 
----------- 
- Fixed VSPdf clipping issue [TFS 10612]
- Fixed issue with optional parameters [TFS 14121]

=========================================================================================       
VSView8 Build Number 8.0.20101.156   Build Date: January 25, 2010
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
* Improved table shading logic to avoid gaps between cells

=========================================================================================       
VSView8 Build Number 8.0.20101.155   Build Date: January 11, 2010
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V1/2010 build
 ** Built using VS2008 and the latest ATL implementation (with security adjustments for Vista/Win7)

=========================================================================================       
VSView8 Build Number 8.0.20093.150   Build Date: October 28, 2009
========================================================================================= 

Corrections 
----------- 
- AfterUserPage event did not fire when tracking the scrollbar 

=========================================================================================       
VSView8 Build Number 8.0.20093.149   Build Date: August 11, 2009
========================================================================================= 

* V3/2009 build

Corrections 
----------- 
 - Marked the control as unsafe for inialization, in compliance with ATL Security Update
	http://msdn.microsoft.com/en-us/visualc/ee309358.aspx
   NOTE: This only affects scenarios where the control is used in Web Pages.
   NOTE: The actual modifications are in the IObjectSafetyImpl declaration.


=========================================================================================       
VSView8 Build Number 8.0.20083.148   Build Date: December 15, 2008
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Improved PDF export of Japanese quotes.


=========================================================================================       
VSView8 Build Number 8.0.20082.147   Build Date: July 14, 2008
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Improved PDF export of Japanese text.


=========================================================================================       
VSView8 Build Number 8.0.20082.146   Build Date: July 3, 2008
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V2/2008 build

 * Now uses EnumPrinters API instead of older GetProfileString 
   GetProfileString is not supported under some specific scenarios such as Vista services
   (issue TTP ID: 17326)


=========================================================================================       
VSView8 Build Number 8.0.20081.145   Build Date: May 13, 2008
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V1/2008 build

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
VSView8 Build Number 8.0.20073.135   Build Date: December 18, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
* Improved licensing

=========================================================================================       
VSView8 Build Number 8.0.20073.134   Build Date: August 28, 2007
========================================================================================= 

Corrections 
----------- 
- Improved Pdf export of Japanese characters

=========================================================================================       
VSView8 Build Number 8.0.20073.133   Build Date: August 6, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V2/2007 build
 - Text property not rendering some single-character strings (e.g. bullets, checkboxes, tabs) (AXPRV000161, AXPRV000155)
 - Under Vista, scrollbars flickered if the user moves the mouse over the toolbar buttons (AXPRV000151)


=========================================================================================       
VSView8 Build Number 8.0.20072.132   Build Date: March 16, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V2/2007 build
 * Support AngleArc command in VSPdf filter (AXPRV000149).

=========================================================================================       
VSView8 Build Number 8.0.20071.130   Build Date: December 5, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V1/2007 build

=========================================================================================       
VSView8 Build Number 8.0.20063.130   Build Date: December 5, 2006
========================================================================================= 

Corrections 
----------- 
 * Fixed GDI leak under IIS
 * Rendering some text files could result in strange chars appearing
   (in systems without Asian language support installed) (417-55)


=========================================================================================       
VSView8 Build Number 8.0.20063.128   Build Date: October 17, 2006
========================================================================================= 

Corrections 
----------- 
 * Delta fixes
    - VSPdf did not render some raster fonts correctly. (AXPRV000135)
	- Fixed memory leak in ATL library


=========================================================================================       
VSView8 Build Number 8.0.20063.125   Build Date: September 22, 2006
========================================================================================= 

Corrections 
----------- 
 * Clean up empty temp file created by SaveDoc method (AXPRV000145)


=========================================================================================       
VSView8 Build Number 8.0.20063.124   Build Date: August 2, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V3/2006 build
 * Improved text positioning in VSPdf filter. The improvement is noticeable with certain
   barcode fonts (AXPRV000127)

Corrections 
----------- 
 * none


=========================================================================================       
VSView8 Build Number 8.0.20062.124   Build Date: March 27, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V2/2006 build

Corrections 
----------- 
 * VSPdf: text color was not being handled correctly in all cases (Delta id AXPRV000135)


=========================================================================================       
VSView8 Build Number 8.0.20061.123   Build Date: December 15, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none

Corrections 
----------- 
 * VSPdf:
   - Text in PDF document was rendered with wrong color.  (AXPRV000129) 
     (note: this happened only when the text was rendered immediately after an opaque clipped rectangle)
 * VSDraw:
   - SaveDoc method caused memory leak  (AXPRV000125) 


=========================================================================================       
VSView8 Build Number 8.0.20061.122   Build Date: December 7, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V1/2006 build


=========================================================================================       
VSView8 Build Number 8.0.20053.121   Build Date: October 7, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * VSPrinter: load TrackMouseEvent proc dynamically to allow registration under Win95


=========================================================================================       
VSView8 Build Number 8.0.20053.120   Build Date: July 12, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V3/2005 build

Corrections 
----------- 
 * VSPrinter:
    - Improved TableCell propery to allow setting really long strings (use heap instead of stack)
 * VSPdf: 
    - Improved rendering of EMR_PIE metafile records (when using the same point, should draw full ellipse) (AXPRV000098) 
    - Improved clipping of metafile PolyPolygons and PolyBeziers (AXPRV000086)
    - Improved rendering partial table rows (AXPRV000093)
 * VSDraw:
    - DrawCircle method didn't automatically invalidate the control (AXPRV000074)


=========================================================================================       
VSView8 Build Number 8.0.20052.118   Build Date: May 20, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none

Corrections 
----------- 
 * VSPdf: improved string clipping


=========================================================================================       
VSView8 Build Number 8.0.20052.117   Build Date: January 17, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * T2/2005 build
 * Added support for full-justification to C1Pdf
 * Improved accessibility support

Corrections 
----------- 
 * VSPrinter: improved rendering of borders on tables with merged cells and solid bkg
 * VSPrinter: improved rendering of tables with KeepWithNext (didn't work on second-to-last row, AXPRV000066)
 * VSPdf: improved kerning for Japanese fonts
 

=========================================================================================       
VSView8 Build Number 8.0.20051.116   Build Date: November 16, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q1/2005 build
 * Allow use of scrollbars to page through document in preview mode (like Adobe Reader)
 * Updated NavBar appearance to use flat styles

Corrections 
----------- 
 * VSDraw: allow setting ScaleWidth/ScaleHeight to small values (>= 2 is now OK)


=========================================================================================       
VSView8 Build Number 8.0.20044.115   Build Date: September 16, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none

Corrections 
----------- 
 - VSPrinter
 * fixed problem with VSPrinter.TableCell(tcRowKeepTogether): in some cases, 
   when the space left at the bottom of the page was close but not enough to fit a 
   single line, the control could skip the row.
 * fixed problem that caused TableCell(tcBackColor/tcForeColor) to return wrong values
 
 - VSPdf
 * fixed problem that prevented wide VSPrinter borders from rendering to Pdf



=========================================================================================       
VSView8 Build Number 8.0.20044.114   Build Date: July 24, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q4/2004 build


Corrections 
----------- 
 * none


=========================================================================================       
VSView8 Build Number 8.0.20043.114   Build Date: July 24, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * improved precision in some pdf GDI operations


Corrections 
----------- 
 * fixed text clipping problem introduced in VSPdf 113


=========================================================================================       
VSView8 Build Number 8.0.20043.113   Build Date: May 24, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q3 drop
 * support for full RTF justification in VSView
   no new properties are required. use a fully-justified rtf string (\qj) or
   set the TextAlign property to one of the justified settings, e.g:
        Private Sub Command1_Click()
        With Me.VSPrinter1
            Dim s$
            s = "This is a long string that is supposed to render as {\b RTF}. "
            s = s & s & s
            s = s & s & s
            s = s & s & s
            s = s & s & s
    
            .StartDoc
            .TextAlign = taJustTop ' << use full justification
            .TextRTF = s
            .EndDoc
        End With
        End Sub
 * support for clipping in VSPdf


Corrections 
----------- 
 * fixed pdf rendering of metafile arcs


=========================================================================================       
VSView8 Build Number 8.0.20042.112   Build Date: May 5, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none

Corrections 
----------- 
 ** VSPrinter: handle arrow keys when hosted in user controls
 ** VSPrinter: better handling of bad urls set at design time: 
               the control will fire the Error event but no fatal exception
 ** VSPrinter: slightly better check of table row heights when KeepTogether = false
 ** VSPdf: don't try to map raster fonts (e.g MS Sans Serif) as TrueType
     

=========================================================================================       
VSView8 Build Number 8.0.20042.111   Build Date: February 16, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q2/2004 build

Corrections 
----------- 
 ** none
     

=========================================================================================       
VSView8 Build Number 8.0.20041.111   Build Date: February 8, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** Improved handling of some fonts (e.g. MS Sans Serif) in pdf output
 ** Improved font handling for pdf output under Win98
     

=========================================================================================       
VSView8 Build Number 8.0.20041.109   Build Date: December 16, 2003 
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** Improved appearance of non-solid lines in pdf output

Corrections 
----------- 
 ** none
     

2004 ///////////////////////////////////////////////////////////////////////////
*** 8.0.20041.108 - October 22
    Q1/2004 drop

2003 ///////////////////////////////////////////////////////////////////////////
*** 8.0.20034.108 - October 17
    Added support for 32-bpp images in VSPdf export
    Improved handling of temp files (improved threading efficiency/reliability)
*** 8.0.20034.107 - Aug 28
    Added support for non-standard containers
*** 8.0.20034.106 - Aug 7
    q4/2003 drop
*** 8.0.20032.105 - June 10
    Fixed license expiration problem
*** 8.0.20032.104 - June 4
    Fixed problem with default size of "Display" device (8.5x11")
*** 8.0.20032.103 - May 16
	Export links to RTF
	Fixed exporting DBCS strings
	Fixed problem with MouseLink event (was firing too often)
*** 8.0.20032.103 - Apr 2
    Added check in save to file to guarantee reset on ReadyState
    Improved precision of table cell fills
*** 8.0.20032.102 - Feb 21
    Fixed problem with FocusTrack in VSViewPort
    Improved initialization of TableSep property in MFC view objects
*** 8.0.20032.100 - Feb 20
    q2/2003 drop
*** 8.0.20031.100 - Jan 27
    xxxx    Synchronized build number with version 7
    7523    Fixed licensing issue with VSPDF8 and CreateFromFile
    7517    Fixed problem rendering RTF text with Paragraph property
    7381    Fixed problem in ClientToPage method in documents with mixed orientation

2002 ///////////////////////////////////////////////////////////////////////////
*** 8.0.20031.11 - Dec 22
    Fixed accessibility to register under Win95/NT

*** 8.0.20031.9 - Nov 28
    PDF: Added support for PolyBeziers

*** 8.0.20031.8 - Nov 7
    2003/Q1 build

*** 8.0.20024.8 - Nov 4
    Fixed exporting bullet chars to PDF
    Improved precision of PDF underlines
    



////////////////////////////////////////////////////////////////////////////////
//
// What's new in VSView 8.0
//
////////////////////////////////////////////////////////////////////////////////


All controls
------------

Subscription licensing scheme, new About box, all incremental updates and fixes 
applied to previous versions. If you haven't been dowloading the latest patches
from our web site periodically, this is an easy way to get everything in one step.



VSPrinter8, VSDraw8, VSViewPort8 controls
-----------------------------------------

New properties:

string  AccessibleName           Gets or sets the name of the control used by accessibility client applications.
string  AccessibleDescription    Gets or sets the description of the control used by accessibility client applications.
string  AccessibleValue          Gets or sets the value of the control used by accessibility client applications.
Variant AccessibleRole           Gets or sets the role of the control used by accessibility client applications.

These new properties support Microsoft's Active Accessibility effort. Use them to make your 
applications more friendly to people with physical impairments, and to comply with US 
regulations.



VSPDF8 control
--------------

New properties:

string Author       Gets or sets the name of the person who created the PDF document.
string Creator      Gets or sets the name of the application that created the PDF document.
string Title        Gets or sets the title of the PDF document.
string Subject      Gets or sets the subject of the PDF document.
string Keywords     Gets or sets the keywords associated with the PDF document.

These new properties allow you to add information to the PDF files you create. For example,
you can add keywords and make the files easier to retrieve.


Hyperlinks: The VSPDF8 controls supports hyperlinks. Use the new AddLink method in VSPrinter
to add links to a document, then export to PDF as usual. Users will be able to click the links
and move to other URLs or locationsd within the document. See VSPrinter8 notes below for more 
details.

string LinkTag, AnchorTag

These new properties allow you to customize the mechanism used to interpret link and anchor
tags embedded in VSPrinter documents. They are conceptually similar to the OutlineTag property
that already existed in version 7 of the control.



VSPrinter8 control
------------------

Support for Hyperlinks, including HTML and PDF export. Hyperlinks support allows you to build 
navigation into your documents. Hyperlink support consists of the following new object model
elements, detailed below:

1) AddLink and AddLinkTarget methods.
2) AutoLinkNavigate property.
3) MouseLink event.


** AddLink Method

	Syntax: AddLink(LinkText As String, LinkTarget As String, Formatted As Boolean)

	Adds a hyperlink to the document.

	LinkText is the text that will be displayed in the document.

	LinkTarget is a URL (e.g. "http://www.componentone.com") or local reference (e.g. "#myTarget")
	that defines the link target (destination).

	Formatted defines whether the text should be automatically formatted to look like an HTML link
	(underlined and rendered in blue).

	If the AutoLinkNavigate property is set to true, moving the mouse over hyperlinks will change
	the mouse pointer into a hand, and clicking it will either open the link target in a new window
	or bring the target into view.

	If you add links to local references, make sure you add the link target to the document using the
	AddLinkTarget method.

	Hyperlinks are automatically exported to PDF and HTML.

	Note: Internally, hyperlinks are represented as custom tags with the following format: 
	"%PDFLink|<linkTarget>". If the linkTarget starts with a pound sign "#" then the link is a local 
	reference, otherwise it is a URL or file name. If you want, you can create links using the 
	StartTag and EndTag methods to add tags using the specified format.

** AddLinkTarget Method

	Syntax: AddLinkTarget(TargetText As String, TargetName As String)

	TargetText is the text that will be displayed in the document.

	TargetName is the name of the target which can be reached by clicking on hyperlinks that reference it.

	Hyperlinks that point to local references should precede the target name with a pound sign "#". The 
	target name itself should not include the pound sign. For example:

	vp.AddLinkTarget "Home"
	GenerateDocumentContent vp ' << your routine
	vp.AddLink "Go back home", "#Home", True

	Note: Internally, link targets are represented as custom tags with the following format: 
	"%PDFName|<linkTarget>". If you want, you can create link targets using the StartTag and EndTag 
	methods to add tags using the specified format.

** AutoLinkNavigate Property

	Syntax: boolean AutoLinkNavigate

	Gets or sets whether the control should automatically detect hyperlinks embedded in the preview document,
	fire the MouseLink event, and change the cursor or perform the navigation automatically.


** MouseLink Event

	Syntax: Private Sub MouseLink(Link As String, Clicked As Boolean, ByRef Cancel As Boolean)

	Fired when the user clicks or or moves the mouse over a hyperlink (see AutoLinkNavigate property).

	Event parameters:
	Link: string containing the link target (e.g. "http://componentone.com")
	Clicked: true if the user clicked the link, false if he just moved the mouse over it.
	Cancel: set to true to prevent navigation to the link or changing the mouse cursor to show a hand.
