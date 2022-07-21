/////////////////////////////////////////////////////////////////////////////////////////
//
// Readme file for ComponentOne VSReport8 control
//
// VSRpt8.ocx
//
/////////////////////////////////////////////////////////////////////////////////////////


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20121.187   Build Date: May 24, 2012
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V1/2012 build

Corrections 
------------------------------------------- 
- Parameter dialog did not allow user to type strings longer than the textbox (TFS 8802)
- Section OnFormat and OnPrint properties did not show editor button on Property window (TFS 14927)

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20113.186   Build Date: October 6, 2011
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V3/2011 build

Corrections 
------------------------------------------- 
- Fixed an issue related to sub-report design time persistence

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20111.185   Build Date: Feb 22, 2011
========================================================================================= 

Corrections 
------------------------------------------- 
- Improved error-handling when rendering subreports []

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20111.184   Build Date: Jan 26, 2011
========================================================================================= 

Corrections 
------------------------------------------- 
 - Static link to VS2008 runtime library (instead of VS2010)
   This is required to support Win2000 and other older operating systems
   see http://support.microsoft.com/kb/2005279


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20111.183   Build Date: Jan 7, 2011
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V1/2011 build
 ** Built using VS2010 and the latest ATL implementation 

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20103.182   Build Date: October 27, 2010
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 - Removed dependency introduced in build 181 (VS2010)

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20102.181   Build Date: September 8, 2010
========================================================================================= 

Corrections 
----------- 
- Built using VS2010 and the latest ATL implementation
- Improved localization handling (version 177 caused some subtle changes in behavior)

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20102.178   Build Date: June 28, 2010
========================================================================================= 

Corrections 
----------- 
- Fixed problem in report designer's Group editor [11338]

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20101.177   Build Date: February 11, 2010
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V1/2010 build
 ** Built using VS2008 and the latest ATL implementation 
    (with security adjustments for Vista/Win7 and performance improvements)
 
=========================================================================================       
VSRpt8.ocx Build Number 8.0.20093.173   Build Date: October 29, 2009
========================================================================================= 

Corrections 
----------- 
- Fixed layout problem with CanGrow subreports (TFS 5005)

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20093.172   Build Date: August 11, 2009
========================================================================================= 

* V3/2009 build

Corrections 
----------- 
 - Marked the control as unsafe for inialization, in compliance with ATL Security Update
	http://msdn.microsoft.com/en-us/visualc/ee309358.aspx
   NOTE: This only affects scenarios where the control is used in Web Pages.
   NOTE: The actual modifications are in the IObjectSafetyImpl declaration.

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20082.171   Build Date: July 16, 2008
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- V2/2008 build
- Improved Pdf export of Japanese fonts


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20081.170   Build Date: May 13, 2008
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
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
VSRpt8.ocx Build Number 8.0.20081.158   Build Date: February 5, 2008
========================================================================================= 

Corrections 
----------- 
 * Font property was not being properly re-initialized when clearing the report definition

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20073.156   Build Date: December 18, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Improved licensing

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20073.155   Build Date: September 26, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Improved Pdf export of Japanese characters

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20073.154   Build Date: August 6, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V3/2007 build
 - Repeating group headers in subreports did not always align correctly (AXRPT000151)

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20072.153   Build Date: March 7, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V2/2007 build

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20071.152   Build Date: January 4, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Added new method RenderToFileWithLock. This is similar to RenderToFile, except it's
	thread-safe for use in ASP applications.
 * Changed report designer to always save the main report's Font name

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20071.151   Build Date: October 17, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V1/2007 build


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20063.151   Build Date: October 17, 2006
========================================================================================= 

Corrections 
----------- 
* Delta fixes:
  - Empty Html file being generated when exporting a report with image in subreport (AXRPT000143) 
  - ReportDesinger could hang when rendering a report with 'MaxPage' set to other number than 0 or actual pages (AXRPT000150) 
  - Fixed memory leak problem in ATL library


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20063.148   Build Date: September 6, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * fixed gdi leak in field object.

Corrections 
----------- 
 * Allow setting DataSource.Recordset property while rendering subreports.

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20063.147   Build Date: Jul 31, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * improved Designer About Box (do not display registration buttons in OEM editions)

Corrections 
----------- 
 * none

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20062.146   Build Date: May 11, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none

Corrections 
----------- 
 * fixed problem in field breaking that occasionally caused fields to render into
   page footer section (AXRPT000119)

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20062.145   Build Date: March 27, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V2/2006 build

Corrections 
----------- 
  * Fixed logic in footer fields that refer to footer fields in other sections (issue id 339-55)
  * Improved logic used for parsing percentages (convert from string back to double as needed, AXRPT000132) 
  * Improved rendering of subreports in page footer sections (AXRPT000136)


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20061.144   Build Date: December 7, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V1/2006

Corrections 
----------- 
 * Japanese version only: field fonts would sometimes change to [MS P Gothic].  (AXRPT000053) 
   (this was problem related to the font's CharSet)


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20053.143   Build Date: November 28, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none
 
Corrections 
----------- 
 * Script expressions using DataSource.Recordset did an extra AddRef and prevented 
   releasing the Recordset properly (sample usr_14301).

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20053.142   Build Date: July 12, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none
 
Corrections 
----------- 
 * fixed drawing borders around multi-column subreports
 * improved design time support for C++ projects under Visual Studio 7/8
   (used to show invisible controls at runtime)

=========================================================================================       
VSRpt8.ocx Build Number 8.0.20053.140   Build Date: July 12, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V3/2005 build
 
Corrections 
----------- 
 * improved Pdf export to handle curly quotes


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20052.139   Build Date: March 28, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * V2/2005 build
 
Corrections 
----------- 
 * improved rendering of background colors after column breaks (AXRPT000059)
 * improved handling of OnOpen/OnClose events for subreports (AXPRV000064)
 * improved field breaking logic (AXRPT000060)
 * improved rendering of hyperlinks (AXRPT000093)


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20051.137   Build Date: January 22, 2005
========================================================================================= 

Corrections 
----------- 
 * Using filters longer than 1024 characters caused exception


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20051.136   Build Date: November 17, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q1/2005 build
 * Honor section BackColor across page breaks within fields and sections (AXRPT000044)
 
Corrections 
----------- 
 * Improved db cursor handling when rendering PageFooter.  (AXPRV000045)
 * Render overlay fields (with page counts) with MaxPages.  (AXRPT000055)


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20044.135   Build Date: August 23, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** improved report designer keyboard handling for compatibility with VB/Visual Studio/Access:
    arrow keys move selected fields in 5-pixel increments
    CTRL arrow keys move selected fields in 1-pixel increments
    SHIFT arrow keys resize selected fields in 5-pixel increments
    SHIFT+CTRL arrow keys resize selected fields in 1-pixel increments

Corrections 
----------- 
 ** improved rendering of solid sections in multi-column reports
 ** fixed bug in report designer (Group editor show cursor next to two rows)
 ** fixed bug in report designer (could enable export/delete commands when no report available)


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20044.134   Build Date: July 21, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q4/2004 build

Corrections 
----------- 
 ** none


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20043.134   Build Date: July 21, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none

Corrections 
----------- 
 ** improved pdf output filter to include clipping and better precision on some GDI operations


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20043.133   Build Date: July 2, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none

Corrections 
----------- 
 ** fixed problem with group headers in multi-column (down layout) reports


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20043.131   Build Date: May 5, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q3/2004 drop

Corrections 
----------- 
 ** none


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20042.131   Build Date: April 25, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none

Corrections 
----------- 
 ** fix page header shift to right while rendering subreports


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20042.130   Build Date: March 22, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none

Corrections 
----------- 
 ** improved HTML rendering of sections with multiple cangrow subreports


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20042.126   Build Date: February 26, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q2/2004 drop

Corrections 
----------- 
 ** none


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20041.126   Build Date: February 8, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none

Corrections 
----------- 
 ** Fixed font export to pdf under Win98


=========================================================================================       
VSRpt8.ocx Build Number 8.0.20041.125   Build Date: February 2, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * none

Corrections 
----------- 
 ** Fixed gdi leak when rendering unbound images

// 2004 ////////////////////////////////////////////////////////////////////////////
8.0.20041.124 - October 23
    Q1/2004 build

// 2003 ////////////////////////////////////////////////////////////////////////////
8.0.20032.124 - September 2
    Fixed problem with RunningTotals (wasn't resetting properly at StartDoc)
8.0.20032.123 - May 28
    Added support for non-standard containers
8.0.20032.121 - May 9
    Fixed licensing expiration problem
    Fixed small memory leak (noticeable when creating several thousand instances of the control on server apps)
8.0.20032.120 - May 8
    Q3/2003 drop
8.0.20032.120 - May 7
    xxxx    Fixed problem with multiple CanGrow/Shrink subreport fields in a single section
8.0.20032.117 - April 29
    Improved loading of images from database (take packaged or raw gif, jpg, etc)
8.0.20032.116 - April 12
	xxxx	Major speed optimization for subreports (smarter DataSource handling)
8.0.20032.115 - March 31
    xxxx    Fixed persistence for subreports with names that contain ampersands (wasn't encoding)
    xxxx    Fixed problem with IIF statements in Access import
8.0.20032.114 - March 10
    xxxx    Improved handling of Date fields in international Parameter dialog
8.0.20032.113 - March 5
    7635    Improved handling of foreign recordsets in subreports
    xxxx	Improved loading of embedded images in international mdb files
8.0.20032.112 - Feb 20
    Q2/2003 drop
8.0.20031.112 - Jan 26
    xxxx    Fixed licensing issue when rendering reports to file without VSPrinter
8.0.20031.111 - Jan 13
    xxxx    Allow rendering raw image fields (Dispatch as well as packaged OLE objects)


// 2002 ////////////////////////////////////////////////////////////////////////////

8.0.20031.110 - Dec 26, 2002
    Extra check when loading pictures
8.0.20031.109 - Dec 20, 2002
    Fixed licensing for Designer redistribution
    Added extra check for opening Oracle data sources
8.0.20031.106 - Dec 9, 2002
    Fixed column break problem within multi-column subreports
    Fixed License button in About Box 
8.0.20024.104 - Dec 4, 2002
    Improved error checking on RenderToFile/Cancel
    Synchronized last version number to V7
8.0.20024.10 - Nov 11, 2002
    Fixed Tag persistence problem
8.0.20024.8 - Oct 30, 2002
    Fixed loading deep-nested subreports at design time
8.0.20024.7 - Oct 21, 2002
    Fixed problem with subreport grouping/aggregates
    Added RTF export option (control and designer)



////////////////////////////////////////////////////////////////////////////////////
//
// What's new in VSView Reporting Edition 8.0
//
////////////////////////////////////////////////////////////////////////////////////


Subscription licensing scheme, new About box, all incremental updates and fixes 
applied to previous versions. If you haven't been dowloading the latest patches
from our web site periodically, this is an easy way to get everything in one step.


VSReport8 control
-----------------

1) ZOrder

    Field object has a new ZOrder property (type long) and SetZOrder method that allow
    you to control the field's rendering order, so when two fields overlap you can 
    determine which one is 'in front'.

    The SetZOrder method adjusts the ZOrder property automatically, so there's rarely
    any need to set the ZOrder property manually.

2) Hyperlinks

    The Field object now has a LinkTarget property that allows you to specify a URL
    to be visited when the link is clicked (in VSPrinter, HTML, or PDF reports).

    Here's some more details for the online help:

    Syntax:
    string Field.LinkTarget

    Description:
    Gets or sets an expression that evaluates to a URL to be visited when the field is clicked.

    Applies to:
    Field object

    Notes:

    If the LinkTarget is set to a non-empty string, the control will render the field as
    usual (using the Text and Picture properties), but the field will work as a hyperlink.

    When the user clicks on the field, the specified URL will be opened. If the report is 
    being viewed in a VSPrinter control, remember to set the VSPrinter.AutoLinkNavigate
    property to True so the links will work automatically. If the report is exported to 
    HTML or PDF, the links are translated automatically.
    
    The LinkTarget property is always calculated. The Text property, on the other hand, 
    is calculated only if the Calculated property is set to true.

    Examples:

    hdrField.Calculated = False
    hdrField.Text = "Click here to visit the ComponentOne home page"
    hdrField.LinkTarget = "http://www.componentone.com"

    field.Calculated = True
    field.Text = "ProductName"
    field.LinkTarget = """http://www.componentone.com/products/"" & ProductID"

3) RTF export

    The RenderToFile method now accepts a 'vsrRTF' setting to export RTF reports. The designer
    also supports RTF export.

4) Tag persistence

    The control now persists the Tag property for Fields, Sections, and Groups.

5) Anchor property (Field object)

    This was added to Version 7, but not documented.

    Syntax:
    AnchorSettings Field.Anchor

    Description:
    Gets or sets how the field position and dimensions should be affected when the containing section changes size
    as a result of CanGrow or CanShrink.

    Applies to:
    Field object

    Notes:

    When a section grows or shrinks as a result of processing the CanGrow and CanShrink properties, fields usually
    retain their original position ans size (Top and Height properties). 

    In some cases, you may want to force fields to grow with the section, regardless of their content. For example, 
    you may have vertical lines that should span the height of the section, or horizontal borders that should be 
    positioned with respect to the bottom of the section. You can achieve this with the Anchor property.

    Valid settings for this property are:

    AnchorSettings:
          vsrATop
              Default. When the section grows, the distance between the top of the field and the top of the section 
              remains the same.
          vsrABottom
              When the section grows, the field is pushed down. The distance between the bottom of the field and
              the bottom of the section remains the same.
          vsrATopAndBottom
              When the section grows, the field is stretched vertically. The distance between the top of the field 
              and the top of the section remain the same, and so does the distance between the bottom of the field
              and the bottom of the section.

    The diagram below shows the effect of the Anchor property on a field.

    BEFORE
    ============================================================ section top 
    +----------------+ +----------------+ +----------------+
    |    vsrATop     | |   vsrABottom   | |vsrATopAndBottom|
    +----------------+ +----------------+ +----------------+
    ============================================================ section bottom

    If this section grew as a result of processing the CanGrow property, it would be rendered like this:

    AFTER
    ============================================================ section top 
    +----------------+                    +----------------+
    |    vsrATop     |                    |                |
    +----------------+                    |                |
                                          |                | 
                                          |vsrATopAndBottom|
                                          |                |
                       +----------------+ |                |
                       |   vsrABottom   | |                |
                       +----------------+ +----------------+
    ============================================================ section bottom

    

VSReport8 Designer
------------------

The designer has new toolbar buttons and menu options that allow you to adjust the
field's ZOrder, bringing them to the front or sending them to the back.
