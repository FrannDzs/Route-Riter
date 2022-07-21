////////////////////////////////////////////////////////////////////////////////////////////////////
//
// Readme file for VSFlex 8.0
//
////////////////////////////////////////////////////////////////////////////////////////////////////

=========================================================================================       
VSFlex8 Build Number 8.0.20173.318 Build Date: March 14, 2018
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- String date values are now interpretted correctly based on locale. (TFS-308921)
  This issue was introduced in build 311.

=========================================================================================       
VSFlex8 Build Number 8.0.20173.317 Build Date: January 23, 2018
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- Added secondary general purpose buffer to elimate issue with cell merging (TFS-302185)

=========================================================================================       
VSFlex8 Build Number 8.0.20173.316 Build Date: November 30, 2017
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- Updated Japanese AboutBox

=========================================================================================       
VSFlex8 Build Number 8.0.20173.315 Build Date: October 25, 2017
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- Updated AboutBox

=========================================================================================       
VSFlex8 Build Number 8.0.20173.314 Build Date: October 23, 2017
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- Updated AboutBox

=========================================================================================       
VSFlex8 Build Number 8.0.20173.313 Build Date: October 7, 2017
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- Updated AboutBox

=========================================================================================       
VSFlex8 Build Number 8.0.20173.312 Build Date: October 2, 2017
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- Updated AboutBox

=========================================================================================       
VSFlex8 Build Number 8.0.20172.311 Build Date: July 07, 2017
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- none

Corrections 
------------------------------------------- 
- The XLFilter and sort has improved time conversions.  Previously, some time conversions
  with leading zero were not converted correctly leading to improper sorting.  TFS-269506.

=========================================================================================       
VSFlex8 Build Number 8.0.20171.310 Build Date: April 28, 2017
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- none

Corrections 
------------------------------------------- 
- The XLFilter now has improved time conversions.  Previously, some time conversions
  would alway result in ~00:00:00 times.  TFS-256130.

=========================================================================================       
VSFlex8 Build Number 8.0.20163.309 Build Date: January 10, 2017
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- none

Corrections 
------------------------------------------- 
- VSFlexGrid (OLEDB) once again honors Recordset filtering.  This function was lost in
  builds 307 and 308 due to build environment issues which have been corrected. TFS-228097.

=========================================================================================       
VSFlex8 Build Number 8.0.20163.308 Build Date: December 27, 2016
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- Updated AboutBox

=========================================================================================       
VSFlex8 Build Number 8.0.20162.307 Build Date: August 08, 2016
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- FlexGrid now always marked as Safe For Initialization.  This allows for usage within
  Microsoft Office products without unnecessary user notification.

=========================================================================================       
VSFlex8 Build Number 8.0.20162.306 Build Date: July 21, 2016
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V2/2016 build

Corrections 
------------------------------------------- 
- Corrected ColImageList handle type in 64-bit builds

=========================================================================================       
VSFlex8 Build Number 8.0.20161.305 Build Date: January 6, 2016
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V1/2016 build

Corrections 
------------------------------------------- 
- Wheel scrolling did not work if mouse was set to scroll one screen at a time (TFS 141385)

=========================================================================================       
VSFlex8 Build Number 8.0.20141.304 Build Date: October 28, 2015
========================================================================================= 

Corrections 
------------------------------------------- 
- Rendering RTL text was broken in build 255, correct behavior restored in 304 (TFS 133780)

=========================================================================================       
VSFlex8 Build Number 8.0.20141.303 Build Date: July 15, 2015
========================================================================================= 

Corrections 
------------------------------------------- 
- Rendering grids without gridlines to VSPrinter could fail to update the cursor position.

=========================================================================================       
VSFlex8 Build Number 8.0.20141.302 Build Date: Oct 9, 2014
========================================================================================= 

Corrections 
------------------------------------------- 
- ADO could throw exceptions when adding new records.

=========================================================================================       
VSFlex8 Build Number 8.0.20141.300 Build Date: Feb 7, 2014
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V1/2014 build

Corrections 
------------------------------------------- 
- Fixed buffer overrun vulnerability in ComboList/ColComboList properties (TFS 50857)

=========================================================================================       
VSFlex8 Build Number 8.0.20132.297 Build Date: September 13, 2013
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V2/2013 build

Corrections 
------------------------------------------- 
- Fixed bug when exporting long strings (len > 8k) to XLS (TFS 40127)

=========================================================================================       
VSFlex8 Build Number 8.0.20121.296 Build Date: July 16, 2012
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V2/2012 build

Corrections 
------------------------------------------- 
- Fixed Unicode export to XLS (TFS 19987, 17763)

=========================================================================================       
VSFlex8 Build Number 8.0.20121.295 Build Date: May 25, 2012
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V1/2012 build

Corrections 
------------------------------------------- 
- Fixed painting issue in some WinXP/7 themes (TFS 18281)
- Fixed caret issue when selecting and calling EditCell multiple times in the same function (TFS 19286)

=========================================================================================       
VSFlex8 Build Number 8.0.20113.291 Build Date: November 25, 2011
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V3/2011 build

Corrections 
------------------------------------------- 
- ScrollTip did not always appear (TFS 17919)
- Improved scrollbar extent calculation (TFS 18207)

=========================================================================================       
VSFlex8 Build Number 8.0.20112.288 Build Date: October 10, 2011
========================================================================================= 

Corrections 
------------------------------------------- 
- Addressed binding errors that occur when repeatedly binding under XP (TFS 16995)
- Fixed currency formatting issue with Euro sign (TFS 17590)

=========================================================================================       
VSFlex8 Build Number 8.0.20112.285  Build Date: September 2, 2011
========================================================================================= 

Corrections 
------------------------------------------- 
- Mouse events did not fire correctly when resizing or freezing rows/columns (TFS 15393)
- Binding could fail when using large data sources (TFS 15043, 16995)
- Error opening saved excel files with over 32767 rows (TFS 16821)

=========================================================================================       
VSFlex8 Build Number 8.0.20112.284  Build Date: August 19, 2011
========================================================================================= 

Corrections 
------------------------------------------- 
- Improved performance of SaveGrid method with SaveLoadSettings.flexFileExcel option (TFS 16664)

=========================================================================================       
VSFlex8 Build Number 8.0.20112.283  Build Date: August 5, 2011
========================================================================================= 

Corrections 
------------------------------------------- 
- EditWindow property returned NULL immediately after calling EditCell (TFS 16050)

=========================================================================================       
VSFlex8 Build Number 8.0.20112.282  Build Date: June 10, 2011
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V2/2011 build

=========================================================================================       
VSFlex8 Build Number 8.0.20111.281  Build Date: February 26, 2011
========================================================================================= 

Corrections 
------------------------------------------- 
- Fixed issue with hand cursor (TFS 14237)

=========================================================================================       
VSFlex8 Build Number 8.0.20111.280  Build Date: January 26, 2011
========================================================================================= 

Corrections 
------------------------------------------- 
 - Static link to VS2008 runtime library (instead of VS2010)
   This is required to support Win2000 and other older operating systems
   see http://support.microsoft.com/kb/2005279

=========================================================================================       
VSFlex8 Build Number 8.0.20103.276  Build Date: Dec 23, 2010
========================================================================================= 

Corrections 
------------------------------------------- 
 - Restored chaptered rowset support required for hierarchical and filtered rowsets
   (broken in 264 with the latest ATL update) (TFS 13539, 13471)
 - SaveExcel method broken in 264 (TFS 13812)
 - Outlining broken in build 263 (maximum number of subtotal levels back to 128) (TFS 13708)

=========================================================================================       
VSFlex8 Build Number 8.0.20103.273  Build Date: Dec 3, 2010
========================================================================================= 

Corrections 
------------------------------------------- 
 - Removed dependency introduced in build 254 (VS2010)

=========================================================================================       
VSFlex8 Build Number 8.0.20103.271  Build Date: Nov 4, 2010
========================================================================================= 

Corrections 
----------- 
 - Improved F2 detection (TFS 7528)
 - Support editing very long strings (no longer using stack-bound _alloca) (TFS 12386)

=========================================================================================       
VSFlex8 Build Number 8.0.20103.266  Build Date: Sep 22, 2010
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V3/2010 build
 - Improved mouse handling while resizing rows/columns

=========================================================================================       
VSFlex8 Build Number 8.0.20102.264  Build Date: August 24, 2010
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V2/2010 build
 * Built using VS2010 and the latest ATL implementation (reduced ocx size)
 - eliminated mouse move flicker under Win7 / classic screen appearance
 - improved StartEdit behavior on selection change events

=========================================================================================       
VSFlex8 Build Number 8.0.20102.263  Build Date: July 7, 2010
========================================================================================= 

Corrections 
----------- 
- Increased maximum number of subtotal levels from 128 to 256. [8374]

=========================================================================================       
VSFlex8 Build Number 8.0.20101.262  Build Date: April 5, 2010
========================================================================================= 

Corrections 
----------- 
- ADO binding was broken (caused by changes in the latest ATL data classes)
- Fixed problem in Excel shared string table parser (old but hard to reproduce bug)


=========================================================================================       
VSFlex8 Build Number 8.0.20101.261   Build Date: January 11, 2010
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V1/2010 build
 ** Built using VS2008 and the latest ATL implementation (with security adjustments for Vista/Win7)


=========================================================================================       
VSFlex8 Build Number 8.0.20093.258  Build Date: August 11, 2009
========================================================================================= 

* V3/2009 build

Corrections 
----------- 
 - Marked the control as unsafe for initialization, in compliance with ATL Security Update
	http://msdn.microsoft.com/en-us/visualc/ee309358.aspx
   This only affects scenarios where the control is used in Web Pages.
   NOTE: The actual modifications are in the IObjectSafetyImpl declaration and in the 
   use of PROP_ENTRY_TYPE macros.

=========================================================================================       
VSFlex8 Build Number 8.0.20092.257  Build Date: June 12, 2009
========================================================================================= 

Corrections 
----------- 
 - When rendering to metafile in Text AX Container, grid was rendered in black and white.

=========================================================================================       
VSFlex8 Build Number 8.0.20092.256  Build Date: May 19, 2009
========================================================================================= 

* V2/2009 build

=========================================================================================       
VSFlex8 Build Number 8.0.20081.255  Build Date: October 16, 2008
========================================================================================= 

Corrections 
----------- 
 - Fixed memory leak when exporting XLS files (TTP 5120).
 - DAO grid threw exception when binding to sources with more than 170 columns (TTP17110)
 - Handle timer messages while dragging/resizing rows/columns (TTP5116)
	NOTE:
	The grid now allows timer messages while the user moves or resizes rows and columns. 
	This means timer event handlers get called during these actions, but if you change 
	any UI elements while handling these events you need to	refresh the UI elements by 
	calling their refresh method. For example:

	Private Sub Timer1_Timer()
	    Me.Label1 = Time
	    Me.Label1.Refresh ' label won't refresh without this call
	End Sub


=========================================================================================       
VSFlex8 Build Number 8.0.20081.254  Build Date: August 21, 2008
========================================================================================= 

Corrections 
----------- 
 - Property grid in custom property page did not work correctly.


=========================================================================================       
VSFlex8 Build Number 8.0.20081.253  Build Date: July 21, 2008
========================================================================================= 

Corrections 
----------- 
 - Setting Cell(flexcpText) to "vbNullString" caused an exception.


=========================================================================================       
VSFlex8 Build Number 8.0.20081.252  Build Date: June 10, 2008
========================================================================================= 

Corrections 
----------- 
 - Setting ColFormat to "%" didn't work (it did in previous builds)
 - Setting Cell(flexcpFontBold) before the control became visible didn't work (it did in previous builds)


=========================================================================================       
VSFlex8 Build Number 8.0.20081.250  Build Date: May 13, 2008
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
* V1/2008 build
- Rebuilt for "NX" compliance.

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
VSFlex8 Build Number 8.0.20073.242  Build Date: December 18, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
- Improved licensing.


=========================================================================================       
VSFlex8 Build Number 8.0.20073.242  Build Date: November 13, 2007
========================================================================================= 

Corrections 
----------- 
- Restored semantics of GetNodeRow(row, flexNTFirstChild) and GetNodeRow(row, flexNTLastChild)
  This returns the first and last rows (not necessarily node rows).
  This is the original behavior. It was changed in build 226 to address issue AXFLX000126.
  But that was really a non-issue, and the change affected existing code so it was removed.

=========================================================================================       
VSFlex8 Build Number 8.0.20073.241  Build Date: September 17, 2007
========================================================================================= 

Corrections 
----------- 
 - Excel export didn't take into account the settings of ColWidthMin and ColWidthMax properties


=========================================================================================       
VSFlex8 Build Number 8.0.20073.240  Build Date: August 6, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V3/2007 build

Corrections 
----------- 
 - Some font colors were not exported correctly to Excel files (AXFLX000291)
 - Some cells were not clipped correctly when rendering to printer with ExtendLastCol = true (AXFLX000281)
 - Buttons parameter in Ole Drop event handler did not contain the button that was pressed (AXFLX000290)

=========================================================================================       
VSFlex8 Build Number 8.0.20072.239  Build Date: April 19, 2007
========================================================================================= 

Corrections 
----------- 
 ** allow 256 levels in subtotals (previously allowed only 128, AXFLX000293)

=========================================================================================       
VSFlex8 Build Number 8.0.20072.238  Build Date: March 7, 2007
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V2/2007 build

=========================================================================================       
VSFlex8 Build Number 8.0.20071.237  Build Date: December 19, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V1/2007 build

=========================================================================================       
VSFlex8 Build Number 8.0.20063.237  Build Date: December 19, 2006
========================================================================================= 

Corrections 
----------- 
 ** Handle SHRFMLA records when loading xls files (AXFLX000284)

=========================================================================================       
VSFlex8 Build Number 8.0.20063.235  Build Date: August 10, 2006
========================================================================================= 

Corrections 
----------- 
 ** improved Excel output filter to left-align cells with ADO string types
	adChar (129) adWChar (130) adVarChar (200) adLongVarChar (201) adVarWChar (202) adLongVarWChar (203)
 ** improved Excel filter to handle localized built-in date formats
 ** improved DAO data-binding to avoid cursor positioning error after refresh (AXFLX000273)

=========================================================================================       
VSFlex8 Build Number 8.0.20063.234  Build Date: June 16, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V3/2006 build

Corrections 
----------- 
 ** Honor column alignment in editor when starting with empty string (AXFLX000242)

=========================================================================================       
VSFlex8 Build Number 8.0.20062.233  Build Date: May 10, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** Fixed Excel export of Dates to certain international locales (e.g. German)
 ** Fixed memory leak associated with ColKey property (AXFLX000260)
 ** Fixed scrolling problem associated with ColPosition property (AXFLX000261)


=========================================================================================       
VSFlex8 Build Number 8.0.20062.232  Build Date: March 2, 2006
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V2/2006 build
 ** Improved rendering of right-aligned mergeable cells

Corrections 
----------- 
 ** Fixed thread-safety problem in string formatting routines.
 ** Honor SheetBorder color in RenderControl (border doesn't have to be black)


=========================================================================================       
VSFlex8 Build Number 8.0.20061.231  Build Date: Dec 20, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** Honor formatting with no leading zeros (e.g. ".000")


=========================================================================================       
VSFlex8 Build Number 8.0.20061.230  Build Date: Nov 28, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** Fixed bug in Excel SST export feature (added in build 226).
 ** Load TrackMouseEvent proc dynamically to allow registration under Win95 (AXFLX000206)


=========================================================================================       
VSFlex8 Build Number 8.0.20053.228   Build Date: Oct 31, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V1/2006 build

Corrections 
----------- 
 ** CSV export used \r instead of \r\n in some cases (AXFLX000205)
 ** Cell(cpValue) did not return correct values for string starting with decimal point under Win98 (AXFLX000180)
 ** Added extra checks to prevent fatal error under Win98 when closing app while handing event (AXFLX000222)

=========================================================================================       
VSFlex8 Build Number 8.0.20053.226   Build Date: June 17, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V3/2005 build
 ** Export to Excel now supports the creation of shared string tables, which can save cells
    with up to about 8k of text (the previous limit was 256 characters when saving).
    SST also reduces the size of the xls files.
 ** Allow resizing of columns that are too wide to fit the control

Corrections 
----------- 
 ** fixed selection when removing all rows (should change Row property to -1, was broken in build 221)
 ** fixed problem with international formatting (broken in build 216 with precision increase)
 ** fixed db-cursor to handle RightToLeft locales (when GridLinesFixed = flexGridDataGrid)
 ** fixed some international date formatting in Excel export (AXFLX000116)
 ** GetNode method returned invalid (non-node) row when called on non-node row with
    flexNTFirstChild/flexNTFirstChild parameters (AXFLX000126)


=========================================================================================       
VSFlex8 Build Number 8.0.20052.223   Build Date: June 9, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** fixed memory leak in sort method (with certain data types, when grid alrady sorted)
 ** fixed auto clipboard functions in Unicode version
 ** fixed mixed-mode (tristate) display with XP themes
 ** fixed problem with checkbox logic when grid has only fixed rows
 
=========================================================================================       
VSFlex8 Build Number 8.0.20052.222   Build Date: April 26, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** allow MergeSpill to work on fixed columns.

=========================================================================================       
VSFlex8 Build Number 8.0.20052.221   Build Date: March 11, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** allow loading Excel files while they are open in Excel (AXFLX000114)

Corrections 
----------- 
 ** fixed mouse selection problem with merged cells and invisible rows/cols.
 ** fixed problem with PrintGrid with RightToLeft and ExtendLastCol (AXFLX000142)
 ** fixed inconsistent selection behavior when deleting last row in listobx grids (AXFLX000149)
 ** fixed autosize rows with content ending with \r\n (AXFLX000132)
 ** honor "short time" format in excel export (AXFLX000110)
 ** support international currencies in Excel export (AXFLX000109, AXFLX000059)
 ** improved excel export (was locking some styles under Excel 2000) (AXFLX000106)
 ** fixed ",." format (was broken in build) (AXFLX000060)
 ** improved handling of Japanese currency symbol in Excel export (AXFLX000109)
 ** fixed painting of combo lists with more than 37000 items (AXFLX000160)


=========================================================================================       
VSFlex8 Build Number 8.0.20052.222   Build Date: February 20, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** V2/2005 build
 ** Allow resizing rows/columns past the edges of the control (like Explorer etc)
 ** Added new "flexCFBindToBinaryFields" setting to ControlFlagsSettings enumeration.

    By default, the VSFlexGrid will not show binary fields when it is bound to a data source.
    This is by design, because the grid doesn't know how to display or edit these fields 
    (they could be images, for example).

    However, the Microsoft DataGrid does show binary fields, and some users have requested
    compatibility on the VSFlexGrid. This new flag provides that option. For example:

    ' ** show bound binary fields (e.g images, arrays, etc)
    Me.VSFlexGrid1.Flags = Me.VSFlexGrid1.Flags Or flexCFBindToBinaryFields
    
    ' ** hide bound binary fields (default behavior)
    Me.VSFlexGrid1.Flags = Me.VSFlexGrid1.Flags And (Not flexCFBindToBinaryFields)

    Note that the grid still won't show binary fields by default. You need to set the
    flexCFBindToBinaryFields flag in order to get that behavior.

Corrections 
----------- 
 ** none


=========================================================================================       
VSFlex8 Build Number 8.0.20051.219   Build Date: January 20, 2005
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** improved accessibility object (compliance with AccExplorer 2.0)

Corrections 
----------- 
 ** fixed international and percentage formatting (broken in build 216)


=========================================================================================       
VSFlex8 Build Number 8.0.20051.217   Build Date: December 21, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** fixed hot-tracking to take merged cells into account


=========================================================================================       
VSFlex8 Build Number 8.0.20051.216   Build Date: December 1, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** Q1/2005 build
 ** Improved formatting of currency values (more precision)
    the new code can handle values such as 922337203685477.5732 without round-off errors
 ** Added hot-tracking to mouse handling: when using XpThemes, the control will highlight
    fixed cells using the tracking theme.

Corrections 
----------- 
 ** Improved Excel export filter:
    - earlier versions wouldn't import into MS Access
    - had problems saving invisible rows
    - improved font size calculation to reduce round-off errors
 ** Fixed painting of combo buttons when flexSBEditing (could in some cases show button when it shouldn't)


=========================================================================================       
VSFlex8 Build Number 8.0.20044.215   Build Date: October 13, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** right-to-left didn't display ellipsis properly


=========================================================================================       
VSFlex8 Build Number 8.0.20044.214   Build Date: October 7, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** cell images were not displayed in fixed cells when using XP themes 


=========================================================================================       
VSFlex8 Build Number 8.0.20044.213   Build Date: September 19, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** added check to prevent C++ UserControls crash under WinXP 2/Unicode
 ** Node.Move didn't always fire the selection events properly
 ** ValueMatrix("x123") returned 123 rather than zero
 ** improved mouse handling when freezing rows/cols on multiple grids
 ** improved behavior of ESCAPE key when cancelling freezes


=========================================================================================       
VSFlex8 Build Number 8.0.20044.211   Build Date: August 31, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** updated licensing code


=========================================================================================       
VSFlex8 Build Number 8.0.20044.210   Build Date: July 23, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** Added support for loading/saving UTF-16 encoded text files (unicode)
       This support is provided only on the unicode versions of the control
       (VSFlex8u and VSFlex8n).

       - Loading unicode text files (e.g. saved by Excel) is automatic. The
         control will detect the unicode byte-order indicator 0xFFFE at the
         start of the file and will automatically switch to unicode mode. For example:

            ' load tab-delimited text file (regular or Unicode)
            fg.LoadGrid fileName, flexFileTabText

       - By default, the grid will save text files with regular encoding (unicode
         characters are converted into '?' (this is the same behavior as Excel). To
         save unicode text files, you have to specify the "u" flag in the options
         parameter when calling the SaveGrid method. For example:

            ' save tab-delimited text file (regular text)
            ' "vf" parameter means: Visible only, load Fixed cells
            fg.SaveGrid fileName, flexFileTabText, "vf" 

            ' save tab-delimited text file (unicode)
            ' "uvf" parameter means: Unicode, Visible only, load Fixed cells
            fg.SaveGrid fileName, flexFileTabText, "uvf" 

Corrections 
----------- 
 ** none


=========================================================================================       
VSFlex8 Build Number 8.0.20043.207   Build Date: July 9, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** fixed ATL bug that caused some flicker when grid got the focus (in C++ projects only)


=========================================================================================       
VSFlex8 Build Number 8.0.20043.206   Build Date: June 3, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** fixed LoadGrid(binary), was broken when dealing with merged cells (in build 205)


=========================================================================================       
VSFlex8 Build Number 8.0.20043.205   Build Date: April 24, 2004              ** BAD BUILD
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** Q3/2004 drop

Corrections 
----------- 
 ** improved combobox handling in multi-monitor systems
 ** handle byref strings in PrintGrid
 ** fix OldCol parameter in SelChange event when setting data source
 ** Clear(all) now clears checkboxes
 ** improved handling of special chars in LoadGrid(textfile)
 ** improved saving very large Unicode grids with SaveGrid()
  

=========================================================================================       
VSFlex8 Build Number 8.0.20042.205   Build Date: April 24, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** Read xls files with far-east and rich text information.


=========================================================================================       
VSFlex8 Build Number 8.0.20042.204   Build Date: April 2, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** none

Corrections 
----------- 
 ** Allow resetting the ColImageList property by setting it to zero.


=========================================================================================       
VSFlex8 Build Number 8.0.20042.203   Build Date: February 16, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** Q2/2004 build

Corrections 
----------- 
 ** none


=========================================================================================       
VSFlex8 Build Number 8.0.20041.203   Build Date: January 16, 2004
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** None

Corrections 
----------- 
 ** Improved Excel filter to account for 64k row/255 col limits on Excel sheets


=========================================================================================       
VSFlex8 Build Number 8.0.20041.202   Build Date: December 8, 2003 
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** None

Corrections 
----------- 
 ** Node.Sort method didn't work when SubtotalPosition = Below.
 ** Fixed memory leak that affected chaptered rowsets
        (showed when repeatedly changing the recordset's Filter property)
 ** SaveGrid (data) shouldn't save checkbox settings for boolean columns


=========================================================================================       
VSFlex8 Build Number 8.0.20041.201   Build Date: November 26, 2003 
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 ** None

Corrections 
----------- 
 ** Excel export: cells with back color had their styles 'locked' in Excel. This was fixed.


=========================================================================================       
VSFlex8 Build Number 8.0.20041.200   Build Date: November 11, 2003 
========================================================================================= 

Enhancements/Documentation/Behavior Changes 
------------------------------------------- 
 * Q1/2004 drop

 * Improved support for XP themes
        If the application is theme-enabled, the grid will now use themes to display
        checkboxes and collapse/expand buttons in addition to drop-down buttons and 
        scrollbars.
        There's a new setting for the Appearance property, flexXPThemes.
        If you set Appearance to flexXPThemes and the application is theme-enabled, 
        then the control will paint fixed cells using themes as well.
 
 * New property: MergeCellsFixed As MergeSettings 
		Allows users to set different merging criteria for fixed vs scrollable cells
		Setting MergeCells automatically sets MergeCellsFixed to the same value
		(for compatibility)
		If the MergeCells and MergeCellFixed settings are different, files saved
		with SaveGrid(All) method will not be read by older versions of the control

 * New property: GroupCompare As MergeCompareSettings
		Returns or sets the type of comparison used when grouping cells.
		By default, the Subtotal method will group cells when there is an exact match 
		between adjacent cells. This property allows you to control the comparison
		parameters (case-insensitive and trimming, like MergeCompare).

 * New method: CellBorderRange
		Similar to CellBorder, but allows user to specify the range instead of using the selection
        CellBorderRange
                Row1 As Long, Col1 As Long, Row2 As Long, Col2 As Long,
		        Color As Long, Left As Integer, Top As Integer, Right As Integer, Bottom As Integer, 
                Vertical As Integer, Horizontal As Integer
        Row1, Col1, Row2, Col2: range where the border will be applied
        Color: color for the border
        Left, Top, Right, Bottom: width of the outside border, in pixels
        Vertical, Horizontal: width of the inside border, in pixels

 * New methods for clipboard support: 
		Cut(), Copy(), Paste(), Delete()

 * New method for finding rows using regular expressions:
        FindRowRegex(Pattern As String, Row As Long, Col As Long) As Long
        Pattern containing the regular expression to look for
            (see the Pattern property in the VBScript Regex object for regular expression syntax)
        Row row where the search should start (use -1 to start at the first scrollable row)
        Col column to search
        Returns the index of the row that contains a match or -1 if no match was found

 * New properties for custom sort images:
        SortAscendingPicture and SortDescendingPicture

 * New settings for Flags property (runtime only)
		flexCFAutoClipboard: causes the control to handle clipboard keys automatically
			Copy:	Ctrl+C, Ctrl+Ins
			Cut:	Ctrl+X, Shift+Del	(if editable)
			Paste:	Ctrl+V, Shift+Ins	(if editable)
			Delete:	Del					(if editable)
		flexCFNoEditIndent: 
			causes the edit control to use the old behavior: align to the left, no indent

 * More control when saving text files
        The SaveGrid/LoadGrid methods have a variant parameter that is used
        to specify the sheet name when exporting to Excel (string), or whether to in 
        save/load fixed cells as well as scrollable when loading/saveing txt files
        (bool). Starting in build 200, the control also recognizes a string parameter 
        when saving/loading text files. If the string contains an 'f', fixed cells will
        be included when saving/loading. If the string contains a 'v', only visible 
        cells will be saved to the text files.
        For example:
            fg.Save("c:\export\)

 * Improved text editor behavior: honor ColIndent, Alignment
		If you want to prevent that and keep the old behavior,
		set fg.Flags = fg.Flags | flexCFNoEditIndent

 * Version property now returns 801
        If you want your app to run correctly with older builds of VSFlex8, check the Version
        property before using these new features. For example:
            if fg.Version >= 801 then fg.Copy

Corrections 
----------- 
 ** None



** 8.0.20033.193 - August 26
    Improved saving multi-line entries to Excel (cr/lf -> lf), need to set WordWrap = true
    Fixed rendering text in outline button bar when tree style is complete leaf
    Fixed multi-col combos (broken in 191)
    Improved precision of numeric formatting
    Added licensing support for non-standard containers

** 8.0.20033.192 - July 24
    Fixed problem with combo boxes in .NET 1.1
    Improved word wrap logic in Excel export

** 8.0.20033.191 - June 26
    Fixed problem in Excel output of custom date formats

** 8.0.20033.190 - June 10
    Fixed licensing expiration problem

** 8.0.20033.189 - May 27
	Q3 drop

** 8.0.20032.189 - May 24
	Fixed formatting problem with parenthesis and percentages "(###)"
	Improved logic for deleting rows when databound (keep RowData, cell styles)

** 8.0.20032.188 - May 11
    Support hidden rows/columns when saving/loading Excel files
	Improved support for DBCS fonts in Excel export
	Fixed bug that caused TopRow to be set to a value > 1 when instantiating in IE and FrozenCols > 0

** 8.0.20032.187 - April 25
    Several improvements to Excel filter: <<DOC>>
    - support more than 30k rows
    - support word-wrapping
    - support ColorAlternate, ColorFrozen
    - added options for saving fixed rows/columns/translated combo values:
        fg.SaveGrid "book1.xls"
        fg.SaveGrid "book1.xls", "sheetName"
        fg.SaveGrid "book1.xls", flexXLSaveFixedCells
        fg.SaveGrid "book1.xls", flexXLSaveFixedRows
        fg.SaveGrid "book1.xls", flexXLSaveFixedCols
        fg.SaveGrid "book1.xls", flexXLSaveRaw
        fg.SaveGrid "book1.xls", flexXLSaveFixedCells Or flexXLSaveRaw
        (the options are flags in the SaveExcelSettings enumeration)

    Improved handling of null/empty strings when updating data source
    Fixed problem with EditText value when editing checkboxes with the mouse

** 8.0.20032.186 - April 11
    Support Unicode chars in Excel formats (needed for Euro)
    Fixed formatting problem introduced in 184

** 8.0.20032.185 - April 7 *** BAD BUILD ***
    Honor number of lines set for mouse wheel

** 8.0.20032.184 - April 3 *** BAD BUILD ***
	Fixed bug in formatting of international numbers (introduced in 183)

** 8.0.20032.183 - March 26 *** BAD BUILD ***
    Improved DBCS handling in Excel export
	Improved Excel filter to load results of formulas that return strings
	Improved formatting of international currencies

** 8.0.20032.181 - March 20
	Improved handling of Font.* properties (still need to call Refresh after changing)
    Fixed clip property bug introduced in 180
	Improved support for DBCS characters in Excel import/export
	Added support for Windows XP themes
		The support is automatic. If the grid is used in an application 
		that has the appropriate manifest, and is running under Windows XP,
		then the Combo and Edit buttons and editors will automatically 
		be displayed using the current theme.

** 8.0.20031.180 - Feb 25
    Fixed problem with closing IME editor after input
    Made TopRow/LeftCol work even when Redraw = false (like C1FlexGrid)
    MouseUp fired only once on DblClick (should fire twice like other controls)
    Added new setting for PictureType = flexPictureEnhMetafile
    RUN_GEN_002_014 Spelling errors in help string for WallPaperalignment property
    RUN_GEN_001_002 Improved behavior of MergeCompare = IncludeNulls on empty grids
    RUN_GEN_001_001 Fixed saving DBCS characters to Excel in non-Unicode builds
    RUN_GEN_002_008 Fixed AutoSize rows to account for cells that end in DBCS characters
    RUN_GEN_002_006 Fixed LoadGrid for text files that contain DBCS characters
    RUN_GEN_002_005 Fixed width of drop list with empty columns
    RUN_GEN_002_003 Improved handling of null values in numeric sorts
    RUN_GEN_002_001 Fixed behavior of Clip property when string contains empty rows

** 8.0.20031.179 - Feb 6
	Added new settings for text alignment: GeneralTop, GeneralBottom, GeneralCenter

** 8.0.20031.178 - Jan 31
	Minor fix in Excel export when no font == null (no site)
	Minor fix in selection events when removing all rows

** 8.0.20031.177 - Jan 14
    Made minor improvements to Excel filter (custom format handling)

** 8.0.20031.176 - Jan 8
    Added Flags property (runtime only) to improve compatibility with version 7

    Currently, there's only two flags defined: flexCFNone and flexCFV7SelectionEvents
    
    By default, version 8 fires selection change events whenever the selection
    changes, even when the cell coordinates are invalid (e.g. the selection was
    removed by setting Row = -1 or the old row was removed with Rows=1).

    If you set fg8.Flags = flexCFV7SelectionEvents, then the control will fire
    selection events (BeforeSelChange, AfterSelChange) only when the Row/Col 
    parameters refer to valid coordinates, the same behavior as in version 7.


** 8.0.20031.175 - Jan 5
    Fixed crashing problem with keyboard navigation on merged grids with no rows


////////////////////////////////////////////////////////////////////////////////
//
// 2002
//
////////////////////////////////////////////////////////////////////////////////

** 8.0.20031.174 - Dec 22
    Fixed accessibility to register under Win95/NT
    Improved mouse handling to work with posted messages

** 8.0.20031.173 - Dec 20
    xxxx    Small change in handling MousePointer
    7474    VSFlexString: fixed memory leak

** 8.0.20031.172 - Dec 16
    FindRow could throw an exception instead of returning -1
    Extra error-checking in Excel save code
    Added new setting to MergeCompare property: flexMCIncludeNulls merges empty cells

** 8.0.20031.171 - Dec 6
    Fixed Unicode support in Excel load/save

** 8.0.20031.170 - Dec 2
    Synchronized minor build number with VSFlex7 for easier tracking
    
    ** Always fire BeforeSelChange/AfterSelChange, even when ranges contain negative row/col parameters
       This wasn't done before, initially because of a bug and leter because the fix could break existing
       applications. Flex8 always fires these events, so make sure you don't use the parameters without
       testing them first. For example:

       Private Sub fg_AfterSelChange(newrow, ... oldrow)

            If newRow > -1 Then ' newRow = -1 means 'no selection'
                
                If fg.IsVisible(newRow) Then ...

            End If

       End Sub

** 8.0.20031.11 - Nov 25
    7355    Fix to 6531 in previous version introduced a bad data-binding bug (could hang when sorting datasets)
    7354    Setting Cols within the AfterRowColChange event could fire AfterRowColChange again
    7356    BeforeSelChange event firing twice when using Shift key to select multiple rows in listbox mode
    xxxx    BeforeSelChange _not_ firing when Row/Col didn't change (only RowSel/ColSel)

** 8.0.20031.10 - Nov 15 *** bad build
    6531    Fixed small memory leak when binding to empty recordsets
    6915    Setting GridLinesFixed property to DataGrid only showed cursor in VSFlexGrid Light

*** 8.0.20031.9 - Nov 7
    2003/Q1 build
    Improved formatting behavior to increase compatibility with VB (suppress leading zeros)
        0.### > 0.500   .### > .500


////////////////////////////////////////////////////////////////////////////////////////////////////
//
// What's new in VSFlex 8.0
//
////////////////////////////////////////////////////////////////////////////////////////////////////


Subscription licensing scheme, new About box, all incremental updates and fixes 
applied to previous versions. If you haven't been dowloading the latest patches
from our web site periodically, this is an easy way to get everything in one step.

Excel import/export (using SaveGrid/LoadGrid methods).


All versions of VSFlex8 (ADO, DAO, Light, etc)
----------------------------------------------


** New properties:

string  AccessibleName           Gets or sets the name of the control used by accessibility client applications.
string  AccessibleDescription    Gets or sets the description of the control used by accessibility client applications.
string  AccessibleValue          Gets or sets the value of the control used by accessibility client applications.
Variant AccessibleRole           Gets or sets the role of the control used by accessibility client applications.

These new properties support Microsoft's Active Accessibility effort. Use them to make your 
applications more friendly to people with physical impairments, and to comply with US 
regulations.

boolean IsSearching             Returns a value that indicates whther the ghrid is in auto search mode
StartAutoSearch                 Fired when the grid enters auto search mode
EndAutoSearch                   Fired when the grid exits auto search mode



** New settings:

The SaveGrid and LoadGrid methods have a new setting, flexFileExcel, that allows you to 
save and load XLS (Excel 97) files. For example, the command

	fg.SaveGrid fileName, flexFileExcel

would save the grid contents into an Excel97 file. The command

	fg.LoadGrid fileName, flexFileExcel

would load the first sheet of an Excel97 workbook into the grid.

Notes about the Excel filter:

1) It does not require Excel to be present on the machine.
2) It only supports the Excel97 format (Excel95 and earlier are not supported).
3) It translates cell values (including formula values), fonts, colors, row and column dimensions.
4) It does not translate macros, charts, cell borders, rotated text, and other advanced formatting.
5) Formulas are translated into values.
6) When saving, only a single sheet is created.
7) When loading, only a single sheet is imported (you can specify which one).


** Documentation Update (apply to VSFlex7 hlp contents)

** SaveGrid

Saves grid contents and format to a file.

Syntax     
[form!]VSFlexGrid.SaveGrid FileName As String, SaveWhat As SaveLoadSettings, [ FixedCells As Boolean ]

Remarks


This method saves a grid to a binary or to a text file. The grid may be retrieved later using the 
LoadGrid method. Grids saved to text files may also be read by other programs, such as Microsoft 
Excel or Microsoft Word. 

The parameters for the SaveGrid method are described below:


- 
FileName As String

The name of the file to create, including the path. If a file with the same name already exists, it
is overwritten.


- 
SaveWhat As SaveLoadSettings

This parameter specifies what should be saved. Valid options are:


Constant	Value	Description

flexFileAll	0	Save all data and formatting information to a proprietary binary format.

flexFileData	1	Save only the data, ignoring formatting information to a proprietary binary format.flexFileFormat	2	Save only the global formatting, ignoring the data to a proprietary binary format.

flexFileCommaText	3	Save data to a comma-delimited text file.

flexFileTabText	4	Save data to a tab-delimited text file.

flexFileCustomText	5	Save data to a text file using the delimiters specified by the ClipSeparators property.
flexFileExcel	6	Save all data and formatting information to an Excel97 file.


- Options As Variant  (optional)

When saving and loading text files, this parameter allows you to specify whether fixed cells are saved
and restored. The default is False, which means fixed cells are not saved or restored.

When saving and loading Excel files, this parameter allows you to specify the name or index of the sheet
to be loaded, or the name of the sheet to be saved. If omitted, the first sheet is loaded.

- Notes:
The flexFileFormat option saves global formatting only. It does not save any cell-specific information, 
not even the number of rows and columns. This allows you to use this setting to create formats that can 
be applied to existing grids even if they have different dimensions. 

Because column widths and row heights are related to the number of rows and columns on the grid, they 
are not saved or restored if you use the flexFileFormat option.The following is a list of the properties
that do get saved and restored if you use the flexFileFormat option: 
BackColor, ForeColor, BackColorBkg, BackColorAlternate, BackColorFixed, ForeColorFixed, BackColorSel,
ForeColorSel, TreeColor, SheetBorder, GridLines, GridLinesFixed, GridLineWidth, GridColor, GridColorFixed,
TextStyle, TextStyleFixed, ScrollBars, SelectionMode, RowHeightMin, MergeCells, SubtotalPosition, OutlineBar,
Font, and WordWrap.



If your application requires you to save several grids, you should consider using the Archive method to 
compress and combine them all into a single archive file. You can later use the ArchiveInfo method to 
retrieve information from the archive file.

The flexFileExcel option is new in Version 8. It does not require Excel to be present on the machine. You
can load and save Excel97 sheets (BIFF9 format), one sheet per workbook only (when loadinf, you can specify 
which sheet using the Options parameter).

The Excel filter supports cell values (including formula values), fonts, formats, colors, row and column
dimensions. It does not support features that don't translate into the grid, such as macros, charts, 
rotated text, cell borders, and other advanced formatting.


** LoadGrid

Loads grid contents and format from a file.

Syntax     
[form!]VSFlexGrid.LoadGrid FileName As String, LoadWhat As SaveLoadSettings, [ Options As Variant ]

Remarks


This method loads grid from a file previously saved with the SaveGrid method, comma-delimited text file
(CSV format) such as an Excel text file, or a tab-delimited text file.

 The parameters for the LoadGrid method
are described below:

- FileName As String


The name of the file to load, including the path. This file must have been created by the SaveGrid method,
or an Invalid File Format error will occur (error #321).


- 
LoadWhat As SaveLoadSettings


This parameter specifies what should be loaded. Valid options are:


Constant	Value	Description
flexFileAll	0	Load all data and formatting information available in the file.

flexFileData	1	Load only the data, ignoring formatting information.

flexFileFormat	2	Load only the formatting, ignoring the data.

flexFileCommaText	3	Load data from a comma-delimited text file.

flexFileTabText	4	Load data from a tab-delimited text file.
flexFileCustomText	5	Load data from a text file using the delimiters specified by the ClipSeparators property.
flexFileExcel	6	Load a sheet from an Excel97 file (you can specify which sheet to load using the Options parameter).

- Options As Variant (optional)

When saving and loading text files, this parameter allows you to specify whether fixed cells are saved
and restored. The default is False, which means fixed cells are not saved or restored.

When saving and loading Excel files, this parameter allows you to specify the name or index of the sheet
to be loaded, or the name of the sheet to be saved. If omitted, the first sheet is loaded.

- Notes:
The flexFileExcel option is new in Version 8. It does not require Excel to be present on the machine. You
can load and save Excel97 sheets (BIFF9 format), one sheet per workbook only (when loadinf, you can specify 
which sheet using the Options parameter).

The Excel filter supports cell values (including formula values), fonts, formats, colors, row and column
dimensions. It does not support features that don't translate into the grid, such as macros, charts, 
rotated text, cell borders, and other advanced formatting.
