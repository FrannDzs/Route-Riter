============================================================================
TrueDBGrid Pro Build Number 8.0.20173.378
============================================================================
 * Updated AboutBox info.

============================================================================
TrueDBGrid Pro Build Number 8.0.20173.377
============================================================================
 * Updated AboutBox info.

============================================================================
TrueDBGrid Pro Build Number 8.0.20173.376
============================================================================
 * Updated AboutBox info.

============================================================================
TrueDBGrid Pro Build Number 8.0.20173.375
============================================================================
 * Updated AboutBox info.

============================================================================
TrueDBGrid Pro Build Number 8.0.20163.374
============================================================================
 * Updated AboutBox for licensing purposes.

============================================================================
TrueDBGrid Pro Build Number 8.0.20133.373
============================================================================

Corrections
-----------

 * This build corrects a problem with mouse clicks in the non-cell area causing
   an exception.  (Tfs-142472).


============================================================================
TrueDBGrid Pro Build Number 8.0.20133.372
============================================================================

Corrections
-----------

 * This build corrects a problem with InactiveBackColor and InactiveForeColor
   invocations upon loss of focus.
   (Tfs-56139).


============================================================================
TrueDBGrid Pro Build Number 8.0.20133.371
============================================================================

Corrections
-----------

 * This build resolves an exception issue created in the previous build
   associated with the RightToLeft property in the PropertyPages.  The
   exception occurred due to an unexpected NULL value in the new code.
   (Tfs-50252).

============================================================================
TrueDBGrid Pro Build Number 8.0.20133.370
============================================================================

Corrections
-----------

 * This build correctly restricts the setting of the RightToLeft property
   based on the value of the GetSystemMetrics(SM_MIDEASTENABLED) result.

============================================================================
TrueDBGrid Pro Build Number 8.0.20133.369
============================================================================

Corrections
-----------

 * This build eliminates an exception when checking checkbox column
   values in the BeforeColUpdate and AfterColUpdate events.

============================================================================
TrueDBGrid Pro Build Number 8.0.20132.368
============================================================================

Corrections
-----------

 * This build restores certain grid behaviors when cell edits force row order
   changes in very specific circumstances.  It is possible that side effects
   may be observed for server side data cursors.  (5801 - Revisited)

============================================================================
TrueDBGrid Pro Build Number 8.0.20132.367
============================================================================

Corrections
-----------

 * Build 366 limited scrollbar width to a single value for themes which do
   not support scrollbar width changes. This restriction has been relaxed to
   allow scrollbars with smaller widths.  Using a smaller width suppresses
   OS painting of the scrollbar for theme behaviors. (40455)

============================================================================
TrueDBGrid Pro Build Number 8.0.20132.366
============================================================================

Corrections
-----------

 * As a result of Windows Vista/7/8 themes and Visual Styles, changing the
   scrollbar width could result what appeared to be multiple scrollbars, but
   was actually the Windows imposed theme behavior on system scrollbars. For
   themes which do not support scrollbar width changes (Aero, Basic, etc.)
   the grid now ignores scroll width changes. (40455)

 * Column ButtonPicture and FilterButtonPicture can be cleared by setting
   the properties to NULL or Nothing. (41601)

============================================================================
TrueDBGrid Pro Build Number 8.0.20121.365
============================================================================

Corrections
-----------

 * Using the Columns.Value property in the BeforeColUpdate wasn't returning
   the correct value when using TDBDropdown with value translation and having
   duplicate DisplayValues. (19197)

 * CellTips won't display on secondary monitor in a Dual Monitor scenario. (6140)

 * Fixed incorrect printing with merged cells. (16659)

============================================================================
TrueDBGrid Pro Build Number 8.0.20093.364
============================================================================

Corrections
-----------

 * Unable to set focus to the last column in an Inverted grid that had
   fixed columns.  (4569)

 * Extra row was being inserted into the top of the grid when modifying
   the cell value of a sorted grid and the cell being modifed was the column
   in which the data source was sorted on. (5801)

============================================================================
TrueDBGrid Pro Build Number 8.0.20081.362
============================================================================

Corrections
-----------

 * When more than one grid was placed on a dialog in VS2005 the licensing
   screen was displayed at design time when the grid got focus.  (AXTDG000260)

 * Some printer setting (e.g., duplex, 2 up pagination) were not being
   honored when the grid was printed.  (AXTDG000259)

============================================================================
TrueDBGrid Pro Build Number 8.0.20073.361
============================================================================

Corrections
-----------

 * Fixed exception that was thrown when the grid was hosted in MS Project.

============================================================================
TrueDBGrid Pro Build Number 8.0.20073.360
============================================================================

Corrections
-----------

 * Cancelling an update in the BeforeColUpdate() event on a cell with 
   ValueItems.Presentation = dbgCheckBox was dirtying the cell.  (AXTDBG000250)

 * The Find() method of XArrayDB with XTYPE_STRINGCASESENSITIVE was performing
   a case insensitive search.  (AXTDG000248)

 * Corrected interop issue with the ICursor grid and IE7.  (AXTDG000241)


============================================================================
TrueDBGrid Pro Build Number 8.0.20071.358
============================================================================

Corrections
-----------

 * Clearing a cell in a column attached to a TrueDBDropdown control with translated values was
   populating the cell with the previous selection made in that column.  (AXTDG000243)

============================================================================
TrueDBGrid Pro Build Number 8.0.20063.357
============================================================================

Corrections
-----------

 * Corrected "Data type mismatch" error/exception when updating a column with 
   translated values. (AXTDG000238,AXTDG000239)

 * Changing the underlying datasource on a bound grid when the current cell 
   was bound to a memo field was not clearing the data cache. (AXTDG000240)

============================================================================
TrueDBGrid Pro Build Number 8.0.20063.354
============================================================================

Corrections
-----------

 * The PrintDialog was not correctly reflecting the PrintInfo.NumberOfCopies property when
   printing from PrintPreview.  (AXTDG000232) 

 * Expanding/collapsing a child grid wasn't working under Inverted or Form dataview.  (AXTDG000229)


============================================================================
TrueDBGrid Pro Build Number 8.0.20063.353
============================================================================

Enhancements
------------

 * The grid now supports a TrueDBDropdown using ValueTranslate when the listfield
   column contains duplicate values.  (AXTDG000205)

Corrections
-----------

 * Fixed exception with Hierarchical Dataview when the layout was defined
   at design time and the grid was bound to the data source at run time.  (AXTDG000227)

 * The column header divider lines were not correctly positioned when using
   xp themes.  (AXTDG000223)


============================================================================
TrueDBGrid Pro Build Number 8.0.20062.352
============================================================================

Corrections
-----------

 * Fixed horizontal scrollbar when used with a split configued with dbgNumberOfColumns. (AXTDG000215)

 * Corrected rendering of the RecordSelector on the last row of the grid.  (AXTDG000214)

============================================================================
TrueDBGrid Pro Build Number 8.0.20061.351
============================================================================

Corrections
-----------

 * Columns headers were not rendering correctly when selected in 256 color mode.  (AXTDG000209)

 * Clicking on the rowheader of the Fitlerbar or ColumnHeader rows was causing
   extra calls to UnBoundReadData() events.  (AXTDG000211)


============================================================================
TrueDBGrid Pro Build Number 8.0.20053.347   Build Date: Friday, July 1, 2005
============================================================================

Corrections
-----------

 * Fixed printing/previewing a grid that contained restricted merged columns. (AXTDG000195)

 * Specifying a literal in an EditMask using the '\' escape character was 
   incorrectly applying regional settings to the literal.  (AXTDG000199)
   

============================================================================
TrueDBGrid Pro Build Number 8.0.20052.345   Build Date: Friday, May 20, 2005
============================================================================

Corrections
-----------

 * Removing a row from a clone of a recordset that was bound to the grid was
   clearing the SelBookmarks collection.  It now removes the deleted row
   from the collection. (AXTDG000097)

 * The Selbookmarks collection was reporting incorrect bookmarks in the 
   AfterDelete() event when a row was deleted.  It now removes the deleted 
   row from the collection. (AXTDG000121)

 * Printing/Previewing a grid in GroupBy mode when focus was on the 2nd
   split was incorrectly rendering the grouped cells.  (AXTDG000192)

==============================================================================
TrueDBGrid Pro Build Number 8.0.20052.344   Build Date: Friday, March 11, 2005
==============================================================================

Enhancements
------------

 * Optimized the grids rendering with columns that were using dbgMergeRestricted.

Corrections
-----------

 * Fixed exception when all columns had their AllowFocus property set to false. (AXTDG000158)

 * Split caption was incorrectly rendered when the caption contained non-alphabetical characters.  (AXTDG000170)

 * Alignment property of checkbox columns was incorrect when the grid was themed.  (AXTDG000171)

 * Inputting a key when the cell had the DropdownList property set to true and
   the key did not match an entry in the combo was clearing the contents of the cell (AXTDG000172)

 * The Update() method was displaying multiple validation errors.  (AXTDG000178)

 * A hierarchical grid could not display more than 4096 rows.  (AXTDG000185)

====================================================================================
TrueDBGrid Pro Build Number 8.0.20052.342   Build Date: Wednesday, February 16, 2005
====================================================================================

Corrections
-----------

 * Animated CellTips weren't being displayed under Win2k when using Blend.  (AXTDG000092)

 * Column header alignment was incorrect for a column containing the expand/collapse
   icon for a ChildGrid.  (AXTDG000176)

=================================================================================
TrueDBGrid Pro Build Number 8.0.20052.341   Build Date: Friday, February 11, 2005
=================================================================================

Corrections
-----------

 * Addnew row wasn't rendering correctly for MultiLine mode.  (AXTDG000167)

 * Date Edit mask wasn't working properly under Japanese Locale.  (AXTDG000173)

 * Fixed printing/previewing for an unbound column when the native data type
   was something other than string.  (AXTDG000174)

===================================================================================
TrueDBGrid Pro Build Number 8.0.20052.340   Build Date: Tuesday, February 8, 2005
===================================================================================

Corrections
-----------

 * Fixed focus problem when the WrapRowPointer was set to true.  (AXTDG000153)

 * FetchCellStyle property wasn't retaining the value when set at design time.  (AXTDG000154)

 * ColContaining method was returning -1 for grouped column.  (AXTDG000156)

 * Horizontal Scrollbar was not working correctly when ExtendRightColumn was set.  (AXTDG000157)

 * Bookmark property was returning NULL in the AfterInsert event for a newly added row.  (AXTDG000160)


===================================================================================
TrueDBGrid Pro Build Number 8.0.20052.339   Build Date: Wednesday, February 2, 2005
===================================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------

 * Child grids are now supported for Inverted and Form dataviews.

Corrections
-----------

 * DBCS strings for Radio buttons when using XP themes were not displaying
   correctly.  (AXTDG000108)

 * Fixed horizontal scrollbar when ExtendRightColumn was set to true. (AXTDG000126)

 * When clicking on a column in the AddNew row where the column had a valuelist defined
   with CycleOnClick set to true was not dirtying the row.  (AXTDG000127)

   
==================================================================================
TrueDBGrid Pro Build Number 8.0.20051.336   Build Date: Tuesday, November 30, 2004
==================================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------

 * Added localization string (dbgpTipRefresh) for the Refresh button 
   in the print preview window. (AXTDG000096)

Corrections
-----------

 ** Pressing the Tab key while the dropdown was open was not moving focus to the
    next column.  (AXTDG000112)

 ** Merge cells weren't painting properly after horizontal scrolling when
    the merge property was set to dbgMergeRestricted.

 ** Removing the child grid from it's parent grid was not making the child
    grid visible.

 ** The columns.CellTop property was returning the incorrect value when the
    grid was scrolled.  (AXTDG000120)

 ** The dropdown in the FilterBar wasn't opening when using Alternating row
    styles and the Locked property was set.  (AXTDG000128)

===============================================================================
TrueDBGrid Pro Build Number 8.0.20044.333   Build Date: Monday, August 30, 2004
===============================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------

 * When using the ButtonPicture property, the grid now renders the image
   using the image size.  The height and width will be adjusted if the
   image size is wider or higher than the cell.


Corrections
-----------

 ** Fixed phantom vertical scrollbar being displayed when using 
    multiple splits. (AXTDG000082)

 ** Right clicking the grid and selecting Cut from the menu was not
    working.  (AXTDG000084)


===============================================================================
TrueDBGrid Pro Build Number 8.0.20044.331   Build Date: Thursday, July 29, 2004
===============================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------

 ** Grouping a column and then clicking in a cell the first time was returning
    an incorrect value for the ColContaining() method.  (AXTDG000080)

 ** Fixed GPF when using the FilterBar under Win98.  (AXTDG000078)

 ** Fixed "row flutter" when ScrollTrack was true and dragging the vertical
    scroll thumb to the bottom.

 ** Selecting a value from the combobox in the FilterBar wasn't firing
    the FilterChange() event.  (AXTDG000067)

 ** Fixed text truncation when exporting data and the grid's datasource
    was RDO.  (AXTDG000057)

 ** RecordSelectors were not displaying when ungrouping a grouped grid and
    the grouped grid's layout was loaded using the LoadLayout() method.  (AXTDG000061)


=============================================================================
TrueDBGrid Pro Build Number 8.0.20043.330   Build Date: Tuesday, May 11, 2004
=============================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------

  ** CTRL-Enter wasn't inserting a newline into the grid's cell editor. (AXTDG000065)

  ** The Bookmark parameter of the FormatText() event was incorrect for when
     printing/previewing selected rows.  (AXTDG000058)

  ** Displaying a MessageBox in the UnboundReadData event was throwing an exception.
     (AXTDBG000053)

=================================================================================
TrueDBGrid Pro Build Number 8.0.20042.328   Build Date: Monday, February 16, 2004
=================================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------

 ** Fixed stack overflow when updating an unbound column when the Locktype property of
    the datasource was set to adLockReadOnly.  (AXTDG000001)

 ** The button click event wasn't firing when you clicked on a button in a cell which
    was not on the current row. (AXTDG000031)

==================================================================================
TrueDBGrid Pro Build Number 8.0.20041.327   Build Date: Thursday, February 5, 2004
==================================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------

 ** Changing page orientation in Print Preview now correctly reflows the document for
    the select orientation. (AXTDG000020)


 ** Cells weren't rendering correctly for a hierarchical grid with more than 3 levels.
    (AXTDG000026)

 ** Updating a row using when using the recordset and CursorLocation=adUseServer was
    causing the grid to reposition the top row to the row updated. (8027)


===================================================================================
TrueDBGrid Pro Build Number 8.0.20041.326   Build Date: Thursday, December 11, 2003
===================================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------

 ** Fixed exception when pressing the Backspace/Delete key in a cell that contained
    and editmask. (AXTDG000009/XTDG000010)

 ** Fixed incorrect rendering of the expand/collapse icon in a Hierarchical grid that
    contained more than one level of relations.  (AXTDBG000018)

 ** Cells were incorrectly merged in a Hierarchical grid.

==================================================================================
TrueDBGrid Pro Build Number 8.0.20041.324   Build Date: Tuesday, November 25, 2003
==================================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------

 ** Calling the ExpandChild() when a child grid was already open was collapsing it. (AXTDG000005)

 ** Calling the ExpandChild() method after calling CollapseChild() was not displaying
    the open icon correctly. (AXTDG000007)

 ** Setting the FetchStyle property using the property pages was not persisting. (AXTDBG000008)

 ** Switching to Normal dataview from GroupBy dataview wasn't restoring the state of the
    record selectors. (AXTDG000011)


===================================================================================
TrueDBGrid Pro Build Number 8.0.20041.323   Build Date: Thursday, November 06, 2003
===================================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------

 ** Fixed GPF when the mouse was inside the grids client area using 
    XP Themes and switching dataviews. (8237)

 ** Assigning a picture to a button (ButtonPicture property) wasn't 
    displaying the image when using XP Themes. (8287)

 ** Setting focus to the grid using the Tab key when the form was 
    initially displayed wasn't correctly showing the marquee. (8295)

 ** Calling the Collapse() method of the grid to hide the child grid and
    then calling the Expand() method wasn't displaying the child grid. (8201)

 ** Calling the AutoSize() method for a column when you only have one 
    column in the split and the SizeMode is set to dbgNumberOfColumns 
    was reducing the size of the split when called more than once. (8200)

 ** FilterButtonClick event wasn't firing. (8230)

 ** Setting the CellStyle.Locked property in the FetchCellStyle event 
    wasn't working for checkbox columns. (7134)

 ** Setting the width of TrueDBDropdown in the DropDownOpen() event 
    wasn't working properly. (8199)

 ** Fixed GPF when pressing the backspace key after entering an invalid 
    character for an editmask that had more than one character selected. (8178)


==================================
TrueDBGrid Pro Build 8.0.20034.322
==================================

Corrected Problems
------------------

- The Bookmark property wasn't returning the bookmark of the newly added
  row in the AfterInsert() event. (8197)

==================================
TrueDBGrid Pro Build 8.0.20034.321
==================================

Enhancements/Documentation Changes
----------------------------------

- Added licensing support for non-standard containers.

Corrected Problems
------------------

- None.

==================================
TrueDBGrid Pro Build 8.0.20034.320
==================================

Enhancements/Documentation Changes
----------------------------------

- Cell selection is now available with an Inverted grid.

Corrected Problems
------------------

- Font property wasn't returning the correct charset.  (7990)

- Persisting a GroupBy grid wasn't correctly preserving the RecordSelector state when saving
  the layout.  (7906)

- Fixed paste operation.  (7702)

- ExportToDelimitedFile method was throwing an exception with invisible columns. (6973)


==================================
TrueDBGrid Pro Build 8.0.20033.319
==================================

Enhancements/Documentation Changes
----------------------------------

- Added DirectionAfterTab property.  Controls the new cell location after a tab key has been entered.

- TabAction property has been enhanced with a new enumeration: dbgGridColumnNavigation.  Setting the
  property to this value causes the focus to go to the next control in the tab order when the Tab key
  is entered on the last column of the last row.

- Columns.FetchStyle is now an enumeration with the following values:

   		dbgFetchCellStyleNone = 0          - No events will be fired to retrieve cell styles (default).
		dbgFetchCellStyleColumn = 1        - Events will fire for rows not including the AddNewRow.
		dbgFetchCellStyleAddNewColumn = 2  - Events will fire for rows including the AddNewRow.

  For the case of dbgFetchCellStyleAddNewColumn, the bookmark into the event will be null for the
  addnew row.

- Added WrapRowPointer property.  This is a boolean and effects the next cell location when the current
  cell is on the first or last row and the cursor up or down key is pressed.

Corrected Problems
------------------

==================================
TrueDBGrid Pro Build 8.0.20033.316
==================================

Enhancements/Documentation Changes
----------------------------------

Corrected Problems
------------------

- Fixed GPF in property pages in VC++ when choosing the "All" tab. (7696)

- Fixed copy/paste in a cell. (7702, 7760)

- Tabbing out of the grid when using an External Editor and having a number format
  for the cell, wasn't display the cell as formatted.  (7762)

- BeforeUpdate() event wasn't firing in the correct order for a Hierarchical datasource. (7794)

==================================
TrueDBGrid Pro Build 8.0.20032.311
==================================

Enhancements/Documentation Changes
----------------------------------

Corrected Problems
------------------

- Fixed default size of rows in FormView.

- Fixed UI when rearranging the order of grouped columns.

- Change to/from a MultiLine grid wasn't resetting the horizontal scrollbars.

- Fixed rendering of the connecting line for child grids.

- Setting a background picture of a cell wasn't drawing the Collapse/Expand icons.

- Having a grid that had a child grid wasn't showing the split sizer cursor.

==================================
TrueDBGrid Pro Build 8.0.20032.310
==================================

Enhancements/Documentation Changes
----------------------------------

- Editing vertically centered cells no longer requires WordWrap = true to edit vertically.

Corrected Problems
------------------

- Fixed rendering in Form DataView when column headers were invisible.

- Fixed display problem in MultiLine mode and forcing the visibility of
  scrollbars.

- Displaying a multiline grid as a child grid was not sizing the child grid
  correctly when displayed.

- Fixed exception when using the ExpandChild()/CollapseChild() methods and no
  ChildGrids were being used.

- Fixed problem when using key navigation to exit the FilterBar with multiple splits.

==================================
TrueDBGrid Pro Build 8.0.20032.308
==================================

Enhancements/Documentation Changes
----------------------------------

Corrected Problems
------------------

- Grid wasn't rendering correctly after changing the Desktop Theme.

- Split caption wasn't painted properly under Theme appearance.

- Columns weren't rendered correctly under XP Themes in Inverted dataview as
  as the mouse was moved over column headers.

- Fixed problem when trying to Fix a column that wasn't visible (i.e., not scrolled
  into the client area).

- Setting the Style.BackgroundPicture or Style.ForegroupPicture now throws an error
  if the bitmap is too large (approximately 256k).

==================================
TrueDBGrid Pro Build 8.0.20032.306
==================================

Enhancements/Documentation Changes
----------------------------------

Corrected Problems
------------------

- Fixed print problem when the screen was set to a non-square resolution (e.g., 1280x1024).  (6105)

- Printing an Inverted grid was causing an unspecified dll error.

==================================
TrueDBGrid Pro Build 8.0.20032.305
==================================

Enhancements/Documentation Changes
----------------------------------

 - The ValueItems.MatchCase property is now applicable to the built in combo.

Corrected Problems
------------------

 - Scrolling the grid horizontally in Normal dataview and then setting the
   dataview to Inverted was not setting the topmost column correctly. (7535).

 - Fixed unspecified dll error when printing an empty grid with ValueTranslate = true. (7487)

 - The expand/collapse icons were not being painted after switching dataview. (7404)

 - GroupHeadClick event wasn't firing when all the columns were grouped. (7490)

 - Fixed problem when setting the vertical/horizontal scrollbar size. (7605)

 - Fixed problem initiating edit on the FilterBar when the current row was
   not visible.  (7509)

 - Fixed column dividers not painting properly in multi-line mode and 
   extendrightcolumn property was true. (7616)

 - Toggling between Inverted and Normal dataviews were losing the fix column setting. (7534)

 - Fixed RefetchCell() method not working in unbound mode. (6939, 7617)

==================================
TrueDBGrid Pro Build 8.0.20031.303
==================================

Enhancements/Documentation Changes
----------------------------------

Corrected Problems
------------------

 - Reposition the current row position in code and having a cell locked by setting the 
   the style property in the FetchCellStyle() event was not unlocking the cell on the new row.

 - Setting the DirectionAfterEnter to None was preventing the default button's click event
   to fire on the form when the enter key was pressed.

 - FetchCellStyle() event was firing twice for the first column when the grid's datasource
   was changed.

==================================
TrueDBGrid Pro Build 8.0.20031.301
==================================

Enhancements/Documentation Changes
----------------------------------

Corrected Problems
------------------

 - Fixed scrolling problem in FormView.

 - Fixed font problem.  Setting a font with a different charset
   was improperly setting all font charsets.

==================================
TrueDBGrid Pro Build 8.0.20031.300
==================================

Enhancements/Documentation Changes
----------------------------------

Corrected Problems
------------------

 - Fixed problem when using the Tab key to navigate the grid while the grid contained some fixed columns.

==================================
TrueDBGrid Pro Build 8.0.20031.299
==================================

Enhancements/Documentation Changes
----------------------------------

Corrected Problems
------------------

 - The ExportToDelimtedFile() method wasn't correctly exporting column headers in the correct order.

 - The horizontal scrollbar button wasn't working correctly in GroupBy DataView.

==================================
TrueDBGrid Pro Build 8.0.20031.298
==================================

Enhancements/Documentation Changes
----------------------------------

Corrected Problems
------------------

 - Selecting a row which contained a merged column wasn't highlighting the entire merged column.

 - Checkboxes were displayed incorrectly when the appearance property was set to XP Themes.

 - Fixed links in the about box.

==================================
TrueDBGrid Pro Build 8.0.20031.297
==================================

Enhancements/Documentation Changes
----------------------------------

 - Licensing update.

Corrected Problems
------------------

==================================
TrueDBGrid Pro Build 8.0.20031.296
==================================

Enhancements/Documentation Changes
----------------------------------

Corrected Problems
------------------

 - In some circumstances, refreshing the datasource when using the Floating Editor
   Marquee was causing incorrect text to be displayed in the editor after the refresh.

 - Editing an empty cell that was vertically centered wasn't display the caret.

 - Keypress event wasn't firing for the enter keywhen using a Marquee other than 
   Floating Editor.

 - Printing or exporting the grid whose cell value was being translated by the 
   dropdown control, wasn't translating values for cells that hadn't been visited.

 - Unable to copy cells to clipboard if a column was locked.

 - Columns(x).Text property was not returning the translated value if the translation
   was being done by the dropdown control and the cell hadn't been visited.

 - Decreasing the height of the grid was changing the toprow.

 - Resizing the parent grid while a child grid was open, wasn't correctly painting the
   parent grid.

 - Setting the ConvertEmptyCell property wasn't working when the grid was bound to a
   shaped datasource.

 - The grid was selecting cells when the MultiSelect property was set to extended and
   it received a mouse move message without first getting a mouse down.

 - Fixed problem with scrollbars not painting properly with multiple splits.

 - Grid caption wasn't displaying when using XP Themes.

 - Columns captions using XP Themes weren't painted properly when the mouse was moved 
   from the caption to the vertical scrollbar.

==================================
TrueDBGrid Pro Build 8.0.20024.293
==================================

Enhancements/Documentation Changes
----------------------------------

This build of TrueDBGrid now supports XP themes.  A new option in available for the
Appearance property (dbgXPTheme).  Setting this property displays the grid using XP 
theme elements when run on an operating system that supports Themes.  If this property
is set and the operating system does not support themes, it renders using the traditional
3D look.


You can now fix a column (keep it from scrolling) using the Column.Fix(boolean fix, short pos) method.
This method takes two arguments.

fix - True/False.  True to fix the column.
pos - Where to position the column if you fix more than one.


Split specific property ShowCollapseExpandIcons.  You can use this to disable the displaying of
expand/collapse icons.

Two new methods, ExpandChild()/CollapseChild() to display/hide child grids programmatically.

A new property MatchCase on the ValueItems object.  Setting this to true allows one to only
do case sensitive searches for Translations.

The Column.Merge property was enhanced to allow restricted type merging if the previous
column is merged and it's merge range is less than the current columns merge range.

Printing has been enhanced to print Form and Inverted dataviews.

Column Selection and AutoSize() are now supported in Inverted Dataview.

Corrected Problems
------------------

6856 - If MarqueeStyle is set to Highlight Cell, grid generates an error "property 
	   is not available  in this context" in BeforeColChange event handler.

6841 - If MarqueeStyle is default and you Alt+Tab to another application, then         
	   AfterColUpdate event does not fire.

6970 - If you set to form view, then the column.Left and the column.Width will          
	   return the same results as in the normal view.

6454 - In Inverted view, can not select the record by clicking the recordselector.

7020 - Using the Update method in the AfterColEvent in UnboundClassic mode was
	   throwing and exception.

7067 - Unable to select from the dropdown using the keyboard if the cell text did
	   not match an entry in the dropdown.

7070 - Fixed horizontal scrollbar position when ExtendRightColumn was set to true.

7071 - Fixed paint problem with a grid that contained no columns.

