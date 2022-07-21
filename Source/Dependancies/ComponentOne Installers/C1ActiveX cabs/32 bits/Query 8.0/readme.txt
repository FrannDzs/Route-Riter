==========================================================================
ComponentOne Query Control Version 8.0.20141.38     Build Date: 03/21/2014
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None

Corrections
-----------
-- IE11 compatibility

==========================================================================
ComponentOne Query Control Version 8.0.20111.37     Build Date: 03/10/2011
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None

Corrections
-----------
-- Fixed bug in VB6: Crash opening a form containg a C1Query component.


==========================================================================
ComponentOne Query Control Version 8.0.20111.35     Build Date: 01/06/2011
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
Rebuilt C1Query8.ocx using Visual Studio 2010, for compatibility
with 64 bit.

==========================================================================
ComponentOne Query Control Version 8.0.20091.34     Build Date: 01/08/2008
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None

==========================================================================
ComponentOne Query Control Version 8.0.20082.33     Build Date: 04/30/2008
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
-- Now ComponentOne Query Control can work with DEP on 
   (DEP="data execution prevention")

==========================================================================
ComponentOne Query Control Version 8.0.20082.32     Build Date: 04/21/2008
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None

==========================================================================
ComponentOne Query Control Version 8.0.20081.31     Build Date: 01/15/2008
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
Licensing change.

==========================================================================
ComponentOne Query Control Version 8.0.20081.30     Build Date: 01/03/2008
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None

==========================================================================
ComponentOne Query Control Version 8.0.20073.29     Build Date: 01/03/2008
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
Licensing change.

==========================================================================
ComponentOne Query Control Version 8.0.20072.28     Build Date: 03/02/2007
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None

==========================================================================
ComponentOne Query Control Version 8.0.20071.27     Build Date: 10/11/2006
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None

==========================================================================
ComponentOne Query Control Version 8.0.20063.26     Build Date: 08/24/2006
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------
Repaired VC++ tutorial projects.

==========================================================================
ComponentOne Query Control Version 8.0.20062.24     Build Date: 04/20/2006
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------
None

==========================================================================
ComponentOne Query Control Version 8.0.20061.23     Build Date: 12/18/2005
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------
None


==========================================================================
ComponentOne Query Control Version 8.0.20053.22     Build Date: 08/17/2005
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------
None


==========================================================================
ComponentOne Query Control Version 8.0.20052.21     Build Date: 04/19/2005
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------
-- Fixed bug: C1QueryFrame.Item.ID property returned incorrect value 
   after BuildSQL method call under certain circumstances.

-- Fixed a problem in the CheckBoxes sample.



==========================================================================
ComponentOne Query Control Version 8.0.20052.20     Build Date: 03/08/2005
==========================================================================

Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------
-- Fixed bug: Last edited Lookup value was not always saved in the 
   Schema Designer when schema is saved using Save/Save As menu item.


==========================================================================
ComponentOne Query Control Version 8.0.20051.19     Build Date: 11/03/2004
==========================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------
-- Fixed bug: C1Query generated incorrect SQL
statements for "None" and "Not All" connectives in complex conditions.


==========================================================================
ComponentOne Query Control Version 8.0.20051.18     Build Date: 11/01/2004
==========================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------
-- Fixed bug: LookupValues event did not work correctly in C++ and .NET projects.

-- Fixed a problem with Windows XP Service Pack 2: Security warning dialog box
    was displayed when a C1QueryFrame control is loaded.

-- Fixed a bug in the TDBCombo sample.


==========================================================================
ComponentOne Query Control Version 8.0.20044.17     Build Date: 08/03/2004
==========================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------
None


==========================================================================
ComponentOne Query Control Version 8.0.20043.16     Build Date: 06/20/2004
==========================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------
None


Corrections
-----------
Fixed bug: C1QueryFrame.ShowFolderFields did not work correctly with custom PathSeparator property value.



==========================================================================
ComponentOne Query Control Version 8.0.20043.15     Build Date: 05/01/2004
==========================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------

- Added property: C1Query.PathSeparator. This property value is used as separator between folders and folder names in folder paths.
This separator appears in C1QueryFrame controls when the FullFolderNames property is set to True. It is dot (".") by default.
Setting this property to a string other than ".", you can include dots in folder and folder field names.


----------------------
Documentation changes:
----------------------

----------
New topic:
----------
  Under Property Reference->Quick Reference for All Properties:

PathSeparator property
----------------------

Syntax

  C1Query.PathSeparator = string

Notes

  Read/Write at run time and design time. Property applies to C1Query control.

Description

This property value is used
as separator between folders and folder names in folder paths.
This separator appears in C1QueryFrame controls when the FullFolderNames property is set to True. It is dot (".") by default.
Setting this property to a string other than ".", you can include dots in folder and folder field names.

----------------
Modified topics:
----------------

FullFolderName property
-----------------------

To the following line in Description section:

This property returns full dot-separated folder name of a field represented by the Field object

add:

(C1Query.PathSeparator is used instead of dot, if set by the programmer)


FieldNameToFieldID Method
-------------------------

To the following line in Arguments section:

FullName is the full dot-separated name of a folder field in a C1Query schema.

add:

(C1Query.PathSeparator is used instead of dot, if set by the programmer)


FullFieldNames property
-----------------------
To the following line in Description section:

If this property is set to True, full dot-separated names including full folder name are used in query display.

add:

(C1Query.PathSeparator is used instead of dot, if set by the programmer)


General Appearance/Behavior Features
------------------------------------

To the following line:

The FullFieldNames property set to True makes the control show folder fields with full dot-separated folder path.

add:

(C1Query.PathSeparator is used instead of dot, if set by the programmer)


C1QueryFrame Object
-------------------
To the following line in C1QueryFrame Object Properties table:

If True, the control shows full dot-separated field names containing folder names.

add:

(C1Query.PathSeparator is used instead of dot, if set by the programmer)


SetLookup Method
----------------
To the following line in Arguments section:

FullFieldName is a string argument specifying the full dot-separated name of a folder field in a C1Query schema.

add:

(C1Query.PathSeparator is used instead of dot, if set by the programmer)


ShowFolderFields Method
-----------------------

To the following in Arguments section:

FolderPath specifies a dot-separated name of a folder field or an entire folder in a C1Query schema.

add:

(C1Query.PathSeparator is used instead of dot, if set by the programmer)


Quick Reference for All Properties
----------------------------------

To the following lines:

If True, the control shows full dot-separated field names containing folder names.

Full dot-separated folder name of a field represented by this Field object.

add:

(C1Query.PathSeparator is used instead of dot, if set by the programmer)


Field Object
------------

To the following line in Field Object Properties table:

Full dot-separated folder name of a field represented by this Field object.

add:

(C1Query.PathSeparator is used instead of dot, if set by the programmer)




Corrections
-----------

Fixed bugs:
 - Field.Exists property returned incorrect value if called before query generation/menu item selection.
 - QueryFrame.LoadFromXML/LoadFromXMLFile worked incorrectly in DataSource mode.




==========================================================================
ComponentOne Query Control Version 8.0.20042.14     Build Date: 02/03/2004
==========================================================================


Enhancements/Documentation/Behavior Changes
-------------------------------------------
None

Corrections
-----------
None



===============================================
ComponentOne Query Control version 8.0.20041.13
===============================================

==================
Corrected problems
==================

-- Fixed bug: Setting Field.DisplayName in DataSource mode changed database 
   field name used in the generated SQL statement, making the statement invalid.
-- Fixed bug: Exception occured under certain circumstances in C1QueryFrame
   control when the user adds a new query item using context menu.


===============================================
ComponentOne Query Control version 8.0.20041.12
===============================================
Only version number change.


===============================================
ComponentOne Query Control version 8.0.20034.11
===============================================

==================
Corrected problems
==================

-- QueryItem.HasLookup property worked incorrectly in DataSource mode.


===============================================
ComponentOne Query Control version 8.0.20033.10
===============================================
==========
What's New
==========

-- Added a folder field Tag property. This string property can be used to store arbitrary
   additional information associated with the field. It can be set for folder fields
   in Schema Designer at design time and retrieved (and changed) at run time as Field.Tag.

-- Added QueryItem.Tag (string) property. It can be used to store arbitrary additional information
   associated with query items. This property is persisted along with other QueryItem properties
   by C1QueryFrame SaveToXML/LoadFromXML methods.

==================
Corrected problems
==================

-- SetLookup method worked incorrectly in DataSource mode when some of the fields are hidden with 
   Field.DataSourceName = "".

-- LoadFromXML method worked incorrectly in DataSource mode.

-- QueryItem.Alias property was not persisted by C1QueryFrame SaveToXML/LoadFromXML methods.


==============================================
ComponentOne Query Control version 8.0.20033.9
==============================================
==========
What's New
==========

-- Field.HasLookup property added.

-- Schema Designer: Move up/Move down commands added for Folders toolwindow context menu.

==================
Corrected problems
==================

-- Fixed minor Schema Designer bug: changing folder field name in Folder
   Designer does not refresh field name in the Folders toolwindow.

-- Bug with drap&drop support in Folders toolwindow (names are set back to
   View field names instead of folder field names).


==============================================
ComponentOne Query Control version 8.0.20032.8
==============================================
==========
What's New
==========

-- It is now possible to use C1Query in web applications, to submit queries from a thin
   browser-based client to the server and generate SQL statement on the server.
   A new sample WebQueryASP (see subdirectory Samples\ASP) shows how to do that using ASP
   (Active Server Pages). C1Query is used to generate SQL query statements from user input received
   from a browser client. C1Query is used only on the server, the client is a generic HTML browser,
   no ActiveX or other plug-ins required. See readme.txt in the sample directory for details.

-- Published the XML structure of a C1Query schema, see HTML file c1query_schema_format.htm.
   This allows to create schema at run time without using Schema Designer.

-- Added new sample RunTimeSchemaInit (subdirectory Samples\VB) showing how to create and use a schema
   at run time without using Schema Designer.

-- It is now possible to specify a join condition in Schema Designer as an arbitrary SQL condition.
   In this case, the condition determines the entire join condition, without differentiating between
   the right and left parts and operator. This new feature allows to specify joins 
   with special operators or with expressions in both parts, which was not possible before.
   To specify a SQL condition in a join, check the "Use SQL condition" check box in the join
   editor dialog and type the SQL condition text in the text box. In the SQL condition, use
   operators and functions according to the SQL syntax, and use the following notation for
   fields:

   [ViewName]![TableName].[FieldName]

   For example, join Orders.EmployeeID = Employees.EmployeeID can be represented in this form as

   [Orders]![Orders].[EmployeeID] = [Employees]![Employees].[EmployeeID]

==================
Corrected problems
==================
-- C1QueryFrame.ExitEdit(True) did not properly finish lookup editing.

==============================================
ComponentOne Query Control version 8.0.20031.6
==============================================
==========
What's New
==========

Subscription licensing scheme, new About box, all incremental updates and fixes 
applied to previous versions. If you haven't been dowloading the latest patches
from our web site periodically, this is an easy way to get everything in one step.

==================
Corrected problems
==================

-- C1Query now works properly as an ActiveX control in .NET environment.

-- Fixed Tutorial 8. It was failing on any other constant type than string
   (try condition Categories.CategoryID = 1).

-- Advanced elementary condition with left and right constants:
   DoBeforeEditing2 raised script error.

-- Fixed bug in schema designer:
   Relationship designer->Binary relationship->Changing right view crashed
   if there were join conditions with right side expression.

-- Fixed bug introduced in version 1.1 that caused a crash under certain
   circumstances in editing a query after generating SQL with errors,

-- Fixed bug in the schema designer:
   Crash in Drag&Drop a field from a View designer to the empty area of the Folders toolwindow. 

-- Add advanced elementary condition, set right side kind to Field but leave it empty.
   Then change comparison to Empty or Non-empty. After that SQL generation failed
   (it expected a real field in the right side, which it should not).

-- The Boolean field allowed comparisons such as <, >, etc.
   Now the Boolean type allows only Equal, Not equal, Empty and non-empty comparisons.

-- Fixed bug in database structure import: Table list empty when UserID is non-empty
   and tables belong to a non-empty schema different from the UserID. For example,
   in the standard SQL Server Northwind sample database, the schema is "dbo", and
   the UserID is usually different from "dbo".

-- In DataSource mode, after clearing query (C1QueryFrame.Clear), the
   C1Query.ClauseWhere property retains its previous value, is not cleared (it should
   be set to ""), and subsequent BuildQuery calls do not change
   ClauseWhere while query remains empty.
   
-- If field alias is already enclosed (by the user, in schema designer) in quotes,
   as in "[My Alias]", then it is copied to the resulting SQL as is, without substituting
   non-alphanumeric characters with underscores. For example, alias "[My Alias]"
   was represented as _My_Alias_ before, now it is represented as [My Alias].
   
