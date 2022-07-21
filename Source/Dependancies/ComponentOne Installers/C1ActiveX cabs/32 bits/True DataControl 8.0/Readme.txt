==================================================
ComponentOne True DataControl version 8.0.20091.34
==================================================
No changes, only version number changed.

==================================================
ComponentOne True DataControl version 8.0.20082.33
==================================================
-- Now True DataControl can work with DEP on 
   (DEP="data execution prevention")

==================================================
ComponentOne True DataControl version 8.0.20082.32
==================================================
No changes, only version number changed.

==================================================
ComponentOne True DataControl version 8.0.20081.31
==================================================
Licensing change.

==================================================
ComponentOne True DataControl version 8.0.20081.30
==================================================
No changes, only version number changed.

==================================================
ComponentOne True DataControl version 8.0.20073.29
==================================================
Licensing change.

==================================================
ComponentOne True DataControl version 8.0.20072.28
==================================================
No changes, only version number changed.

==================================================
ComponentOne True DataControl version 8.0.20071.27
==================================================
No changes, only version number changed.

==================================================
ComponentOne True DataControl version 8.0.20063.26
==================================================
No changes, only version number changed.

==================================================
ComponentOne True DataControl version 8.0.20062.25
==================================================
No changes, only version number changed.

==================================================
ComponentOne True DataControl version 8.0.20061.24
==================================================
No changes, only version number changed.

==================================================
ComponentOne True DataControl version 8.0.20053.23
==================================================
No changes, only version number changed.

==================================================
ComponentOne True DataControl version 8.0.20052.22
==================================================
No changes, only version number changed.

==================================================
ComponentOne True DataControl version 8.0.20051.21
==================================================
-- Fixed bug: Occurs only in Windows 98: Editing a True DBGrid cell bound to a TData field with ModificationMode = '1 - Immediate',
   whole cell is selected after very keystroke.

==================================================
ComponentOne True DataControl version 8.0.20044.18
==================================================
-- Fixed bug: TrueData Control crashed when created in a container that uses aggregation, e.g., in ActiveX Control Test Container.

==================================================
ComponentOne True DataControl version 8.0.20043.17
==================================================
-- Fixed bug: Validation rule coded in WillChangeRecord event was not enforced on closing form, if the user enters incorrect value
   in a control and closes the form without leaving that control.

-- Fixed bug: Crash setting the value of an invisible field (Field.Visible=False) in code (setting TData.Fields("fieldname").Value).

==================================================
ComponentOne True DataControl version 8.0.20042.16
==================================================
No changes, only version number changed.

==================================================
ComponentOne True DataControl version 8.0.20041.15
==================================================
No changes, only version number changed.

==================================================
ComponentOne True DataControl version 8.0.20034.14
==================================================
==================
Corrected problems
==================
-- Fixed bug: Single quote in double quotes in an expression text generates an error.
   For example, an expression A = " a'b" is evaluated as A = " a'b"" and generates an error.

-- Minor changes in licensing.

==================================================
ComponentOne True DataControl version 8.0.20033.13
==================================================
Licensing fix.

==================================================
ComponentOne True DataControl version 8.0.20033.12
==================================================
==================
Corrected problems
==================
-- Fixed bug: TData with only bound text boxes, with a SQL Server data source, e.g., Customers table in Northwind.
   Enter illegal record, e.g., enter existing (duplicate) CustomerID, press Next button.
   A generic "Error occured" error is reported instead of the specific database error (UPDATE failed because of
   duplicate key, etc).

==================================================
ComponentOne True DataControl version 8.0.20032.11
==================================================
Minor changes in licensing and AboutBox.

==================================================
ComponentOne True DataControl version 8.0.20031.10
==================================================
==================
Corrected problems
==================

-- Fixed bug: If stored procedure fails on some parameter values,
   creating recordset can fail at design time.

=================================================
ComponentOne True DataControl version 8.0.20031.9
=================================================
==================
Corrected problems
==================
-- Fixed a bug (caused by an MS Client Cursor Engine bug):
   With CursorLocation=adUseClient, LockType = adLockBatchOptimistic,
   add a new row (call TData.Recordset.AddNew), set the row's primary key to a non-unique value,
   call TData.Recordset.UpdateBatch. An error occurs (correct behavior). Trap the error, continue.
   Change the primary key value to a unique one, call TData.Recordset.UpdateBatch again. The update
   is performed correctly (data in the database changed), but TData.Recordset.UpdateBatch
   returns with error "Row handle referred to a deleted row or a row marked for deletion".

=================================================
ComponentOne True DataControl version 8.0.20024.5
=================================================

==========
What's New
==========

-- You can bind True DataControl at design time to any
   ADO/OLE DB data source (for example, to VB DataEnvironment)
   using new properties DataSource and DataMember.
   Binding with DataSource/DataMember makes TData use it
   as SourceRecordset instead of creating SourceRecordset
   at run time.

-- True DataControl now supports side-by-side specification.

-- Subscription licensing scheme, new About box

-- All incremental updates and fixes applied to previous versions.
   If you haven't been dowloading the latest patches from our web
   site periodically, this is an easy way to get everything in one step.


==================
Corrected problems
==================

-- TData.Recordset.Find worked incorrectly when TData.SourceRecordset.Filter/Sort is set.

-- Eliminated memory leak refreshing TData.

