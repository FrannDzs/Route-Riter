
TrueDBGrid 8.0c

What is it?

 TrueDBGrid 8.0c is the verion 8.0 ActiveX control with different guids and classnames.
 Version 8.0c provides an upgrade path from TrueDBGrid 7.0 without having to change your 
 version 7.0 code.   The current migration path from verion 7.0 to 8.0 involves changing 
 code when using the FetchCellStyle property.  This isn't needed for 8.0c.

 Unlike verion 8.0, the FetchStyle property is defined as a variant.  So you can start 
 using the new FetchCellStyle enum (dbgFetchCellStyleAddNewColumn) if you wish.  Later 
 conversions to the standard 8.0 version will not be impacted.

 All enhancements and corrections are refelected in both verions of the control.


How do I migrate from Version 7 to 8.0c?

 You can use the migration utility and migrate your project from 7.0 to 8.0.  To migrate 
 your application to 8.0c, use the provided migration utility, Migrate8to8c.exe.

