Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("BARCODES")> 
<Assembly: AssemblyDescription("Contains Objects Meant to Facilitate the Manipulation of Barcodes of Diverse Formats")> 
<Assembly: AssemblyCompany("TransTech Software, Inc.")> 
<Assembly: AssemblyProduct("")> 
<Assembly: AssemblyCopyright("2009")> 
<Assembly: AssemblyTrademark("")> 
<Assembly: CLSCompliant(True)> 

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("BF966935-F362-4BFF-AF96-B83B5D5B88CA")> 

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
' You can specify all the values or you can default the Build and Revision Numbers 
' by using the '*' as shown below:

<Assembly: AssemblyVersion("1.1.*")> 
' V1.1.* - 1. Adds the following fields to the definition of a ScanList record...
'               iScanError
'               dProcessedDate
'               iProcessed
'       -   2. Implements _Records as an in-memory copy of the current file, whether opened for reading or writing
