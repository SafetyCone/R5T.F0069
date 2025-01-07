# R5T.F0069
Excel automation types library.

## Goals

1. Provide Excel Application, Workbook, Worksheet, Range, and other types wrapping interop types.
    * Wrapping types reference their parent wrapping types.
2. To provide InXContext() methods for querying and modifying Excel files.


## Prior Work

* R5T.Private.Palmyra (uses R5T.Excel in an application)
* R5T.Excel (best prior)
* R5T.Private.Agrippina (exploration)
* Minex/Common/Libraries/Excel
* Florence – Public/Common/Libraries/Excel


## Design

* Only one library with the COM reference, all Excel-related functionality (including wrapping types) are in that one library.
* Sub-libraries for values and types that do not use any interop functionality.
* No interop types are exposed from types in the main R5T.F0069 namespace. This is for ease of use.
  * (In a client library, you *can* use instances of types that expose interop types in their interface without the client library needing the COM reference as long as you only use the parts of the interface that do *not* expose the interop types. But as soon as you use any of the part that exposes interop types, you need to add the COM reference to the client project.)
* However, interop types *are* exposed in the R5T.F0069.Interop sub-namespace.
    * Wrappers use the "internal" access modifier for properties exposing their underlying interop type.
    * Public extension methods in a different namespace can then allow public access to the underlying interop type.
    => This way, the underlying interop type instance *is* available if a consuming library really wants it, but it won't be "tripped" over when using the main namespace.
* Most type functionality is provided via extension methods, calling registered functionality methods, that call Interop-namespace functionality methods.



## Excel COM Automation

## Orphaned Excel Processes

All Excel COM automation object wrappers need to be disposable.

Links:

    - https://www.add-in-express.com/creating-addins-blog/release-excel-com-objects/

### Only Windows Executables

This library is a .NET Standard library, which implies cross-platform capabilities if used from a .NET Core or .NET 5.0+ application entry-point. However, Excel COM automation only works if Excel is present on the machine *and* COM automation is available. Since COM automation is a Windows-only technology, the entry-point application must be Windows-only.


### How To

To set a COM automation reference to the Excel interop libraries, make sure you have a .NET 6.0 project ([see the note below](#and-.net-standard) for .NET Standard).

The right-click project references, add a COM reference, and select "Microsoft Excel 16.0 Object Library".

Do not mess with the "Embed Interop Types" option, leave it as true.


### Special considerations due to COM automation limitations

It is not possible to use the ordinary "selector" library methodology for COM automation. For regular .NET assemblies, assembly references are transitive. This is to say that consumer library C can use types from provider library A, even though it only has a reference to library B which in turn references A. For COM automation, this appears not to be possible; all projects using interop types must directly reference the COM automation library.

This is not to say that all projects referencing a project that internally uses interop types must reference the COM automation, instead it means that if the referenced project exposes an interop type the referencing project wants to use it (for example, a library has a function producing an instance of an interop type and the consumer calls that function and now has a variable of the interop type, the consumer must reference the COM automation.


### .NET 5.0+ is ok

COM automation was not available in .NET Core. Back then the entry point application used to have to use the Windows-only .NET Framework.

However, Excel COM automation (like all COM automation) is Windows-only. Thus is was not really a problem that the .NET Standard library can only be consumed from .NET Framework, which only exists on Windows.

However, now


### Any CPU is ok as of .NET 6.0

Relative to x32 or x64, previously the Excel interop assemblies had a limitation that the entry-point application bit-ness (x32 or x64) MUST match the installed Excel application bit-ness. Otherwise a cryptic error appeared at runtime:

    System.InvalidCastException: 'Unable to cast COM object of type 'Microsoft.Office.Interop.Excel.ApplicationClass' to interface type 'Microsoft.Office.Interop.Excel._Application'. This operation failed because the QueryInterface call on the COM component for the interface with IID '{000208D5-0000-0000-C000-000000000046}' failed due to the following error: Invalid class string (Exception from HRESULT: 0x800401F3 (CO_E_CLASSSTRING)).'

### And .NET Standard

This library references the COM assembly "Microsoft Excel 16.0 Object Library", showing that YES, it is possible for .NET Standard library to perform COM automation.

While adding a COM automation reference is not available from the Visual Studio right-click menu for .NET Standard projects, it *is* possible to add a COM reference by either:

* Manually add the COM reference to the project:

    <Project Sdk="Microsoft.NET.Sdk">
      ...
      <ItemGroup>
        <COMReference Include="Microsoft.Office.Interop.Excel">
          <WrapperTool>tlbimp</WrapperTool>
          <VersionMinor>9</VersionMinor>
          <VersionMajor>1</VersionMajor>
          <Guid>00020813-0000-0000-c000-000000000046</Guid>
          <Lcid>0</Lcid>
          <Isolated>false</Isolated>
          <EmbedInteropTypes>true</EmbedInteropTypes>
        </COMReference>
      </ItemGroup>
    </Project>

* Or, temporarily make the project a .NET 6.0 project, use the right-click menu to add the COM reference, and then switch the project back to .NET Standard.
