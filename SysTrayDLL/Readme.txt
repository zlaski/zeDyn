To use SysTrayDll in your own projects, follow the steps below:

1)Compile the dll:
  In the VB IDE
  Click on SysTrayDll in the project group window
  click File/Make SysTrayDll.dll

2)Register the dll: 
  In Windows:
  Click on the start menu & select Run
  Type Regsvr32 {drive}:\{dll path}\SysTrayDll.dll

3)Reference the dll in your own project:
  In the VB IDE click Project/References 
  Select SysTrayDll from the list and click OK
  
4)In the General Declarations portion of a form or class module, add:
  Dim WithEvents SysTray as SysTrayDll.SysTray

5)In the form_load procedure add
  Set SysTray = New SysTrayDll.SysTray

=============
The SysTray object is now ready to use. See the frmTest.frm for an example fo how to use it.

