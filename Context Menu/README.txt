Here are some instructions for using the context menu extension.

ContextMenuExt.dll is designed to work with FindFile.exe.

1. Put the latest version of findfile where you will be using it. 
Startup the new findfile and then shut it down, this will add a new 
entry in the registry.

2. Put ContextMenuExt.dll in windows\system

3. Then Run -> regsvr32 contextmenuext.dll

either from a console window or from the 'Run' dialog on the windows start
button on the taskbar. This will register the dll for shell menu context
handlers.

4. Right click on any file in explorer and select 'FindFile Search...'

5. To uninstall the context menu extension Run -> regsvr32 /u contextmenuext.dll

You can of course change ContextMenuExt.dll to run anything else.

John

