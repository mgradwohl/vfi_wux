# vfi_wux
Port of VFI to Windows App SDK + WinUI

This is the port of VFI from the Windows NT Resource Kit to Windows App SDK + WinUI in C++.

The code builds.

Main files/classes to care about are MainWindow (of course), helpers, CWiseFile.

TODO
clean up includes
move to std::wstring and other modern C++ types
add the ability to add at least one file and get the information showing in the listview
determine the best way to spin up a thread to calculate the CRC
