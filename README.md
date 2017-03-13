VBA Code Helper
===============

[![The MIT License](https://img.shields.io/badge/license-MIT-orange.svg?style=flat-square)](http://opensource.org/licenses/MIT)
[![The Latest Release](https://img.shields.io/badge/release-v1.0-blue.svg?style=flat-square)](https://github.com/osevoso/VBACodeHelper/releases/tag/v1.0)

Overfiew
--------

This is simple add-in for VBA IDE in 64- and 32-bit host apps (such as Excel, Word, AutoCAD etc), which provides **VBA code indent tool** via Сode Window context menu items, and for additionally, two hotkey:

>Ctrl + Alt + *"Gray Plus" key*

for comment code lines,

>Ctrl + Alt + *"Gray Minus" key*

for uncomment code lines.

---

Download
--------

Archived binaries for 64- and 32-bit host apps are available at the [Releases][1] page.

>Note that for 64- and 32-bit apps, you must use a corresponding bit DLL.

Install
--------

>Before install, be shure what hotkeys of your background services, OS shell and other add-ins is not same with the add-in hotkeys.

1.  Unpack DLL into local drive folder (*C:\Addons\VBACH*, for example)

2.  Run with administrator credits at Windows command line: 

    For 64-bit host apps on 64-bit OS and for 32-bit host apps on 32-bit OS

        C:\Windows\regsvr32.exe C:\Addons\VBACH\VBACodeHelper.dll

    For 32-bit host apps on 64-bit OS

        C:\Windows\SysWOW64\regsvr32.exe C:\Addons\VBACH\VBACodeHelper.dll

Deinstall
--------

Use the same command with "/**u**" command line key, for example: 

    regsvr32.exe /u C:\Addons\VBACH\VBACodeHelper.dll
    
-Restrictions
---------

Code lines, which concatenated by undescore at the end of phisical line, would be concat into single phisical code line. 
If its undesirable for you, use "Selected Lines" indent mode for skip their where necessary.    

Credits 
--------

The mechanics of indenting VBA code is based on the original algorithm of Michael Ciurescu author [VBTools AddIn][2].

[1]: https://github.com/osevoso/VBACodeHelper/releases/
[2]: http://www.vbforums.com/showthread.php?479449-VBTools-AddIn-Auto-indent-VB-code-!
