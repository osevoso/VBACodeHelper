VBA Code Helper
===============

[![The MIT License](https://img.shields.io/badge/license-MIT-orange.svg?style=flat-square)](http://opensource.org/licenses/MIT)
[![The Latest Release](https://img.shields.io/badge/release-v1.0.2-blue.svg?style=flat-square)](https://github.com/osevoso/VBACodeHelper/releases/tag/v1.0.2)

Overfiew
--------

This is simple add-in for VBA IDE in 64- and 32-bit host apps (such as Excel, Word, AutoCAD etc), which provides **VBA code indent tool** via Сode Window popup menu and, for additionally - some hotkeys for frequently used operations:

Сomment code lines: <kbd>CTRL</kbd> <kbd>ALT</kbd> <kbd>NUM +</kbd>

UnComment code lines: <kbd>CTRL</kbd> <kbd>ALT</kbd> <kbd>NUM -</kbd>

Toggle bookmark: <kbd>CTRL</kbd> <kbd>ALT</kbd> <kbd>NUM *</kbd>

Go to next bookmark: <kbd>CTRL</kbd> <kbd>`~</kbd>

Go to previous bookmark: <kbd>CTRL</kbd> <kbd>SHIFT</kbd> <kbd>`~</kbd>.

---

Download
--------

Packed binaries for 64- and 32-bit host apps are available at the [Releases][1] page.

>Note that for 64- and 32-bit apps, you must use a corresponding bit DLL.

Install
--------

>Before install, be shure that hotkeys of your background services, OS shell and other add-ins is not same with the add-in hotkeys.

1.  Unpack DLL into local drive folder (*C:\Addons\VBACH*, for example)

2.  Run with administrator credits at Windows command line: 

    For 64-bit host apps on 64-bit OS and for 32-bit host apps on 32-bit OS

        C:\Windows\regsvr32.exe C:\Addons\VBACH\VBACodeHelper.dll

    For 32-bit host apps on 64-bit OS

        C:\Windows\SysWOW64\regsvr32.exe C:\Addons\VBACH\VBACodeHelper.dll

3.  Turn on "Loaded" checkbox in VBE add-in's dialog.

Update
--------

Just replace the addin's dll where it is located.

Deinstall
--------

Use the same command with "/**u**" command line key, for example: 

    regsvr32.exe /u C:\Addons\VBACH\VBACodeHelper.dll
    
Restrictions
---------

Code lines, which concatenated by undescore at the end of phisical line, would be concat into single phisical code line. 
If its undesirable for you, use "Selected Lines" indentation mode where it is necessary.    

Credits 
--------

The mechanics of indenting VBA code is based on the original algorithm of Michael Ciurescu author [VBTools AddIn][2].

[1]: https://github.com/osevoso/VBACodeHelper/releases/
[2]: http://www.vbforums.com/showthread.php?479449-VBTools-AddIn-Auto-indent-VB-code-!
