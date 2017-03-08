VBA Code Helper
===============
Overfiew
--------

This is simple add-in for VBA IDE in 32- and 64-bits host applications (such as Excel, Word, AutoCAD...), which provides **VBA code indent tool** via Сode Window context menu items, and for additionally, two hotkey:

>Ctrl + Alt + *"Gray Plus key"*

for comment code lines,

>Ctrl + Alt + *"Gray Minus key"*

for uncomment code lines.

---

Download
--------

Archived binaries for 32- and 64-bites host applications are available at the [Releases][1] page.

>Note that for 64 and 32-bits host applications must use relevant DLL of bits.

Install
--------

>**Look out!** Before install, be shure what hotkeys of your background services and OS shell is not same with the add-in hotkeys to avoid undesirable surprises.

1.  Unpack DLL into preffered folder (*C:\Addons\VBACH*, for example)

2.  Run with administrator credits at Windows command line: 

    For 64-bits host applications on 64-bits OS and for 32-bits host applications on 32-bits OS

        C:\Windows\regsvr32.exe C:\Addons\VBACH\VBACodeHelper.dll

    For 32-bits host applications on 64-bits OS

        C:\Windows\SysWOW64\regsvr32.exe C:\Addons\VBACH\VBACodeHelper.dll

Deinstall
--------

Use the same command with "/**u**" command line key, for example: 

    regsvr32.exe /u C:\Addons\VBACH\VBACodeHelper.dll

Restrictions
--------

Code lines, which concatenated by undescore at the end of phisical line, would be concat into single phisical code line. If its undesirable for you, use "Selected Lines" indent mode for skip their where necessary.

Credits 
--------

The mechanics of indenting VBA code is based on the original algorithm of Michael Ciurescu author [VBTools AddIn][2].

[1]: https://github.com/osevoso/VBACodeHelper/releases/
[2]: http://www.vbforums.com/showthread.php?479449-VBTools-AddIn-Auto-indent-VB-code-!
