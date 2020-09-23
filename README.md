<div align="center">

## HTMLHelpLauncher


</div>

### Description

Allows you to show the new compiled HTML help files (.chm) within a vb application.
 
### More Info
 
The filename of the compiled HTML help file.

The window handler of the help file.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Neil Ault](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/neil-ault.md)
**Level**          |Intermediate
**User Rating**    |4.4 (44 globes from 10 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/neil-ault-htmlhelplauncher__1-9679/archive/master.zip)





### Source Code

```
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~ SUBJECT:   HTML Help Launcher
'~~~ AUTHOR:   Neil Ault (Neil.Ault@btinternet.com)
'~~~ CREATED:   11/07/2000
'~~~
'~~~ DESCRIPTION: Allows you to launch the new compiled HTML help
'~~~       files (.chm) within your visual basic apps. You
'~~~       need to have the file hhctrl.ocx installed on
'~~~       your machine which normally comes with Internet
'~~~       Explorer.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Option Explicit
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
'Constants used by HtmlHelp
Const HH_DISPLAY_TOPIC = &H0
Const HH_SET_WIN_TYPE = &H4
Const HH_GET_WIN_TYPE = &H5
Const HH_GET_WIN_HANDLE = &H6
Const HH_DISPLAY_TEXT_POPUP = &HE   'Display string resource ID or text in a pop-up window.
Const HH_HELP_CONTEXT = &HF      'Display mapped numeric value in dwData.
Const HH_TP_HELP_CONTEXTMENU = &H10  'Text pop-up help, similar to WinHelp's HELP_CONTEXTMENU.
Const HH_TP_HELP_WM_HELP = &H11    'Text pop-up help, similar to WinHelp 's HELP_WM_HELP.
'Opens the compiled help file
Private Sub ShowHelpFile(strFilename As String)
Dim hwndHelp As Long
  'The return value is the window handle of the created help window.
  hwndHelp = HtmlHelp(hWnd, strFilename, HH_DISPLAY_TOPIC, 0)
End Sub
```

