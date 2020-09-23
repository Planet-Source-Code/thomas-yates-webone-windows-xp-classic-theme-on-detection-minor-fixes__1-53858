<div align="center">

## WebOne Windows XP Classic Theme On Detection  \*Minor Fixes\*


</div>

### Description

*MINOR FIX* NOW YOU CAN DETECT Whether Windows XP is running in classic style or without visual styles. The function called can also be used to tell theme color if themes are on. If themes are not on then the function will fail and this code will trap the error and return False. REQUIRES "themeui.dll" Should be on Windows XP versions because this is the theme system. THIS FILE IS NOT THE ONE USED FOR DRAWING THEMES. Please let me know of any problems. I reply to all emails. Also check out my other cool submissions. THIS IS DONE IN ONLY 14 LINES OF REAL CODE NOT COMMENTS
 
### More Info
 
No Inputs

Look under X:/Windows/System32/ for the Themeui.dll file where X = Windows System Drive

Boolean

May not work in systems that are not XP. Please let me know!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Thomas Yates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thomas-yates.md)
**Level**          |Beginner
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/thomas-yates-webone-windows-xp-classic-theme-on-detection-minor-fixes__1-53858/archive/master.zip)

### API Declarations

None. Uses the dll for calls


### Source Code

```
Dim TheCurrentTheme As New Theme.Theme
Dim Manager As New ThemeManager
Public Function ClassicThemeOn() as Boolean
Dim Testit ' Used to catch if windows is in classic mode
 Set TheCurrentTheme = Manager.SelectedTheme
 'Test to see if windows is in classic style
 On Error Resume Next
 Testit = TheCurrentTheme.VisualStyleColor
 If Err.Number = -2147024894 Then
  'Error number is the number caused
  'when themed.VisualStyleColor Fails
  'when in clasic mode
  ClassicThemeOn = True
 Else
  ClassicThemeOn = False
 End If
 On Error GoTo 0
End Function
```

