<div align="center">

## Activate the Screensaver


</div>

### Description

This lesson will show you how to launch the screensaver with code. It's very simple and only takes a couple of lines of code to accomplish. The first part goes in a module.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Faisal  Arif](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/faisal-arif.md)
**Level**          |Beginner
**User Rating**    |2.6 (13 globes from 5 users)
**Compatibility**  |VB 5\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/faisal-arif-activate-the-screensaver__1-21479/archive/master.zip)





### Source Code

```

(General) (Declarations)
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
 (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As _
 Long, ByVal lParam As Long) As Long
Public Const WM_SYSCOMMAND = &H112&
Public Const SC_SCREENSAVE = &HF140&
To actually activate the screensaver only takes one line of code. You can put it anywhere you want, but for my example, I'm placing it in the Click event of a command button.
Command1 Click
Private Sub Command1_Click()
 Call SendMessage(Me.hWnd, WM_SYSCOMMAND, SC_SCREENSAVE, _
  0&)
End Sub
```

