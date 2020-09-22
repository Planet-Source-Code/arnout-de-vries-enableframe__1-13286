<div align="center">

## enableFrame


</div>

### Description

When you disable a Frame then all controls in it are disabled to, nice feature, but to the users it still looks like if the controls as enabled.

So I wrote a little subroutine which en/disables all controls in a frame. Can be handy sometimes :-)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Arnout de Vries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/arnout-de-vries.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/arnout-de-vries-enableframe__1-13286/archive/master.zip)





### Source Code

```
Option Explicit
Public Sub enableFrame(curFrame As Frame)
 ' purpose:
 '  set the .enabled property of all controls on a frame to
 '  the same state as the enabled state of the current frame
 Dim ctl As Control
 ' Loop through all controls on the current form
 For Each ctl In curFrame.Parent.Controls
  On Error Resume Next        ' error checking, because not every control has
                    ' a container property
  If ctl.Container.hWnd = curFrame.hWnd Then
   If Err.Number = 0 Then      ' if we didn't receive an error code, proceed
    ctl.Enabled = curFrame.Enabled ' state of control same as Frame
    If TypeOf ctl Is Frame Then   ' if the control is a frame itself then
     enableFrame ctl        ' enter this procedure again for the current frame
    End If
   End If
  End If
 Next ctl
End Sub
```

