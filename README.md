<div align="center">

## QueryUnload


</div>

### Description

How to prevent a form/app from unloading , even if you use the taskmanager? Then try this code... (It works for me ;)
 
### More Info
 
Make sure you DO have a way of closing your form/app


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Riaan Aspeling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/riaan-aspeling.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/riaan-aspeling-queryunload__1-1548/archive/master.zip)





### Source Code

```
Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
'To cancel the unload make the cancel = true. Don't do it
'on the vbAppTaskManager one though.
 Dim ans As String
 Select Case UnloadMode
  Case vbFormControlMenu 'Value 0
'This will be called if you select the close from the little icon
'menu on top and left of the form.
   cancel = False
  Case vbFormCode 'Value 1
'This will be called if your code requested a unload
   cancel = False
  Case vbAppWindows 'Value 2
'vbAppWindows is triggered when you shutdown Windows and your app is still
'running. Added by Jim MacDiarmid
   cancel = False
   End
  Case vbAppTaskManager 'Value 3
'You have to allow the taskmanager to close the program, else you get
'that nasty 'App not responding, close anyway' dialog :<
'The clever way arround it would be to restart your program
'This would be used for a password screen!
   cancel = False
   x = Shell(App.Path & "\" & App.EXEName, vbNormalFocus)
   End
  Case vbFormMDIForm 'Value 4
'This code is called from the parent form
   cancel = False
 End Select
End Sub
```

