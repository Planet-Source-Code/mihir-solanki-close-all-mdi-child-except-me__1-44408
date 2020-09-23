<div align="center">

## Close all MDI Child except me


</div>

### Description

This code will close all other MDI Child except the current activated one.Simply put this procedure in Module or Form itself..There are two different way to use it.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mihir Solanki](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mihir-solanki.md)
**Level**          |Advanced
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mihir-solanki-close-all-mdi-child-except-me__1-44408/archive/master.zip)





### Source Code

```
Option Explicit
'+++++++++++++++++++++++++++++++++++++
' First Style
' Use private procedure in Form
'+++++++++++++++++++++++++++++++++++++
Private Sub Form_Activate()
  UnloadOthers
End Sub
Private Sub UnloadOthers()
  Dim frm As Form
  For Each frm In Forms
    If frm.Name <> Me.Name And Not (TypeOf frm Is MDIForm) Then
      Unload frm
    End If
  Next
End Sub
'+++++++++++++++++++++++++++++++++++++
' Second Style
' Use Public Procedure in Module
'+++++++++++++++++++++++++++++++++++++
'Form Code
Private Sub Form_Activate()
  UnloadOthers me.Name
End Sub
'Module Code
Public Sub UnloadOthers(frmName as string)
  Dim frm As Form
  For Each frm In Forms
    If frm.Name <> frmName And Not (TypeOf frm Is MDIForm) Then
      Unload frm
    End If
  Next
End Sub
```

