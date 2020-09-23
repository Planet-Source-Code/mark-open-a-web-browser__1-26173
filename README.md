<div align="center">

## Open a web browser


</div>

### Description

This is a simple way to run internet explorer from within VB and have it surf automatically to the webpage you specify.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mark](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark.md)
**Level**          |Beginner
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mark-open-a-web-browser__1-26173/archive/master.zip)





### Source Code

```
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const conSwNormal = 1
ShellExecute hwnd, "open", "http://webaddress.com", vbNullString, vbNullString, conSwNormal
```

