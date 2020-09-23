<div align="center">

## Default Browser


</div>

### Description

A simple API call to open up the default browser to the url of your choice. VERY simple and useful code.
 
### More Info
 
URL string, hwnd


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[FluffyDave](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/fluffydave.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/fluffydave-default-browser__1-11517/archive/master.zip)

### API Declarations

```
'API for OpenBrowser
Public Const SW_SHOWDEFAULT = 10
Declare Function ShellExecute Lib "shell32.dll" Alias _
 "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
 ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
```


### Source Code

```
Public Function OpenBrowser(strURL As String, lngHwnd As Long)
 OpenBrowser = ShellExecute(lngHwnd, vbNullString, strURL, vbNullString, _
  "c:\", SW_SHOWDEFAULT)
End Function
```

