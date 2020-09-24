<div align="center">

## 16 and 32 bit functions to create


</div>

### Description

16 AND 32 bit functions to read/write ini files--very useful!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Qaid](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-qaid.md)
**Level**          |Unknown
**User Rating**    |4.4 (109 globes from 25 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-qaid-16-and-32-bit-functions-to-create__1-174/archive/master.zip)

### API Declarations

```
'****************************************************
'* INI_sm.BAS                   *
'****************************************************
Option Explicit
#If Win16 Then
    Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer
    Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
#Else
' NOTE: The lpKeyName argument for GetProfileString, WriteProfileString,
'    GetPrivateProfileString, and WritePrivateProfileString can be either
'    a string or NULL. This is why the argument is defined as "As Any".
'     For example, to pass a string specify  ByVal "wallpaper"
'     To pass NULL specify          ByVal 0&
'    You can also pass NULL for the lpString argument for WriteProfileString
'    and WritePrivateProfileString
' Below it has been changed to a string due to the ability to use vbNullString
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If
```


### Source Code

```
Create a new module called: INI_SM.BAS
Add an attribute:
Attribute VB_Name = "ini_sm"
Add this code:
'*******************************************************
'* Procedure Name: sReadINI              *
'*=====================================================*
'*Returns a string from an INI file. To use, call the *
'*functions and pass it the Section, KeyName and INI  *
'*File Name, [sRet=sReadINI(Section,Key1,INIFile)].  *
'*val command.                     *
'*******************************************************
Function ReadINI(Section, KeyName, filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function
'*******************************************************
'* Procedure Name: WriteINI              *
'*=====================================================*
'*Writes a string to an INI file. To use, call the   *
'*function and pass it the sSection, sKeyName, the New *
'*String and the INI File Name,            *
'*[Ret=WriteINI(Section,Key,String,INIFile)].     *
'*Returns a 1 if there were no errors and       *
'*a 0 if there were errors.              *
'*******************************************************
Function writeini(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function
```

