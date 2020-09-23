<div align="center">

## String Table Resource file viewer


</div>

### Description

Use the Win32 API functions to retreive from a resource library (exe or dll) a string within a String Table with its specified ID and LanguageID.

Gives the ability to write real multilingual VB applications by using a resource file. You can define in the resource file more than one columns for the String Table type and return the string based on the ID and the LanguageID.

This code id bypassing the standard VB function LoadResString where you don't have the ability to specifiy a language.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-08-25 16:35:36
**By**             |[Florian Santi Development](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/florian-santi-development.md)
**Level**          |Intermediate
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[String Tab269409252001\.zip](https://github.com/Planet-Source-Code/florian-santi-development-string-table-resource-file-viewer__1-27535/archive/master.zip)

### API Declarations

```
LoadLibrary
FreeLibrary
RtlMoveMemory
FindResourceEx
LoadResource
LockResource
SizeofResource
FreeResource
```





