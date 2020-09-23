<div align="center">

## TranslateColor


</div>

### Description

Windows works with normal colors and system colors. Visual basic can not handle the system colors and work with them as normal RGB colors. Here's an interesting API I found that translates System Colors to 'normal' colors. I made a little prog arround this to show how it is done. WARNING ! I work with win98 so I don't know if it works with WIN95. If it doesn't, could you please tell me. My adress is: stephan.swertvaegher@planetinternet.be    The .dll that contains this API is olepro32.dll. Download the zip and find out...
 
### More Info
 
System Color

Normal RGB color


<span>             |<span>
---                |---
**Submitted On**   |2000-08-20 23:06:26
**By**             |[stephane swertvaegher](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephane-swertvaegher.md)
**Level**          |Intermediate
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD91528202000\.zip](https://github.com/Planet-Source-Code/stephane-swertvaegher-translatecolor__1-10852/archive/master.zip)

### API Declarations

```
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
```





