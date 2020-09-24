<div align="center">

## Formatted MsgBoxes


</div>

### Description

I got really tired of having MsgBoxes that reached across the screen to deliver one or two lines of text, so I decided to write a function that would keep them about the same width. The function will add line breaks to keep the length of each line about 70, with variation according to content. I tried to keep the context of the input about the same by only breaking at spaces, dashes, and underscores. The code's a little sloppy, in my opinion, but sadly I'm not even sure how to clean it up.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[unsticky](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/unsticky.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/unsticky-formatted-msgboxes__1-62252/archive/master.zip)





### Source Code

```
Public Function FormattedMsg(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "MsgBox") As VbMsgBoxResult
If Len(Prompt) > 80 Then
 Dim L As Long, R As Long, LChr As String, RChr As String, Break As Long, tmp(1) As String
 tmp(0) = Prompt
 Prompt = ""
 Do Until tmp(0) = ""
 L = 69
 R = 70
 LChr = Mid(tmp(0), L, 1)
 RChr = Mid(tmp(0), R, 1)
 Do Until LChr = " " Or LChr = "-" Or LChr = "_" Or L = 1
  L = L - 1
  LChr = Mid(tmp(0), L, 1)
 Loop
 Do Until RChr = " " Or RChr = "-" Or RChr = "_" Or R = Len(tmp(0))
  R = R + 1
  RChr = Mid(tmp(0), R, 1)
 Loop
 Break = IIf(70 - L < R - 70, L, R)
 tmp(1) = Left(tmp(0), Break) & vbCrLf
 tmp(1) = IIf(Left(tmp(1), 1) <> " ", tmp(1), Mid(tmp(1), 2))
 tmp(0) = Mid(tmp(0), Break)
 If Len(tmp(0)) < 76 Then
  tmp(1) = tmp(1) & IIf(Left(tmp(0), 1) <> " ", tmp(0), Mid(tmp(0), 2))
  tmp(0) = ""
 End If
 Prompt = Prompt & tmp(1)
 Loop
End If
FormattedMsg = MsgBox(Prompt, Buttons, Title)
End Function
```

