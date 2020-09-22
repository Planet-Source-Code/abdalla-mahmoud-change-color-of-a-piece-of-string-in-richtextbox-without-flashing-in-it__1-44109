<div align="center">

## Change Color Of a Piece Of String In RichTextBox Without Flashing In It


</div>

### Description

Hi All .. This Function Solve Problm Of Flashing During Change the Color of a Piece of String in RichTextBox .. It's Useful For CodeBoxes ..
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Abdalla Mahmoud](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/abdalla-mahmoud.md)
**Level**          |Beginner
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/abdalla-mahmoud-change-color-of-a-piece-of-string-in-richtextbox-without-flashing-in-it__1-44109/archive/master.zip)





### Source Code

```
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Function ColorRtfString(ByVal SelStart As String, ByVal SelLength As Long, ByVal Color As Long)
Dim OldPos As Long
Call LockWindowUpdate(RichTextBox.hWnd)
'Locking Editing (It Ignores Flashing )
OldPos = RichTextBox.SelStart
RichTextBox1.SelStart = SelStart
RichTextBox1.SelLength = SelLength
RichTextBox1.SelColor = Color
RichTextBox1.SelStart = OldPos
RichTextBox1.SelLength = 0
'Unlocking Editing
Call LockWindowUpdate(0)
End Function
```

