<div align="center">

## Picturebox Scroller


</div>

### Description

This is a simple function which smoothly scrolls text in a picturebox. Unlike other entries it only needs one picturebox. Each time the function is called it adds the text to the bottom (or top) of the picturebox and scrolls the rest of the box (without moving the box itself). Used with a timer control it creates a very versatile scrolling routine. It can easily be modified to scroll anything that has an HDC. The font property on the picturebox controls the look of the text. BitBlt api based on code submitted by MO. I used this code for a phone tracking program than needed to constantly scroll incoming call data in a picturebox.
 
### More Info
 
a PictureBox control, some text and the direction

Simply add a picturebox called picture1 and a timer called timer1 to a form and paste the code into the form's code module. The timer.interval changes the speed of the scrolling (of course).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Paul Wheeler](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-wheeler.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/paul-wheeler-picturebox-scroller__1-11912/archive/master.zip)

### API Declarations

```
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
```


### Source Code

```
Option Explicit
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Sub Form_Load()
 Timer1.Enabled = True
 Timer1.Interval = 100
End Sub
Sub Timer1_Timer()
 Static i As Integer
 i = i + 1
 If i < 10 Then
 ScrollText Picture1, "Just a simple test #" & i, True
 Else
 ScrollText Picture1, "", True
 End If
End Sub
Sub ScrollText(pic As PictureBox, txt As String, up As Boolean)
 Dim ret As Long, vHeight As Long
 If pic.ScaleMode <> 3 Then pic.ScaleMode = 3
 vHeight = pic.TextHeight(txt)
 If up Then
 ret = BitBlt(pic.hDC, 0, -vHeight, pic.ScaleWidth, pic.ScaleHeight, pic.hDC, 0, 0, &HCC0020)
 pic.Line (0, pic.ScaleHeight - vHeight)-(pic.ScaleWidth, pic.ScaleHeight), pic.BackColor, BF
 pic.CurrentY = pic.ScaleHeight - vHeight
 Else 'down
 ret = BitBlt(pic.hDC, 0, vHeight, pic.ScaleWidth, pic.ScaleHeight, pic.hDC, 0, 0, &HCC0020)
 pic.Line (0, 0)-(pic.ScaleWidth, vHeight), pic.BackColor, BF
 pic.CurrentY = 0
 End If
 pic.CurrentX = (pic.ScaleWidth - pic.TextWidth(txt)) / 2 'centers text
 pic.Print txt
End Sub
```

