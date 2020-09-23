<div align="center">

## Earthquake Shake Effect


</div>

### Description

Makes a form shaking.

&lt;&lt; Earthquake Shake Effect &gt;&gt;
 
### More Info
 
My First Code In PSC


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[jBaron](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jbaron.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Jokes/ Humor](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/jokes-humor__1-40.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jbaron-earthquake-shake-effect__1-62444/archive/master.zip)





### Source Code

```
'Put this code in a module
'*********************** Author ****************************
'Earthquake Shake Effect
'Author: jBaron
'E-mail Address: jbaron@freemail.gr
'Website: www.hackedemonia.tk
'Dated on: 05-09-2005
'For any problems,suggestions please E-mail me.
'***********************************************************
' In a command button place the above code:
'example: << Call Shake(Form1,1,1000,100) >>
'
'             a)  b) c) d)
'a) Here we call a form.
'
'b) We give the State: 1 for shaking Horizontally,2 for shaking Vertically
'  3 for shaking both Horizonally and Vertically.
'
'c) We give a number to say for How Long the form will be shaking.
'
'd) We give a number to say How Big the effect will be.
'Warning:
' # If You give a large number (Howlong) it wiil take a long time
'  for the effect to pass out and may conflict your program.
' # Same happens with (Howbig).
Public Function Shake(Aform As Form, State As Integer, Howlong As Long, Howbig As Long)
Dim i As Long
If State = 1 Then
  '--> Shakes Form LeftRight...
  For i = 0 To Howlong
    Aform.Left = Aform.Left + Howbig
    Aform.Left = Aform.Left - Howbig
  Next i
End If
If State = 2 Then
  '--> Shakes Form UpDown...
  For i = 0 To Howlong
    Aform.Top = Aform.Top + Howbig
    Aform.Top = Aform.Top - Howbig
  Next i
End If
If State = 3 Then
  '--> Shakes Form UpDown and LeftRight...
  For i = 0 To Howlong
    Aform.Left = Aform.Left + Howbig
    Aform.Left = Aform.Left - Howbig
    Aform.Top = Aform.Top + Howbig
    Aform.Top = Aform.Top - Howbig
  Next i
End If
End Function
```

