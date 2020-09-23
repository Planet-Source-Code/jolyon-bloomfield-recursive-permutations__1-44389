<div align="center">

## Recursive permutations


</div>

### Description

Takes in a string and spits out all possible permutations of the inputted characters using a simple recursive routine. Good recursive example.
 
### More Info
 
Put the lot onto a form, put a command button "command1" on the form, put a textbox "text1" on the form, and run.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jolyon Bloomfield](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jolyon-bloomfield.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jolyon-bloomfield-recursive-permutations__1-44389/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
Open "C:\windows\desktop\words.txt" For Output As #1
Recurse Text1.Text, ""   ' string so permutate is text1.text
Close #1
Shell "C:\windows\notepad.exe C:\windows\desktop\words.txt", vbNormalFocus
End Sub
Private Sub Recurse(ByVal Letters As String, ByVal Built As String)
Dim I As Integer
If Len(Letters) = 1 Then
Print #1, Built & Letters
Exit Sub
End If
For I = 1 To Len(Letters)
Recurse Mid(Letters, 1, I - 1) & Mid(Letters, I + 1), Built & Mid(Letters, I, 1)
Next I
End Sub
```

