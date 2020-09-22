<div align="center">

## multiple arguments


</div>

### Description

Do you want to pass multiple arguments to your function/method?

There are two ways to do this using unbound array or you can use 'ParamArray'. ParamArray is really a professional way to proceed.

Here is the sample code to make sum of numbers to make you clear how to pass multiple argument to a function.
 
### More Info
 
multiple argument

number


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Deepak Kumar Shaw](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/deepak-kumar-shaw.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/deepak-kumar-shaw-multiple-arguments__1-38960/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
  '*** 1ts Method ***
 Dim i(5) As Integer
  i(0) = 1: i(1) = 2: i(2) = 3: i(3) = 4: i(4) = 5: i(5) = 5:
  MsgBox Add1(i)
  '*** 2nd method ***
  MsgBox Add(1, 2, 3, 4, 5, 5)
End Sub
'*** Here we are using an unbound array to pass int type of array like C/C++ ***
Private Function Add1(i() As Integer) As Integer
Dim tt: tt = Now()
  Dim j As Long, sum As Long
  For j = 0 To UBound(i)
    sum = sum + i(j)
  Next j
 Debug.Print Now() - tt
  Add1 = sum
End Function
'*** This Method use ParamArray for the multiple arguments ***
Private Function Add(ParamArray i()) As Long
Dim tt: tt = Now()
 Dim sum As Long: sum = 0
 For j = 0 To UBound(i)
  sum = sum + i(j)
 Next j
 Debug.Print Now() - tt
  Add = sum
End Function
```

