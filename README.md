<div align="center">

## Column Name Excel


</div>

### Description

Returns the name of the column, given a number.

Eg. 5 -&gt; 'E', 50 -&gt; 'AX'. One line of code!
 
### More Info
 
Number of the column

Name of the column


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\[Pieter\]](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pieter.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VBA MS Excel
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pieter-column-name-excel__1-63730/archive/master.zip)





### Source Code

```
Function ReturnName(ByVal num As Integer) As String
 ReturnName = Split(Cells(, num).Address, "$")(1)
End Function
```

