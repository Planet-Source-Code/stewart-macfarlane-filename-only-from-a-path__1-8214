<div align="center">

## Filename only from a path


</div>

### Description

Ever wanted to extract a filename only from a path including the filename, well this simple function will do it for you! (I know its not ground breaking but its simple easy and useful)
 
### More Info
 
call the function like this:

get_filename_only(filepath$) where filepath$ is a variable containing a filepat and filename

or

get_filename_only("c:\windows\notepad.exe") change the path to suit your needs

put this code into a module (.bas) and call from the program as explained above

the filename only (incl. extension) or a message telling you to check the path if it cant find a valid path in the string you sent to it.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Stewart MacFarlane](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stewart-macfarlane.md)
**Level**          |Beginner
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stewart-macfarlane-filename-only-from-a-path__1-8214/archive/master.zip)





### Source Code

```
Function get_filename_only(filepath)
For x = Len(filepath) To 1 Step -1
  If Mid(filepath, x, 1) = "\" Then
    get_filename_only = Right(filepath, Len(filepath) - x)
    Exit Function
  End If
Next x
get_filename_only = "Please check filepath it may be incorrect)"
End Function
```

