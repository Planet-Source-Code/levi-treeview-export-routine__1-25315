<div align="center">

## TreeView Export Routine


</div>

### Description

Exports a TreeView control (including ALL it's children) into a graphical representation with plain text.
 
### More Info
 
A tree control!

You'll never find anything

'easier. just paste the code into your

'program and call the first function

'and it will return the output of

'everything in your TreeView control

'Ex:Text1.Text = exportTree(MyTreeView)

A string with the formatted

'text to look like the tree control


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Levi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/levi.md)
**Level**          |Beginner
**User Rating**    |3.0 (9 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/levi-treeview-export-routine__1-25315/archive/master.zip)





### Source Code

```
'sample output:
'
'+- 07/21/01 - 11:57:09 PM
'| \- Shell_TrayWnd - ""
'| |- Button - ""
'| |- TrayNotifyWnd - ""
'| | \- TrayClockWClass - ""
'| |- MSTaskSwWClass - ""
'| \- SysTabControl32 - ""
'**************************************
Function exportTree(tree As TreeView) As String
 Dim txtlen As Long, txtbuffer As String, index As Long
 Dim vlines() As Boolean, ret as String
 ret = chr(13) & chr(10)
 ReDim Preserve vlines(0)
 vlines(0) = True
 numchildren = 0
 exportTree = "+- " & Format(Date, "mm/dd/yy") & " - " & Format(Time, "hh:mm:ss AM/PM") & " - "
 Dim depth As Integer, start As Long
 index = Tree.Nodes.Item(1).FirstSibling.index
 Do
 On Error Resume Next
 Err.Clear
 index2 = Tree.Nodes.Item(index).Next.index
 If Err.Number = 91 Then
 On Error GoTo 0
 vlines(0) = False
 exportTree = exportTree & ret & "| \- " & Tree.Nodes.Item(index).Text _
 & getchildren(index, 0, vlines)
 Exit Do
 Else
 On Error GoTo 0
 exportTree = exportTree & ret & "| |- " & Tree.Nodes.Item(index).Text _
 & getchildren(index, 0, vlines)
 index = Tree.Nodes.Item(index).Next.index
 End If
 Loop
 exportTree = exportTree & ret & "|" & ret & "|" & ret & "|" & ret & "|"
End Function
'the following function calls itself over and over for each child and returns with
'ALL of the children and their children, etc. of the current item in the tree
Function getchildren(ByVal index As Long, ByVal childcnt As Integer, ByRef vlines() As Boolean, Optional data As String = "") As String
 Dim children As Integer
 childcnt = childcnt + 1
 ReDim Preserve vlines(childcnt)
 children = Tree.Nodes.Item(index).children
 If children > 1 Then
 vlines(childcnt) = True
 data = data & ret & childspaces(childcnt, vlines) & "|- " & Tree.Nodes.Item(index).Child
 Call getchildren(Tree.Nodes.Item(index).Child.index, childcnt, vlines, data)
 index = Tree.Nodes.Item(index).Child.index
 For i% = 3 To children
 data = data & ret & childspaces(childcnt, vlines) & "|- " & Tree.Nodes.Item(index).Next
 Call getchildren(Tree.Nodes.Item(index).Next.index, childcnt, vlines, data)
 index = Tree.Nodes.Item(index).Next.index
 Next i%
 vlines(childcnt) = False
 data = data & ret & childspaces(childcnt, vlines) & "\- " & Tree.Nodes.Item(index).Next
 Call getchildren(Tree.Nodes.Item(index).Next.index, childcnt, vlines, data)
 ElseIf children = 1 Then
 vlines(childcnt) = False
 data = data & ret & childspaces(childcnt, vlines) & "\- " & Tree.Nodes.Item(index).Child
 Call getchildren(Tree.Nodes.Item(index).Child.index, childcnt, vlines, data)
 End If
 getchildren = data
End Function
'This function is used to insert the correct amount of space from the edge
'to make all the children line up properly
Function childspaces(childcnt As Integer, vlines() As Boolean) As String
 childspaces$ = "| "
 For i% = 1 To childcnt
 childspaces$ = childspaces$ & IIf(vlines(i% - 1), "| ", " ")
 Next i%
End Function
```

