<div align="center">

## A Better Multiple Undo


</div>

### Description

This code adds a multiple undo/redo function to any textbox or RichTextBox. Easy to set up and use, and doesn't require any extra controls or use of the API. Simple and effective.
 
### More Info
 
A textbox (Text1) and two buttons (cmdUndo and cmdRedo).

No side-effects.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Taras Young](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/taras-young.md)
**Level**          |Advanced
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/taras-young-a-better-multiple-undo__1-22456/archive/master.zip)





### Source Code

```
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' A BETTER MULTIPLE UNDO
''' Copyright (C) 2001 Taras Young
''' http://www.snowblind.net/
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''
''' Paste this code into a form, and add a Textbox (Text1) and
''' two command buttons (cmdUndo and cmdRedo).
'''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''
''' If you want to use a RichTextBox, uncomment the lines
''' marked "for richtextboxes" and comment out the lines
''' marked "for normal textboxes" (obviously).
'''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim UndoStack() As String, UndoStage, Undoing
Private Sub cmdRedo_Click()
Undoing = True
 UndoStage = UndoStage + 1
 Text1.Text = UndoStack(UndoStage)      'for normal textboxes
' Text1.rtfText = UndoStack(UndoStage)    'for richtextboxes
Undoing = False
End Sub
Private Sub cmdUndo_Click()
Undoing = True               'prevent doubling-up
 UndoStage = UndoStage - 1         'go back a stage
 If UndoStage <= 0 Then UndoStage = 0    'protection from errors
'For normal textboxes, use:
 Text1.Text = UndoStack(UndoStage)     'replace current text with
                      'new text
''For richtextboxes, use:
' Text1.rtfText = UndoStack(UndoStage)   'replace current text with
'                      'new text
Undoing = False
End Sub
Private Sub Form_Load()
ReDim UndoStack(0)       'must be redimmed for UBound to work
End Sub
Private Sub Text1_Change()
' Records the last changes made
ReDim Preserve UndoStack(UBound(UndoStack) + 1) 'increase the stack size
'For normal textboxes:
UndoStack(UBound(UndoStack)) = Text1.Text    'add the current state
''For richtextboxes:
'UndoStack(UBound(UndoStack)) = rtfText1.Text  'add the current state
If Not Undoing Then UndoStage = UndoStage + 1  'change the current stage
End Sub
```

