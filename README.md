<div align="center">

## Dynamically Added Objects


</div>

### Description

This short, simple code allows the addition of objects into a container on another form (in this case a PicturBox). It can easily be modified to add the object to different locations by modifying the locations in the dynObject() function.
 
### More Info
 
Create 2 forms. Add a CommandButton to Form1, and add a PictureBox for Form2.

Pressing the first Command1 loads a label on Form2 in the picture box and a command button on Form1. Pressing the new command button will activate a message box.

Pressing Command1 again will cause an error which should be handled in an actual application.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[CJ Schaff](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cj-schaff.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/cj-schaff-dynamically-added-objects__1-13745/archive/master.zip)





### Source Code

```
Dim WithEvents dynbutton As VB.CommandButton
Dim WithEvents dynLabel As VB.Label
Private Sub Form_Load()
 Form2.Show
 Form2.Top = Form1.Top
 Form2.Left = Form1.Left + Form1.Width
End Sub
Private Sub Command1_Click()
 Call dynObjects
End Sub
Public Sub dynObjects()
 'Define label location and properties
   Set dynLabel = Form2.Controls.Add("VB.label", "dynLabel", Form2.Picture1)
    dynLabel.Caption = "Dynamically added label!"
    dynLabel.Visible = True
    dynLabel.BorderStyle = 1
 'Define CommandButton location and properties
   Set dynbutton = Form1.Controls.Add("VB.commandbutton", "dynButton", Form1)
    dynbutton.Caption = "Dynamic Button"
    dynbutton.Visible = True
    dynbutton.Width = 1275
    dynbutton.Font = "MS Sans Serif"
 End Sub
Private Sub dynButton_click()
 MsgBox ("You have pressed a dynamically added button")
End Sub
```

