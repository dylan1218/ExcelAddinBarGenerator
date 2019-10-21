Public CmdBarMenu As CommandBarControl

Function Generate_AddIn_Button(Button_Caption As String, Sub_Call As String, IDValue)

Dim CmdBar As CommandBar _
, CmdBarMenuItem_Summary_Page As CommandBarControl


Set CmdBar = Application.CommandBars("Worksheet Menu Bar")


Set CmdBarMenu = CmdBar.Controls("Tools")   ' Index 6
Set CmdBarMenuItem_Summary_Page = CmdBarMenu.Controls.Add(Type:=msoControlButton)

'Deletes duplicates
Application.DisplayAlerts = False
    With CmdBarMenu
        On Error Resume Next
        .Controls(Button_Caption).Delete
    End With
Application.DisplayAlerts = True


'Sets captions and assigns a Macro
With CmdBarMenuItem_Summary_Page
     .Caption = Button_Caption
     .OnAction = Sub_Call    'Name of the sub it will call when button is pushed
     .FaceId = IDValue
End With


End Function
