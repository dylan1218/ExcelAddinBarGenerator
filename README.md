# ExcelAddinBarGenerator
Simple function to generate add-in buttons from your modules

Usage example:

Sub Auto_Open()

For Each Macro you want a button for:
Call Generate_AddIn_Button("Name for your button" formatted as string, "ModuleName.FunctionName" formatted as string, Icon ID fomatted as integer)

End Sub

Calling this function will add your macro under an "Add-ins" tab at the top of your excel window. The button will be created with your designated name, and icon ID. Duplicates will be deleted, however if you delete one of your functions, you must manually remove the button. There's absolutely a way to automatically remove the button, but don't have that set up -- anyone is  free to push a change to implement that.
