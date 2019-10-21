# ExcelAddinBarGenerator
Simple function to generate add-in buttons from your modules

Refer to UsagEexample.vb for example of how to call the function. Inputs to the function are Generate_AddIn_Button("Button name" as string, "Macro name" as string formatted as ModuleName.MacroName, IconID as integer).

Calling this function will add your macro under an "Add-ins" tab at the top of your excel window. The button will be created with your designated name, and icon ID. Duplicates will be deleted, however if you delete one of your functions, you must manually remove the button. There's absolutely a way to automatically remove the button, but don't have that set up -- anyone is  free to push a change to implement that.
