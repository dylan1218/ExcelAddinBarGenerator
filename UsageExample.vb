Sub Auto_Open() 'Or Private Sub Workboo_Open() in ThisWorkbook

'Addin bar will be generated with buttons for the paramaters entered below. These buttons will then call the functions entered into the second paramater
of the Generate_AddIn_Button function.
    Call Generate_AddIn_Button("Table of Contents", "CustomFunctions.CreateTOC", 21)
    Call Generate_AddIn_Button("Generate Email from Selection", "CustomFunctions.CreateEmail_From_Selection", 24)
    Call Generate_AddIn_Button("Get User information", "CustomFunctions.Get_Email_Title_NoLoop", 22)
    Call Generate_AddIn_Button("Copy visible range to new tab", "CustomFunctions.copyRng_NewTab", 32)
    Call Generate_AddIn_Button("Create files at user designated path", "CustomFunctions.createFoldersAtPath", 33)
    Call Generate_AddIn_Button("Find references this sheet relys on", "CustomFunctions.sheetsRelyingOn", 35)
    Call Generate_AddIn_Button("Find sheets that depend on this sheet", "CustomFunctions.sheetsDependantOn", 36)
    
End Sub
