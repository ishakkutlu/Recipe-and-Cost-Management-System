Attribute VB_Name = "mdlRecovery"
Option Explicit

Sub UnprotectAllSheets()
    If ThisWorkbook.Worksheets(2).ProtectContents Then
        ThisWorkbook.Worksheets(2).Unprotect Password:="123"
    End If
    If ThisWorkbook.Worksheets(3).ProtectContents Then
        ThisWorkbook.Worksheets(3).Unprotect Password:="123"
    End If
End Sub

Sub ProtectAllSheets()
    If Not ThisWorkbook.Worksheets(2).ProtectContents Then
        ThisWorkbook.Worksheets(2).Protect Password:="123", DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
    If Not ThisWorkbook.Worksheets(3).ProtectContents Then
        ThisWorkbook.Worksheets(3).Protect Password:="123", DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
End Sub

'Sub ProtectAllSheets()
'    Dim ws As Worksheet
'
'    ' Loop through all worksheets in the workbook
'    For Each ws In ThisWorkbook.Worksheets
'        ' Check if the sheet is already protected
'        If Not ws.ProtectContents Then
'            ws.Protect Password:="123", DrawingObjects:=True, Contents:=True, Scenarios:=True
'        End If
'    Next ws
'End Sub
'
'Sub UnprotectAllSheets()
'    Dim ws As Worksheet
'
'    ' Loop through all worksheets in the workbook
'    For Each ws In ThisWorkbook.Worksheets
'        ' Check if the sheet is already protected
'        If ws.ProtectContents Then
'            ws.Unprotect Password:="123"
'        End If
'    Next ws
'End Sub


Sub recoverRecipeFiles()
    Dim confirmation As Variant
    Dim recipePath As String
    Dim recipeFileName As String
    Dim wsRecipeIndex As Worksheet
    Dim recipeName As String
    Dim recipeID As String
    Dim lastRow As Long, i As Long
    Dim recipeArray As Variant
    Dim recoveredFiles As Long

    ' Display a confirmation message to the user
    confirmation = MsgBox("This process will attempt to recover lost or deleted recipe files inside the 'Recipes' folder." & vbCrLf & vbCrLf & _
                          "If any recipe files were accidentally deleted or lost, they will be recreated." & vbCrLf & vbCrLf & _
                          "Depending on the number of missing files and their content, this operation may take some time." & vbCrLf & vbCrLf & _
                          "Do you want to proceed?", vbYesNo + vbQuestion, "Confirm Recovery")

    ' If the user selects "No", exit the procedure
    If confirmation = vbNo Then Exit Sub
    
    ' Define the "Recipes" folder path where recipe files will be stored
    recipePath = ThisWorkbook.Path & "\Recipes\"

    ' Ensure that the "Recipes" folder exists, if not, create it
    If Dir(recipePath, vbDirectory) = "" Then
        MkDir recipePath
    End If
    
    ' Set the worksheet reference
    Set wsRecipeIndex = ThisWorkbook.Sheets(3)
    
    ' Find the last row in column A
    lastRow = wsRecipeIndex.Cells(wsRecipeIndex.Rows.Count, "A").End(xlUp).Row
    
    ' Ensure there is valid data in the Recipe Page before proceeding
    If lastRow < 2 Then
        MsgBox "No data found in the Recipe sheet!", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Read Recipe ID and Recipe Name columns into an array for optimized processing
    recipeArray = wsRecipeIndex.Range("A2:B" & lastRow).value ' Load data efficiently

    ' Initialize counter for created files
    recoveredFiles = 0
    
    Application.ScreenUpdating = False
    ' Loop through the array
    For i = 1 To UBound(recipeArray, 1)
        ' Extract Recipe ID and Recipe Name
        recipeID = recipeArray(i, 1)  ' Column A (Recipe ID)
        recipeName = recipeArray(i, 2) ' Column B (Recipe Name)

        ' Construct the expected file name using Recipe Name and Recipe ID
        recipeFileName = recipePath & recipeName & "_" & recipeID & ".xlsx"
        
        ' Check if the recipe file does not already exist
        ' If the file is missing, create it by calling "FindRecipeDetails" procedure
        If Dir(recipeFileName) = "" Then
            Call FindRecipeDetails(recipeID)
            recoveredFiles = recoveredFiles + 1 ' Increment the counter
        End If
    Next i
    Application.ScreenUpdating = True
    
    ' Display a success message after completing the recovery process
    MsgBox "Recovery process completed successfully!" & vbCrLf & _
           "Total recipe files recovered: " & recoveredFiles, vbInformation, "Recovery Completed"
           
End Sub

Sub OpenRecipeFolder()
    Dim recipePath As String
    
    recipePath = ThisWorkbook.Path & "\Recipes\"
    
    ' Ensure the Recipes Folder exists
    If Dir(recipePath, vbDirectory) = "" Then
        MkDir recipePath
    End If
    
    ' Open Recipes Folder in File Explorer
    Shell "explorer.exe " & recipePath, vbNormalFocus
End Sub



