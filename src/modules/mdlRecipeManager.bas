Attribute VB_Name = "mdlRecipeManager"
Option Explicit

' If you want to update the ingredient with the given product ID associated with a recipe, set "recipeUpdateBool" to True.
' If you want to add a new recipe, set "recipeUpdateBool" to False.
Public recipeUpdateBool As Boolean

' Always set "recipeUpdateAssistBool" to True in btnUpdateRecipe_Click when updating a recipe.
' After STEP 1, its value affects the process.
Public recipeUpdateAssistBool As Boolean ' This Boolean variable is False by default.

Public recipePathValidBool As Boolean ' Validate Recipes folder

Sub createRecipeBook(strRecipeName As String, strRecipeID As String, recipeDict As Object, _
                    dblTotalCost As Double, dblTotalAmount As Double, dblTotalFat As Double, dblTotalSugar As Double, dblTotalSalt As Double)
    Dim recipePath As String
    Dim recipeFileName As String
    
    Dim i As Integer
    Dim key As Variant
    
    Dim wsProduct As Worksheet
    Dim rangeProduct As Range
    Dim productID As String
    
    Dim wsRecipeIndex As Worksheet
    Dim indexRange As Range
    Dim indexLastRow As Long
    
    Dim wbRecipe As Workbook
    Dim wsRecipe As Worksheet
    Dim recipeLastRow As Integer
    
    Set wsProduct = ThisWorkbook.Sheets(2)  ' Ingredients (Products) Database Page
    Set wsRecipeIndex = ThisWorkbook.Sheets(3) ' Recipe Index Page
    
    ' ###################################### STEP 1
    If recipeUpdateBool = False Then 'Adding new record
        'Write index ( a new recipe ID)
        indexLastRow = wsRecipeIndex.Cells(wsRecipeIndex.Rows.Count, 1).End(xlUp).Row + 1
        wsRecipeIndex.Cells(indexLastRow, 1).value = strRecipeID
        wsRecipeIndex.Cells(indexLastRow, 2).value = strRecipeName
    End If
    ' ###################################### STEP 1
    
    ' Define the "Recipes" folder path
    recipePath = ThisWorkbook.Path & "\Recipes\"
    
    ' Create the file name using Recipe ID and Recipe Name
    recipeFileName = recipePath & strRecipeName & "_" & strRecipeID & ".xlsx"

    ' Create a new workbook
    Set wbRecipe = Workbooks.Add
    Set wsRecipe = wbRecipe.Sheets(1)
    wsRecipe.Name = "Recipe Page"

    ' Write Recipe Info
    wsRecipe.Cells(2, 2).value = "Recipe ID:"
    wsRecipe.Cells(2, 4).value = strRecipeID
    wsRecipe.Cells(3, 3).value = "Recipe Name:"
    wsRecipe.Cells(3, 4).value = strRecipeName
    
    ' Write Column Headers for Ingredients
    wsRecipe.Cells(5, 2).value = "No."
    wsRecipe.Cells(5, 3).value = "Product ID"
    wsRecipe.Cells(5, 4).value = "Product Name"
    wsRecipe.Cells(5, 5).value = "Brand / Supplier"
    wsRecipe.Cells(5, 6).value = "Cost / Price"
    wsRecipe.Cells(5, 7).value = "Amount (gr)"
    wsRecipe.Cells(5, 8).value = "Amount (%)"
    wsRecipe.Cells(5, 9).value = "Fat (gr)"
    wsRecipe.Cells(5, 10).value = "Sugar (gr)"
    wsRecipe.Cells(5, 11).value = "Salt (gr)"

    'Assist boolean
    If recipeUpdateAssistBool = True Then recipeUpdateBool = False ' for STEP 2
    
'    MsgBox "recipeUpdateBool: " & recipeUpdateBool & vbNewLine & vbNewLine & _
'    "recipeUpdateAssistBool: " & recipeUpdateAssistBool
    
    ' Write Ingredient Data from Dictionary
    recipeLastRow = 6 ' Start writing from row 6
    For Each key In recipeDict.Keys
        wsRecipe.Cells(recipeLastRow, 2).value = recipeDict(key)(0) ' No.
        wsRecipe.Cells(recipeLastRow, 3).value = recipeDict(key)(1) ' Product ID
        wsRecipe.Cells(recipeLastRow, 4).value = recipeDict(key)(2) ' Product Name
        wsRecipe.Cells(recipeLastRow, 5).value = recipeDict(key)(3) ' Brand / Supplier
        wsRecipe.Cells(recipeLastRow, 6).value = recipeDict(key)(4) ' Cost / Price
        wsRecipe.Cells(recipeLastRow, 7).value = recipeDict(key)(5) ' Amount gr
        wsRecipe.Cells(recipeLastRow, 8).value = recipeDict(key)(6) ' Amount Percentage
        wsRecipe.Cells(recipeLastRow, 9).value = recipeDict(key)(7) ' Fat
        wsRecipe.Cells(recipeLastRow, 10).value = recipeDict(key)(8) ' Sugar
        wsRecipe.Cells(recipeLastRow, 11).value = recipeDict(key)(9) ' Salt
        
        ' ###################################### STEP 2
        If recipeUpdateBool = False Then 'Adding new record
            ' Update the Data Page with Recipe ID
            productID = recipeDict(key)(1) ' Get Product ID
            ' Search for the Product ID in column A
            Set rangeProduct = wsProduct.Range("A:A").Find(productID, LookAt:=xlWhole)
            ' Search for the Recipe ID in column A
            Set indexRange = wsRecipeIndex.Range("A:A").Find(strRecipeID, LookAt:=xlWhole)
        
            ' UpdateProductRecipeLink(rangeID As Range, ID As String, ws As Worksheet, col As Integer)
            Call UpdateProductRecipeLink(rangeProduct, strRecipeID, wsProduct, 9)
            Call UpdateProductRecipeLink(indexRange, productID, wsRecipeIndex, 3)
        End If
        ' ###################################### STEP 2
        
        recipeLastRow = recipeLastRow + 1
    Next key
    
    ' Write Total Values
    wsRecipe.Cells(recipeLastRow + 1, 5).value = "Total:"
    wsRecipe.Cells(recipeLastRow + 1, 6).value = dblTotalCost ' Total Cost / Price
    wsRecipe.Cells(recipeLastRow + 1, 7).value = dblTotalAmount ' Total Amount gr
    wsRecipe.Cells(recipeLastRow + 1, 8).value = 100 ' Total Amount Percentage
    wsRecipe.Cells(recipeLastRow + 1, 9).value = dblTotalFat ' Total Fat
    wsRecipe.Cells(recipeLastRow + 1, 10).value = dblTotalSugar ' Total Sugar
    wsRecipe.Cells(recipeLastRow + 1, 11).value = dblTotalSalt ' Total Salt

    ' Format recipe page
    Call formatRecipeBook(wsRecipe, recipeLastRow)
    ' Add a chart
    Call recipeBookChart(wsRecipe)
    
    ' Save the workbook
    Application.DisplayAlerts = False
    wbRecipe.SaveAs recipeFileName
    wbRecipe.Close False
    Application.DisplayAlerts = True
    
    ' Cleanup
    Set wbRecipe = Nothing
    Set wsRecipe = Nothing
    Set wsProduct = Nothing
    Set wsRecipeIndex = Nothing
    Set rangeProduct = Nothing
    Set indexRange = Nothing
    Set recipeDict = Nothing
End Sub

Sub UpdateProductRecipeLink(rangeID As Range, ID As String, ws As Worksheet, col As Integer)
    Dim listID As String
    
    ' If Product ID exists, update index
    If Not rangeID Is Nothing Then
        listID = ws.Cells(rangeID.Row, col).value ' Get existing recipes

        ' If empty, add new ID directly
        If listID = "" Then
            ws.Cells(rangeID.Row, col).value = ID
        Else
            ' If not empty, append new ID with comma separation
            ws.Cells(rangeID.Row, col).value = listID & ", " & ID
        End If
    End If
End Sub

Sub formatRecipeBook(wsRecipe As Worksheet, recipeLastRow As Integer)
    Dim grayColor As Long
    Dim recipeRange As Range
    
    ' Format Recipe Info
    wsRecipe.Range("B2:C2").Merge
    wsRecipe.Range("B3:C3").Merge
    wsRecipe.Range("D2:K2").Merge
    wsRecipe.Range("D3:K3").Merge
    wsRecipe.Range("B2:K3").HorizontalAlignment = xlLeft
    wsRecipe.Range("B2:K3").Font.Bold = True
    wsRecipe.Range("D2:K3").NumberFormat = "@" ' Set as Text format
    
    ' Apply formatting to headers (Bold & Left)
    wsRecipe.Range("B5:K5").Font.Bold = True
    wsRecipe.Range("B5:K5").HorizontalAlignment = xlLeft
    
    ' Apply bold formatting to total row
    wsRecipe.Range("E" & recipeLastRow + 1 & ":K" & recipeLastRow + 1).Font.Bold = True

    ' Apply number formatting (Thousand separator + 2-3 decimal places)
    Set recipeRange = wsRecipe.Range("F6:F" & recipeLastRow + 1) ' Column F (Cost / Price)
    recipeRange.NumberFormat = "#,##0.00" ' Two decimal places with thousand separator
    Set recipeRange = wsRecipe.Range("G6:G" & recipeLastRow + 1) ' Column G (Amount gr)
    recipeRange.NumberFormat = "#,##0.000" ' Three decimal places with thousand separator
    Set recipeRange = wsRecipe.Range("H6:H" & recipeLastRow + 1) ' Column H (Amount Percentage)
    recipeRange.NumberFormat = "#,##0.00" ' Three decimal places with thousand separator
    Set recipeRange = wsRecipe.Range("I6:K" & recipeLastRow + 1) ' Columns I to K (Fat, Sugar, Salt)
    recipeRange.NumberFormat = "#,##0.000" ' Three decimal places with thousand separator
    
    ' Apply cell format
    wsRecipe.Range("B6:B" & recipeLastRow + 1).HorizontalAlignment = xlCenter  ' Column B (No)
    wsRecipe.Range("C6:C" & recipeLastRow + 1).HorizontalAlignment = xlLeft ' Column C (Product ID)
    wsRecipe.Range("C6:C" & recipeLastRow + 1).NumberFormat = "@"  ' Set as Text format (Product ID)
    
    ' AutoFit columns for better visibility
    wsRecipe.Columns("B:K").AutoFit
    wsRecipe.Columns("A").ColumnWidth = 2.5
    
    ' Remove grids
    wsRecipe.Activate
    ActiveWindow.DisplayGridlines = False
    
    ' Add borders without diagonal lines
    ' Define gray color (RGB format)
    grayColor = RGB(200, 200, 200) ' Light gray tone
    
    ' Set the range for the header section
    Set recipeRange = wsRecipe.Range("B2:K3")
    
    ' Apply borders to the header section
    With recipeRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin ' Thin line
        .Color = grayColor ' Apply gray color
        ' Remove diagonal lines
        .item(xlDiagonalUp).LineStyle = xlNone
        .item(xlDiagonalDown).LineStyle = xlNone
    End With
    
    ' Set the range for the main data section
    Set recipeRange = wsRecipe.Range("B5:K" & recipeLastRow - 1)
    
    ' Apply borders to the main data section
    With recipeRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin ' Thin line
        .Color = grayColor ' Apply gray color
        ' Remove diagonal lines
        .item(xlDiagonalUp).LineStyle = xlNone
        .item(xlDiagonalDown).LineStyle = xlNone
    End With
    
    ' Set the range for the total row
    Set recipeRange = wsRecipe.Range("E" & recipeLastRow + 1 & ":K" & recipeLastRow + 1)
    
    ' Apply borders to the total row
    With recipeRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin ' Thin line
        .Color = grayColor ' Apply gray color
        ' Remove diagonal lines
        .item(xlDiagonalUp).LineStyle = xlNone
        .item(xlDiagonalDown).LineStyle = xlNone
    End With

    ' Cleanup
    Set recipeRange = Nothing
End Sub

Sub recipeBookChart(wsRecipe As Worksheet)
    Dim wsChart As Worksheet
    Dim chartObj As ChartObject
    Dim rngChartData As Range
    Dim chartLastRow As Integer

    ' Set the worksheet where the chart will be created
    Set wsChart = wsRecipe
    
    ' Determine the last row for chart data
    chartLastRow = wsChart.Cells(wsChart.Rows.Count, 2).End(xlUp).Row
    Set rngChartData = wsChart.Range("D5:D" & chartLastRow & ",H5:H" & chartLastRow) ' Product Name in D, Amount % in H
    
    ' Delete existing charts before creating a new one
    For Each chartObj In wsChart.ChartObjects
        chartObj.Delete
    Next chartObj
    
    ' Create a new chart within wsRecipe
    Set chartObj = wsChart.ChartObjects.Add(650, 15, 426, 226)
    
    With chartObj.Chart
        .SetSourceData Source:=rngChartData
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Amount Distribution (%)"
        
        ' Ensure legend titles are correctly set
        If .SeriesCollection.Count = 1 Then
            .SeriesCollection(1).XValues = wsChart.Range("D6:D" & chartLastRow)
            .SeriesCollection(1).Name = "Products"
        End If
        
        ' Align the title to the left and adjust font size
        With .ChartTitle
            .Font.Size = 12
            .Left = 60
            '.IncludeInLayout = True
        End With
        
        ' Expand the legend and optimize font size
        With .Legend
            .Font.Size = 10
            .Position = xlLegendPositionRight
        End With

        With .PlotArea
            .Left = 30
        End With
    
        ' Configure data labels inside the chart
        With .SeriesCollection(1)
            .HasDataLabels = True
            With .DataLabels
                .Font.Size = 10
                .ShowPercentage = True
                .ShowValue = False
                .Position = xlLabelPositionInsideEnd
            End With
        End With
    End With
    
    ' Ensure the worksheet remains visible
    wsChart.Visible = xlSheetVisible
    
    ' Cleanup objects
    Set chartObj = Nothing
    Set rngChartData = Nothing
End Sub

Sub updateRecipeIDsInProductPage(recipeID As String, rangeProduct As Range, indexRange As Range, wsProduct As Worksheet)
    Dim oldProductList As String
    Dim oldProductIDArray As Variant
    Dim oldRecipeName As String
    Dim oldRecipeList As String
    Dim oldRecipeIDArray As Variant
    Dim newRecipeList As String
    Dim j As Integer, i As Integer
    
    ' If RecipeID is found, get associated ProductIDs
    If Not indexRange Is Nothing Then
        oldRecipeName = indexRange.Offset(0, 1).value ' Get old recipe name (column B)
        oldProductList = indexRange.Offset(0, 2).value ' Get old product IDs (column C)
        
        ' Convert Product List into an Array
        oldProductIDArray = Split(oldProductList, ",")
        
        ' Loop through productIDArray and remove the recipe id's from the product database
        For i = LBound(oldProductIDArray) To UBound(oldProductIDArray)
            Set rangeProduct = wsProduct.Range("A:A").Find(Trim(oldProductIDArray(i)), LookAt:=xlWhole)
            If Not rangeProduct Is Nothing Then
                oldRecipeList = rangeProduct.Offset(0, 8).value ' Get old recipe IDs (column I)
                ' Convert Product List into an Array
                oldRecipeIDArray = Split(oldRecipeList, ",")

                ' Create a new recipe list excluding the removed RecipeID
                newRecipeList = ""
                
                ' Loop through oldRecipeIDArray and remove the matching recipeID
                For j = LBound(oldRecipeIDArray) To UBound(oldRecipeIDArray)
                    If Trim(oldRecipeIDArray(j)) <> recipeID Then
                        ' Append to newRecipeList if not the recipeID to remove
                        If newRecipeList = "" Then
                            newRecipeList = Trim(oldRecipeIDArray(j))
                        Else
                            newRecipeList = newRecipeList & "," & Trim(oldRecipeIDArray(j))
                        End If
                    End If
                Next j
                ' Update the cell with the new recipe list
                rangeProduct.Offset(0, 8).value = newRecipeList
            End If
        Next i
    End If
    
    ' Delete old recipe file
    Call removeRecipeFile(oldRecipeName, recipeID)

    ' Cleanup
    Set rangeProduct = Nothing
End Sub

Sub recipePathValidation()
    Dim recipePath As String

    ' Define the "Recipes" folder path
    recipePath = ThisWorkbook.Path & "\Recipes\"

    ' Check if the "Recipes" folder exists
    recipePathValidBool = True
    If Dir(recipePath, vbDirectory) = "" Then
        recipePathValidBool = False
        MsgBox "'Recipes' folder not found! Please recover it before proceeding.", vbExclamation, "Missing Folder"
        Exit Sub
    End If
End Sub


