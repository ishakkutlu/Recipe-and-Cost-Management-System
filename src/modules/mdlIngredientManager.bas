Attribute VB_Name = "mdlIngredientManager"
Option Explicit

Sub RemoveProductIDFromRecipeIndex(productID As String, wsProduct As Worksheet, rangeProduct As Range)
    Dim recipeIDs As String
    Dim recipeArray As Variant
    Dim i As Integer, j As Integer
    Dim wsRecipeIndex As Worksheet
    Dim productList As String
    Dim foundCell As Range
    Dim productIDArray As Variant
    Dim newProductList As String
    Dim recipeName As String, recipeID As String
    
    ' Set worksheet references
    Set wsRecipeIndex = ThisWorkbook.Sheets(3)
    
    ' Get recipe IDs in column I
    recipeIDs = wsProduct.Cells(rangeProduct.Row, 9).value
    
    ' If ProductID has associated recipes
    If recipeIDs <> "" Then
        ' Split multiple Recipe IDs into an array
        recipeArray = Split(recipeIDs, ",")
        
        ' Loop through array to get recipe ID's
        For i = LBound(recipeArray) To UBound(recipeArray)
            recipeID = Trim(recipeArray(i))
            
            ' Search for the RecipeID in column A
            Set foundCell = wsRecipeIndex.Range("A:A").Find(recipeID, LookAt:=xlWhole)
            
            ' If RecipeID is found, get associated ProductIDs
            If Not foundCell Is Nothing Then
                recipeName = foundCell.Offset(0, 1).value ' Get recipe name (column B)
                productList = foundCell.Offset(0, 2).value ' Get product IDs (column C)
                
                ' Convert Product List into an Array
                productIDArray = Split(productList, ",")
                
                ' Create a new product list excluding the removed ProductID
                newProductList = ""
                
                ' Loop through productIDArray and remove the matching productID
                For j = LBound(productIDArray) To UBound(productIDArray)
                    If Trim(productIDArray(j)) <> productID Then
                        ' Append to newProductList if not the productID to remove
                        If newProductList = "" Then
                            newProductList = Trim(productIDArray(j))
                        Else
                            newProductList = newProductList & "," & Trim(productIDArray(j))
                        End If
                    End If
                Next j
                
                ' Update the cell with the new product list
                foundCell.Offset(0, 2).value = newProductList
                
                If newProductList = "" Then ' Recipe does'nt include any product ID
                    Call removeRecipeFile(recipeName, recipeID) ' Delete recipe file
                    wsRecipeIndex.Rows(foundCell.Row).Delete ' Delete recipe index record
                End If
            End If
        Next i
    Else
        ' MsgBox "No recipes associated with this Product ID.", vbInformation, "No Updates Needed"
        Exit Sub
    End If
End Sub

Sub CollectAffectedRecipes(productID As String, wsProduct As Worksheet, rangeProduct As Range)
    Dim recipeIDs As String
    Dim recipeArray As Variant
    Dim i As Integer
    
    ' Get recipe IDs in column I
    recipeIDs = wsProduct.Cells(rangeProduct.Row, 9).value
    
    ' If ProductID has associated recipes
    If recipeIDs <> "" Then
        ' Split multiple Recipe IDs into an array
        recipeArray = Split(recipeIDs, ",")

        ' Loop through array to get recipe ID's
        For i = LBound(recipeArray) To UBound(recipeArray)
            Call FindRecipeDetails(Trim(recipeArray(i)))
        Next i
    Else
        ' MsgBox "Product ID " & productID & " has been successfully updated. No recipe updates are needed.", vbInformation, "Update Successful"
        Exit Sub
    End If
End Sub

Sub FindRecipeDetails(recipeKey As String)
    Dim wsRecipeIndex As Worksheet, wsProduct As Worksheet ', wsOutput As Worksheet
    Dim foundCell As Range, productCell As Range
    Dim productList As String
    Dim productIDArray As Variant
    Dim i As Integer, lastRow As Long, rowIndex As Integer
    Dim totalAmount As Double, productAmount As Double
    Dim amountPercentage As Double
    Dim totalCost As Double, totalFat As Double, totalSugar As Double, totalSalt As Double
    Dim recipeDict As Object
    Dim key As Variant
    Dim recipeName As String, recipeID As String
    
    ' Set worksheet references
    Set wsRecipeIndex = ThisWorkbook.Sheets(3)
    Set wsProduct = ThisWorkbook.Sheets(2)
    ' Set wsOutput = ThisWorkbook.Sheets(4)
    
    ' Create dictionary to store product details
    Set recipeDict = CreateObject("Scripting.Dictionary")
    recipeID = recipeKey
    
    ' Initialize total variables
    totalCost = 0
    totalAmount = 0
    totalFat = 0
    totalSugar = 0
    totalSalt = 0
    rowIndex = 1 ' Start numbering from 1
    
    ' Search for the RecipeID in column A
    Set foundCell = wsRecipeIndex.Range("A:A").Find(recipeID, LookAt:=xlWhole)
    
    ' If RecipeID is found, get associated ProductIDs
    If Not foundCell Is Nothing Then
        recipeName = foundCell.Offset(0, 1).value ' Get recipe name (column B)
        productList = foundCell.Offset(0, 2).value ' Get product IDs (column C)
        
        Call removeRecipeFile(recipeName, recipeID) ' Delete old recipe file
        
        ' Convert Product List into an Array
        productIDArray = Split(productList, ",")
        
        ' Calculate total amount for all related products
        For i = LBound(productIDArray) To UBound(productIDArray)
            Set productCell = wsProduct.Range("A:A").Find(Trim(productIDArray(i)), LookAt:=xlWhole)
            If Not productCell Is Nothing Then
                totalCost = totalCost + productCell.Offset(0, 3).value ' Cost / Price
                totalAmount = totalAmount + productCell.Offset(0, 4).value ' Amount (gr)
                totalFat = totalFat + productCell.Offset(0, 5).value ' Fat (gr)
                totalSugar = totalSugar + productCell.Offset(0, 6).value ' Sugar (gr)
                totalSalt = totalSalt + productCell.Offset(0, 7).value ' Salt (gr)
            End If
        Next i
        
        ' Store product details including calculated amount percentage
        For i = LBound(productIDArray) To UBound(productIDArray)
            Set productCell = wsProduct.Range("A:A").Find(Trim(productIDArray(i)), LookAt:=xlWhole)
            If Not productCell Is Nothing Then
                productAmount = productCell.Offset(0, 4).value ' Amount (gr)
                If totalAmount > 0 Then
                    amountPercentage = (productAmount / totalAmount) * 100 ' Convert to percentage
                Else
                    amountPercentage = 0
                End If
                
                ' Store in dictionary with updated order including Amount (%)
                recipeDict.Add Trim(productIDArray(i)), Array( _
                    rowIndex, _
                    productCell.value, _
                    productCell.Offset(0, 1).value, _
                    productCell.Offset(0, 2).value, _
                    productCell.Offset(0, 3).value, _
                    productCell.Offset(0, 4).value, _
                    amountPercentage, _
                    productCell.Offset(0, 5).value, _
                    productCell.Offset(0, 6).value, _
                    productCell.Offset(0, 7).value)
                    
                    ' No. - rowIndex
                    ' Product ID - productCell.value
                    ' Product Name - productCell.Offset(0, 1).value
                    ' Brand / Supplier - productCell.Offset(0, 2).value
                    ' Cost / Price - productCell.Offset(0, 3).value
                    ' Amount (gr) - productCell.Offset(0, 4).value
                    ' Amount (%) - amountPercentage
                    ' Fat (gr) - productCell.Offset(0, 5).value
                    ' Sugar (gr) - productCell.Offset(0, 6).value
                    ' Salt (gr) - productCell.Offset(0, 7).value
                    
                rowIndex = rowIndex + 1 ' Increment for next row
            End If
        Next i
    
        ' Recreate recipe book
        recipeUpdateBool = True 'Updating record
        recipeUpdateAssistBool = False 'Assist boolean
        Call createRecipeBook(recipeName, recipeID, recipeDict, totalCost, totalAmount, totalFat, totalSugar, totalSalt)
    
'        ' Write dictionary data to Sheet 4 (wsOutput)
'        lastRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row + 1 ' Find the next empty row
'
'        For Each key In recipeDict.Keys
'            wsOutput.Cells(lastRow, 1).value = recipeDict(key)(0) ' Product ID
'            wsOutput.Cells(lastRow, 2).value = recipeDict(key)(1) ' Product Name
'            wsOutput.Cells(lastRow, 3).value = recipeDict(key)(2) ' Brand / Supplier
'            wsOutput.Cells(lastRow, 4).value = recipeDict(key)(3) ' Cost / Price
'            wsOutput.Cells(lastRow, 5).value = recipeDict(key)(4) ' Amount (gr)
'            wsOutput.Cells(lastRow, 6).value = recipeDict(key)(5) ' Amount (%)
'            wsOutput.Cells(lastRow, 7).value = recipeDict(key)(6) ' Fat (gr)
'            wsOutput.Cells(lastRow, 8).value = recipeDict(key)(7) ' Sugar (gr)
'            wsOutput.Cells(lastRow, 9).value = recipeDict(key)(8) ' Salt (gr)
'            lastRow = lastRow + 1
'        Next key
'    Else
'        ' MsgBox "Recipe ID not found in Recipe Index.", vbExclamation, "Error"
    End If
    
    ' Cleanup
    Set wsRecipeIndex = Nothing
    Set wsProduct = Nothing
    Set foundCell = Nothing
    Set productCell = Nothing
    Set recipeDict = Nothing
    ' Set wsOutput = Nothing
End Sub

Sub addBorder(ws As Worksheet, wsRange As Range)
    Dim grayColor As Long
    Dim lastRow As Long
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Ensure there is valid data in the sheet before proceeding
    If lastRow < 2 Then Exit Sub
    
    ' Set wsRange = ws.Range("A" & lastRow & ":I" & lastRow)
    
    ' AutoFit columns for better visibility
    wsRange.Columns(2).AutoFit ' Second column in wsRange
    wsRange.Columns(3).AutoFit ' Thirth column in wsRange
    
    If wsRange.Columns(2).ColumnWidth < 25 Then wsRange.Columns(2).ColumnWidth = 25
    If wsRange.Columns(3).ColumnWidth < 25 Then wsRange.Columns(3).ColumnWidth = 25
    wsRange.Columns(wsRange.Columns.Count).ColumnWidth = 60 ' Last column in wsRange
    
    ' Remove grids
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    
    ' Add borders without diagonal lines
    ' Define gray color (RGB format)
    grayColor = RGB(200, 200, 200) ' Light gray tone
    
    ' Apply borders
    With wsRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = grayColor ' Apply gray color
        ' Remove diagonal lines
        .item(xlDiagonalUp).LineStyle = xlNone
        .item(xlDiagonalDown).LineStyle = xlNone
    End With

    ' Cleanup
    Set wsRange = Nothing
End Sub

Sub removeRecipeFile(recipeName As String, recipeID As String)
    Dim recipePath As String
    Dim recipeFileName As String
    
    ' Define the "Recipes" folder path
    recipePath = ThisWorkbook.Path & "\Recipes\"
    
    ' Create the file name using Recipe ID and Recipe Name
    recipeFileName = recipePath & recipeName & "_" & recipeID & ".xlsx"
    
    ' Check if the file exists and delete it
    If Dir(recipeFileName) <> "" Then Kill recipeFileName
End Sub



