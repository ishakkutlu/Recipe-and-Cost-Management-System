VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRecipeManager 
   Caption         =   "Recipe Manager"
   ClientHeight    =   10590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16845
   OleObjectBlob   =   "frmRecipeManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRecipeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim productIDCheckBool As Boolean
Dim productIDValidBool As Boolean
Dim rangeProduct As Range
Dim wsProduct As Worksheet
Dim productID As String

Dim recipeIDCheckBool As Boolean
Dim recipeIDValidBool As Boolean
Dim mandatoryFieldBool As Boolean
Dim wsRecipeIndex As Worksheet
Dim indexRange As Range
Dim recipeID As String

Private Sub productIDCheck() ' Validates and processes the Product ID input
    ' Trim and remove all spaces from the input Product ID
    productID = Replace(Trim(Me.txtProductID.value), " ", "")

    ' Update the textbox value with the cleaned Product ID
    Me.txtProductID.value = productID

    ' Check if Product ID is empty
    productIDCheckBool = True
    If productID = "" Then
        productIDCheckBool = False
        MsgBox "Please enter an Product ID and try again.", vbExclamation, "Missing Input"
        Exit Sub
    End If

    ' Ensure numeric validation before proceeding
    If Not isOnlyNumeric(productID) Then
        productIDCheckBool = False
        MsgBox "Please enter a valid numeric Product ID. Only numeric characters (0-9) are allowed!", vbExclamation, "Invalid Input"
        Me.txtProductID.value = ""
        Exit Sub
    End If
    ' Set the worksheet reference
    Set wsProduct = ThisWorkbook.Sheets(2)
End Sub

Private Sub productIDValid() ' Checks if the Product ID exists in the database
    ' Validate the Product ID first
    Call productIDCheck
    If Not productIDCheckBool Then Exit Sub

    ' Search for the Product ID in column A
    Set rangeProduct = wsProduct.Range("A:A").Find(productID, LookAt:=xlWhole)

    ' Determine whether the Product ID exists
    productIDValidBool = Not rangeProduct Is Nothing
End Sub

Private Sub btnAddProduct_Click()
    Dim item As listItem
    Dim i As Integer
    Dim duplicate As Boolean

    ' Check the Product ID
    Call productIDValid
    If Not productIDCheckBool Then Exit Sub

    ' Display an message if Product ID is not found
    If Not productIDValidBool Then
        MsgBox "Product ID not found!", vbExclamation, "Invalid Operation"
        Me.txtProductID.value = ""
        Exit Sub
    End If
    
    ' Check for duplicate entry in ListView
    duplicate = False
    For i = 1 To Me.lstRecipeItems.ListItems.Count
        ' Compare Product ID with SubItems(1)
        If Me.lstRecipeItems.ListItems(i).SubItems(1) = productID Then
            duplicate = True
            Exit For
        End If
    Next i
    
    If duplicate Then
        MsgBox "This Product ID is already added.", vbExclamation, "Duplicate Entry"
        Me.txtProductID.value = ""
        Exit Sub
    End If

    ' Add product details to ListView
    Set item = Me.lstRecipeItems.ListItems.Add
    item.Text = Me.lstRecipeItems.ListItems.Count ' index number
    item.SubItems(1) = productID ' Product ID
    item.SubItems(2) = wsProduct.Cells(rangeProduct.Row, 2).value ' Product Name
    item.SubItems(3) = wsProduct.Cells(rangeProduct.Row, 3).value ' Brand / Supplier
    item.SubItems(4) = Format(wsProduct.Cells(rangeProduct.Row, 4).value, "#,##0.00") ' Cost
    item.SubItems(5) = Format(wsProduct.Cells(rangeProduct.Row, 5).value, "#,##0.000") ' Amount
    item.SubItems(6) = "0.00" ' Placeholder for Percentage, will be updated later
    item.SubItems(7) = Format(wsProduct.Cells(rangeProduct.Row, 6).value, "#,##0.000") ' Fat
    item.SubItems(8) = Format(wsProduct.Cells(rangeProduct.Row, 7).value, "#,##0.000") ' Sugar
    item.SubItems(9) = Format(wsProduct.Cells(rangeProduct.Row, 8).value, "#,##0.000") ' Salt"
    
    ' Update total values
    Call UpdateTotalValues

    ' Cleanup
    Set wsProduct = Nothing
    Set rangeProduct = Nothing
    
    ' Clear Product ID textbox and selected item after adding
    Me.txtProductID.value = ""
    Me.lstRecipeItems.selectedItem = Nothing
End Sub

Private Sub UpdateTotalValues()
    Dim i As Integer
    Dim totalAmount As Double
    Dim totalCost As Double
    Dim totalFat As Double
    Dim totalSugar As Double
    Dim totalSalt As Double

    ' Reset total values
    totalAmount = 0
    totalCost = 0
    totalFat = 0
    totalSugar = 0
    totalSalt = 0

    ' Loop through ListView items and sum the values
    For i = 1 To Me.lstRecipeItems.ListItems.Count
        totalCost = totalCost + CDbl(Me.lstRecipeItems.ListItems(i).SubItems(4))
        totalAmount = totalAmount + CDbl(Me.lstRecipeItems.ListItems(i).SubItems(5))
        totalFat = totalFat + CDbl(Me.lstRecipeItems.ListItems(i).SubItems(7))
        totalSugar = totalSugar + CDbl(Me.lstRecipeItems.ListItems(i).SubItems(8))
        totalSalt = totalSalt + CDbl(Me.lstRecipeItems.ListItems(i).SubItems(9))
    Next i

    ' Update total labels
    Me.lblTotalCost.Caption = Format(totalCost, "#,##0.00")
    Me.lblTotalAmount.Caption = Format(totalAmount, "#,##0.000")
    Me.lblTotalFat.Caption = Format(totalFat, "#,##0.000")
    Me.lblTotalSugar.Caption = Format(totalSugar, "#,##0.000")
    Me.lblTotalSalt.Caption = Format(totalSalt, "#,##0.000")

    ' Update Percentage for each item based on new totalAmount
    If totalAmount > 0 Then
        For i = 1 To Me.lstRecipeItems.ListItems.Count
            Me.lstRecipeItems.ListItems(i).SubItems(6) = Format((CDbl(Me.lstRecipeItems.ListItems(i).SubItems(5)) / totalAmount) * 100, "#,##0.00")
        Next i
    End If
    
    ' Create a chart
    Call frmRecipeManagerChart
End Sub

Private Sub btnRemoveProduct_Click()
    Dim selectedItem As listItem
    Dim removedProductID As String
    Dim removedProductName As String
    Dim i As Integer
    
    ' Check if an item is selected in ListView
    If Me.lstRecipeItems.selectedItem Is Nothing Then
        MsgBox "Please select a product to remove!", vbExclamation, "No Selection"
        Exit Sub
    End If

    ' Get the selected Product ID & Product Name
    Set selectedItem = Me.lstRecipeItems.selectedItem
    removedProductID = selectedItem.SubItems(1) ' Product ID
    removedProductName = selectedItem.SubItems(2) ' Product Name
    
    ' Remove the selected item
    Me.lstRecipeItems.ListItems.Remove selectedItem.Index

    ' Update sequence numbers after removal
    For i = 1 To Me.lstRecipeItems.ListItems.Count
        Me.lstRecipeItems.ListItems(i).Text = i ' update index number
    Next i
    
    ' Update total values
    Call UpdateTotalValues

    Me.lstRecipeItems.selectedItem = Nothing ' Clear selection
    
End Sub

' Trigger btnRemoveProduct_Click when an item is double-clicked
Private Sub lstRecipeItems_DblClick()
    ' Ensure an item is selected before executing the remove procedure
    If Not Me.lstRecipeItems.selectedItem Is Nothing Then
        Call btnRemoveProduct_Click ' Call the remove button procedure
    End If
End Sub

Private Sub lstRecipeItems_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    ' Clear selection when the user clicks on an empty area of the ListView
    Me.lstRecipeItems.selectedItem = Nothing
End Sub

Private Sub frmRecipeFrame_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Clear selection when the user clicks on an empty area of the ListView
    Me.lstRecipeItems.selectedItem = Nothing
End Sub

Private Sub recipeIDCheck() ' Validates and processes the Recipe ID
    ' Trim and remove all spaces from the Recipe ID
    recipeID = Replace(Trim(Me.txtRecipeID.value), " ", "")

    ' Update the textbox value with the cleaned Recipe ID
    Me.txtRecipeID.value = recipeID

    ' Check if Recipe ID is empty
    recipeIDCheckBool = True
    If recipeID = "" Then
        recipeIDCheckBool = False
        MsgBox "Please enter a Recipe ID and try again.", vbExclamation, "Missing Input"
        Exit Sub
    End If

    ' Ensure numeric validation before proceeding
    If Not isOnlyNumeric(recipeID) Then
        recipeIDCheckBool = False
        MsgBox "Please enter a valid numeric Recipe ID. Only numeric characters (0-9) are allowed!", vbExclamation, "Invalid Input"
        Me.txtRecipeID.value = ""
        Exit Sub
    End If

    ' Set the worksheet reference
    Set wsProduct = ThisWorkbook.Sheets(2)
    Set wsRecipeIndex = ThisWorkbook.Sheets(3) ' Recipe Index is on Sheet3
End Sub

Private Sub recipeIDValid() ' Checks if the Recipe ID exists in the database
    ' Validate the Product ID first
    Call recipeIDCheck
    If Not recipeIDCheckBool Then Exit Sub

    ' Search for the Recipe ID in column A of the Recipe Data sheet
    Set indexRange = wsRecipeIndex.Range("A:A").Find(recipeID, LookAt:=xlWhole)

    ' Determine whether the Recipe ID exists
    recipeIDValidBool = Not indexRange Is Nothing
End Sub

Private Sub btnLoadRecipe_Click() ' Loads data from the database into the form fields
    Dim productIDs As String
    Dim productArray As Variant
    Dim i As Integer
    
    Call UnprotectAllSheets
    
    ' Validate the Recipe ID
    Call recipeIDValid
    If Not recipeIDCheckBool Then GoTo ExitHandler
    
    ' Display an message if Recipe ID is not found
    If Not recipeIDValidBool Then
        MsgBox "Recipe ID not found!", vbExclamation, "Invalid Operation"
        clearUserForm
        GoTo ExitHandler
    End If

    ' Clear List
    Me.lstRecipeItems.ListItems.Clear
    Me.txtRecipeName.value = ""
    Me.txtProductID.value = ""
    
    ' Populate form fields with existing record data
    If Not indexRange Is Nothing Then
        Me.txtRecipeName.value = CStr(wsRecipeIndex.Cells(indexRange.Row, 2).value) ' Recipe Name

        ' Get product IDs in column C
        productIDs = wsRecipeIndex.Cells(indexRange.Row, 3).value
        
        If productIDs <> "" Then
            ' Split multiple Product IDs into an array
            productArray = Split(productIDs, ",")
            
            ' Loop through array to get product ID's
            For i = LBound(productArray) To UBound(productArray) ' Product details
                productID = Trim(productArray(i))
                Me.txtProductID.value = productID
                btnAddProduct_Click 'Call add click
            Next i
        End If
    End If
    ' Cleanup
    Set wsProduct = Nothing
    Set wsRecipeIndex = Nothing
    Set indexRange = Nothing

ExitHandler:
    Call ProtectAllSheets
End Sub

Private Sub btnUpdateRecipe_Click()
    Dim i As Integer
    Dim recipeDict As Object
    Dim itemData As Variant
    
    Call UnprotectAllSheets
    
    ' Recipe path validation
    Call recipePathValidation
    If Not recipePathValidBool Then GoTo ExitHandler
    
    ' Validate the Recipe ID
    Call recipeIDValid
    If Not recipeIDCheckBool Then GoTo ExitHandler

    If Not recipeIDValidBool Then
        MsgBox "Recipe ID does not exist in the database. Please check and try again.", vbExclamation, "Invalid Update Operation"
        Me.txtRecipeID.value = ""
        GoTo ExitHandler
    End If

    ' Trim all spaces
    Call trimAllSpaces

    ' Check if mandatory fields are filled
    Call mandatoryField
    If Not mandatoryFieldBool Then GoTo ExitHandler

    ' Store Ingredient Data in Dictionary
    Set recipeDict = CreateObject("Scripting.Dictionary")
    For i = 1 To Me.lstRecipeItems.ListItems.Count
        itemData = Array( _
            i, _
            CStr(Me.lstRecipeItems.ListItems(i).SubItems(1)), _
            CStr(Me.lstRecipeItems.ListItems(i).SubItems(2)), _
            CStr(Me.lstRecipeItems.ListItems(i).SubItems(3)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(4)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(5)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(6)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(7)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(8)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(9)))

        recipeDict.Add i, itemData
        ' No. - i
        ' Product ID - CStr(Me.lstRecipeItems.ListItems(i).SubItems(1))
        ' Product Name - CStr(Me.lstRecipeItems.ListItems(i).SubItems(2))
        ' Brand / Supplier - CStr(Me.lstRecipeItems.ListItems(i).SubItems(3))
        ' Cost / Price - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(4))
        ' Amount (gr) - Dbl(Me.lstRecipeItems.ListItems(i).SubItems(5))
        ' Amount (%) - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(6))
        ' Fat (gr) - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(7))
        ' Sugar (gr) - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(8))
        ' Salt (gr) - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(9))
    Next i

    ' Check if recipeDict is empty
    If recipeDict.Count = 0 Then
        MsgBox "At least one product must be added to update a recipe. Please add ingredients and try again.", vbExclamation, "Missing Ingredients"
        GoTo ExitHandler
    End If
    
    'Update recipe IDs in data page and delete old recipe file
    Call updateRecipeIDsInProductPage(recipeID, rangeProduct, indexRange, wsProduct)
    
    ' Clear old data
    indexRange.Offset(0, 1).value = "" ' Clear old recipe name (column B)
    indexRange.Offset(0, 2).value = "" ' Clear old product IDs (column C)
    
    ' Add new data
    indexRange.Offset(0, 1).value = CStr(Me.txtRecipeName.value) ' Clear new recipe name (column B)
    
    ' Create recipe book
    Application.ScreenUpdating = False
    recipeUpdateBool = True 'Update recipe
    recipeUpdateAssistBool = True 'Assist boolean
    Call createRecipeBook(CStr(Me.txtRecipeName.value), CStr(Me.txtRecipeID.value), recipeDict, _
                        CDbl(Me.lblTotalCost.Caption), CDbl(Me.lblTotalAmount.Caption), CDbl(Me.lblTotalFat.Caption), _
                        CDbl(Me.lblTotalSugar.Caption), CDbl(Me.lblTotalSalt.Caption))
    Application.ScreenUpdating = True
    
    Set indexRange = Nothing
    Set rangeProduct = Nothing
    Set wsProduct = Nothing
    Set wsRecipeIndex = Nothing
    Set recipeDict = Nothing
    
    ' Confirm to the user
    MsgBox "The recipe file has been updated, and all associated products have been linked successfully.", vbInformation, "Update Successful"
    
    ' Clear form
    clearUserForm

ExitHandler:
    Call ProtectAllSheets
End Sub

Private Sub btnNewRecipe_Click()
    Dim i As Integer
    Dim recipePath As String
    Dim recipeDict As Object
    Dim itemData As Variant
    Dim lastRow As Long
    Dim wsRange As Range
    
    Call UnprotectAllSheets
    
    ' Recipe path validation
    Call recipePathValidation
    If Not recipePathValidBool Then GoTo ExitHandler
    
    ' Validate the Recipe ID
    Call recipeIDValid
    If Not recipeIDCheckBool Then GoTo ExitHandler

    If recipeIDValidBool Then
        MsgBox "The entered Recipe ID already exists in the database. Please use a unique ID to add a new recipe.", vbExclamation, "Duplicate Recipe ID"
        ' Me.txtRecipeID.value = ""
        GoTo ExitHandler
    End If

    ' Trim all spaces
    Call trimAllSpaces

    ' Check if mandatory fields are filled
    Call mandatoryField
    If Not mandatoryFieldBool Then GoTo ExitHandler
  
    ' Store Ingredient Data in Dictionary
    Set recipeDict = CreateObject("Scripting.Dictionary")
    For i = 1 To Me.lstRecipeItems.ListItems.Count
        itemData = Array( _
            i, _
            CStr(Me.lstRecipeItems.ListItems(i).SubItems(1)), _
            CStr(Me.lstRecipeItems.ListItems(i).SubItems(2)), _
            CStr(Me.lstRecipeItems.ListItems(i).SubItems(3)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(4)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(5)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(6)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(7)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(8)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(9)))
        
        recipeDict.Add i, itemData
        ' No. - i
        ' Product ID - CStr(Me.lstRecipeItems.ListItems(i).SubItems(1))
        ' Product Name - CStr(Me.lstRecipeItems.ListItems(i).SubItems(2))
        ' Brand / Supplier - CStr(Me.lstRecipeItems.ListItems(i).SubItems(3))
        ' Cost / Price - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(4))
        ' Amount (gr) - Dbl(Me.lstRecipeItems.ListItems(i).SubItems(5))
        ' Amount (%) - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(6))
        ' Fat (gr) - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(7))
        ' Sugar (gr) - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(8))
        ' Salt (gr) - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(9))
    Next i

    ' Check if recipeDict is empty
    If recipeDict.Count = 0 Then
        MsgBox "At least one product must be added to create a recipe. Please add ingredients and try again.", vbExclamation, "Missing Ingredients"
        GoTo ExitHandler
    End If
    
    ' Create recipe book
    Application.ScreenUpdating = False
    recipeUpdateBool = False 'Adding new recipe
    recipeUpdateAssistBool = True 'Assist boolean
    Call createRecipeBook(CStr(Me.txtRecipeName.value), CStr(Me.txtRecipeID.value), recipeDict, _
                        CDbl(Me.lblTotalCost.Caption), CDbl(Me.lblTotalAmount.Caption), CDbl(Me.lblTotalFat.Caption), _
                        CDbl(Me.lblTotalSugar.Caption), CDbl(Me.lblTotalSalt.Caption))
    Application.ScreenUpdating = True

    ' Add border
    lastRow = wsRecipeIndex.Cells(wsRecipeIndex.Rows.Count, 1).End(xlUp).Row
    Set wsRange = wsRecipeIndex.Range("A" & lastRow & ":C" & lastRow)
    Call addBorder(wsRecipeIndex, wsRange)
    
    Set indexRange = Nothing
    Set rangeProduct = Nothing
    Set wsProduct = Nothing
    Set wsRecipeIndex = Nothing
    Set recipeDict = Nothing
    Set wsRange = Nothing
    
    ' Confirm to the user
    MsgBox "The recipe file has been created, and all associated products have been linked successfully.", vbInformation, "Creation Successful"
    ' Clear form fields after operation
    clearUserForm

ExitHandler:
    Call ProtectAllSheets
End Sub

Private Sub btnDeleteRecipe_Click()
    Dim i As Integer
    Dim recipePath As String
    Dim recipeName As String
    Dim recipeDict As Object
    Dim itemData As Variant
    Dim confirmation As Variant
    Dim lastRow As Long
    Dim wsRange As Range
    
    Call UnprotectAllSheets
    
    ' Recipe path validation
    Call recipePathValidation
    If Not recipePathValidBool Then GoTo ExitHandler
    
    ' Validate the Recipe ID
    Call recipeIDValid
    If Not recipeIDCheckBool Then GoTo ExitHandler

    If Not recipeIDValidBool Then
        MsgBox "Recipe ID does not exist in the database. Please check and try again.", vbExclamation, "Invalid Update Operation"
        Me.txtRecipeID.value = ""
        GoTo ExitHandler
    End If

    ' Trim all spaces
    Call trimAllSpaces

    ' Store Ingredient Data in Dictionary
    Set recipeDict = CreateObject("Scripting.Dictionary")
    For i = 1 To Me.lstRecipeItems.ListItems.Count
        itemData = Array( _
            i, _
            CStr(Me.lstRecipeItems.ListItems(i).SubItems(1)), _
            CStr(Me.lstRecipeItems.ListItems(i).SubItems(2)), _
            CStr(Me.lstRecipeItems.ListItems(i).SubItems(3)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(4)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(5)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(6)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(7)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(8)), _
            CDbl(Me.lstRecipeItems.ListItems(i).SubItems(9)))

        recipeDict.Add i, itemData
        ' No. - i
        ' Product ID - CStr(Me.lstRecipeItems.ListItems(i).SubItems(1))
        ' Product Name - CStr(Me.lstRecipeItems.ListItems(i).SubItems(2))
        ' Brand / Supplier - CStr(Me.lstRecipeItems.ListItems(i).SubItems(3))
        ' Cost / Price - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(4))
        ' Amount (gr) - Dbl(Me.lstRecipeItems.ListItems(i).SubItems(5))
        ' Amount (%) - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(6))
        ' Fat (gr) - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(7))
        ' Sugar (gr) - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(8))
        ' Salt (gr) - CDbl(Me.lstRecipeItems.ListItems(i).SubItems(9))
    Next i

    ' Check if recipeDict or txtRecipeName is empty
    If recipeDict.Count = 0 Or Me.txtRecipeName.value = "" Then
        MsgBox "To delete a recipe, please follow these steps:" & vbNewLine & vbNewLine & _
               "1. Enter the Recipe ID in the designated field." & vbNewLine & _
               "2. Click 'Load Recipe' to load the recipe details." & vbNewLine & _
               "3. Carefully review the ingredients and other details before proceeding with the deletion." & vbNewLine & _
               "Please complete these steps and try again.", vbExclamation, "Missing Ingredients or Recipe Name"
        GoTo ExitHandler
    End If
    
    recipeName = indexRange.Offset(0, 1).value ' Recipe name stored in the Recipe Index page
    ' Confirm deletion from the user
    confirmation = MsgBox("Are you sure you want to delete the recipe with Recipe ID: " & recipeID & " (" & recipeName & ")?", vbYesNo + vbQuestion, "Confirm Deletion")
    If confirmation = vbYes Then
    
        'Update recipe IDs in data page and delete old recipe file
        Call updateRecipeIDsInProductPage(recipeID, rangeProduct, indexRange, wsProduct)
        
        ' Delete the recipe ID row
        wsRecipeIndex.Rows(indexRange.Row).Delete
        
        ' Add border
        lastRow = wsRecipeIndex.Cells(wsRecipeIndex.Rows.Count, 1).End(xlUp).Row
        Set wsRange = wsRecipeIndex.Range("A" & lastRow & ":C" & lastRow)
        Call addBorder(wsRecipeIndex, wsRange)
    
        ' Cleanup
        Set indexRange = Nothing
        Set rangeProduct = Nothing
        Set wsProduct = Nothing
        Set wsRecipeIndex = Nothing
        Set recipeDict = Nothing
        Set wsRange = Nothing
        
        ' Confirm to the user
        MsgBox "Recipe ID: " & recipeID & " (" & recipeName & ") and all associated products links have been deleted successfully!", vbInformation, "Deletion Successful"

        ' Clear form
        clearUserForm
    End If

ExitHandler:
    Call ProtectAllSheets
End Sub

Private Sub btnClearForm_Click()
    clearUserForm
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub clearUserForm()
    Me.lstRecipeItems.ListItems.Clear
    Me.txtRecipeID.value = ""
    Me.txtRecipeName.value = ""
    Me.txtProductID.value = ""
    Call UpdateTotalValues
    Me.lblTotalCost.Caption = ""
    Me.lblTotalAmount.Caption = ""
    Me.lblTotalFat.Caption = ""
    Me.lblTotalSugar.Caption = ""
    Me.lblTotalSalt.Caption = ""
End Sub

Private Sub trimAllSpaces()
    ' Trim all spaces (Text)
    Me.txtRecipeName.value = Trim(Me.txtRecipeName.value)
    ' Remove extra spaces within the text fields and replace them with a single space
    Me.txtRecipeName.value = ReduceSpaces(Me.txtRecipeName.value)
End Sub

Private Sub mandatoryField()
    mandatoryFieldBool = True
    
    ' Ensure the recipe name is provided
    If Me.txtRecipeName.value = "" Then
        mandatoryFieldBool = False
        MsgBox "Please enter recipe name before proceeding.", vbExclamation, "Missing Data"
        Exit Sub
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim colHeader As ColumnHeader

    ' Prevent editing of the first column (Product ID)
    Me.lstRecipeItems.LabelEdit = lvwManual
    
    ' Configure ListView properties
    With Me.lstRecipeItems
        .ColumnHeaders.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .Sorted = True

        ' Add column headers
        Set colHeader = Me.lstRecipeItems.ColumnHeaders.Add(, , "No.", 30)
        Set colHeader = Me.lstRecipeItems.ColumnHeaders.Add(, , "Product ID", 70)
        Set colHeader = Me.lstRecipeItems.ColumnHeaders.Add(, , "Product Name", 150)
        Set colHeader = Me.lstRecipeItems.ColumnHeaders.Add(, , "Brand / Supplier", 140)
        Set colHeader = Me.lstRecipeItems.ColumnHeaders.Add(, , "Cost / Price", 70)
        Set colHeader = Me.lstRecipeItems.ColumnHeaders.Add(, , "Amount (gr)", 70)
        Set colHeader = Me.lstRecipeItems.ColumnHeaders.Add(, , "Amount (%)", 70)
        Set colHeader = Me.lstRecipeItems.ColumnHeaders.Add(, , "Fat (gr)", 50)
        Set colHeader = Me.lstRecipeItems.ColumnHeaders.Add(, , "Sugar (gr)", 60)
        Set colHeader = Me.lstRecipeItems.ColumnHeaders.Add(, , "Salt (gr)", 50)
    End With
End Sub

Private Sub frmRecipeManagerChart()
    Dim i As Integer
    Dim wsChart As Worksheet
    Dim chartObj As ChartObject
    Dim chartFilePath As String
    Dim rngChartData As Range
    Dim chartLastRow As Integer

    ' Update or create Pie Chart
    On Error Resume Next
    Set wsChart = ThisWorkbook.Sheets("ChartData")
    On Error GoTo 0
    
    ' If the sheet does not exist, create it and move it to the last position
    If wsChart Is Nothing Then
        Set wsChart = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsChart.Name = "ChartData"
    End If
    
    ' Clear worksheet (Remove previous data)
    wsChart.Cells.Clear
    
    ' Add headers
    wsChart.Cells(1, 1).value = "Product Name"
    wsChart.Cells(1, 2).value = "Amount %"
    
    ' Transfer ListView data to ChartData sheet
    For i = 1 To Me.lstRecipeItems.ListItems.Count
        wsChart.Cells(i + 1, 1).value = Me.lstRecipeItems.ListItems(i).SubItems(2) ' Product Name
        wsChart.Cells(i + 1, 2).value = CDbl(Me.lstRecipeItems.ListItems(i).SubItems(6)) ' Amount %
    Next i
    
    Me.imgChart.Picture = Nothing ' Remove chart on the form
    If wsChart.Cells(2, 1).value = "" Then Exit Sub
    
    ' Create or update the chart
    chartLastRow = wsChart.Cells(wsChart.Rows.Count, 1).End(xlUp).Row
    Set rngChartData = wsChart.Range("A1:B" & chartLastRow)
    
    ' Delete existing charts before creating a new one
    For Each chartObj In wsChart.ChartObjects
        chartObj.Delete
    Next chartObj
    
    ' Dimension of the chart
    Set chartObj = wsChart.ChartObjects.Add(150, 15, 350, 150) '350, 150)
    
    With chartObj.Chart
        .SetSourceData Source:=rngChartData
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Amount Distribution (%)"
    
        ' Ensure legend titles are correctly set on the first addition
        If .SeriesCollection.Count = 1 Then
            .SeriesCollection(1).XValues = wsChart.Range("A2:A" & chartLastRow) ' Manually set X-axis values
            .SeriesCollection(1).Name = "Products" ' Manually set legend content to "Products"
        End If
    
        ' Align the title to the left and adjust font size
        With .ChartTitle
            .Font.Size = 10
            .Left = 100
        End With
    
        ' Expand the legend and optimize font size
        With .Legend
            .Font.Size = 8
            '.AutoScaleFont = True '
            .Position = xlLegendPositionRight
        End With

        With .PlotArea
            .Left = 100
        End With
        
        ' Configure data labels inside the chart
        With .SeriesCollection(1)
            .HasDataLabels = True
            With .DataLabels
                .Font.Size = 9
                .ShowPercentage = True  ' Display percentage values
                .ShowValue = False      ' Do not display actual numbers
                .Position = xlLabelPositionInsideEnd
            End With
        End With
    End With
    
    ' Save the chart as an image and display it in UserForm
    chartFilePath = ThisWorkbook.Path & "\ChartImage.bmp"
    
    ' Delete existing file first (in case it is corrupted)
    If Dir(chartFilePath) <> "" Then Kill chartFilePath
    
    chartObj.Chart.Export fileName:=chartFilePath, FilterName:="BMP"
    
    ' Load the image into the UserForm
    Me.imgChart.PictureSizeMode = fmPictureSizeModeClip
    Me.imgChart.Picture = LoadPicture(chartFilePath)
    
    ' Hide the worksheet to avoid user interference
    ' wsChart.Visible = xlSheetVeryHidden
    ' wsChart.Visible = xlSheetVisible

    ' Delete existing BMP file
    If Dir(chartFilePath) <> "" Then Kill chartFilePath
    
    ' Cleanup
    Set wsChart = Nothing
    Set rngChartData = Nothing
    Set chartObj = Nothing

End Sub
