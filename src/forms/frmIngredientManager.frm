VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIngredientManager 
   Caption         =   "Ingredient Manager"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9615
   OleObjectBlob   =   "frmIngredientManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmIngredientManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim productIDCheckBool As Boolean
Dim productIDValidBool As Boolean
Dim mandatoryFieldBool As Boolean
Dim rangeProduct As Range
Dim wsProduct As Worksheet
Dim productID As String

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

Private Sub btnLoadProduct_Click() ' Loads data from the database into the form fields
    Call UnprotectAllSheets
    
    ' Validate the Product ID
    Call productIDValid
    If Not productIDCheckBool Then GoTo ExitHandler

    ' Display an message if Product ID is not found
    If Not productIDValidBool Then
        MsgBox "Product ID not found!", vbExclamation, "Invalid Operation"
        clearUserForm
        GoTo ExitHandler
    End If
    
    ' Populate form fields with existing record data
    If Not rangeProduct Is Nothing Then
        Me.txtProductName.value = CStr(wsProduct.Cells(rangeProduct.Row, 2).value)
        Me.txtBrand.value = CStr(wsProduct.Cells(rangeProduct.Row, 3).value)
        Me.txtCost.value = CStr(wsProduct.Cells(rangeProduct.Row, 4).value)
        Me.txtAmount.value = CStr(wsProduct.Cells(rangeProduct.Row, 5).value)
        Me.txtFat.value = CStr(wsProduct.Cells(rangeProduct.Row, 6).value)
        Me.txtSugar.value = CStr(wsProduct.Cells(rangeProduct.Row, 7).value)
        Me.txtSalt.value = CStr(wsProduct.Cells(rangeProduct.Row, 8).value)
    End If
    ' Cleanup
    Set wsProduct = Nothing
    Set rangeProduct = Nothing
   
ExitHandler:
    Call ProtectAllSheets
End Sub

Private Sub btnUpdateProduct_Click() ' Update the ingredient record, including associated recipe indexes and files.
    Call UnprotectAllSheets
    
    ' Recipe path validation
    Call recipePathValidation
    If Not recipePathValidBool Then GoTo ExitHandler
    
    ' Validate the Product ID
    Call productIDValid
    If Not productIDCheckBool Then GoTo ExitHandler
    
    ' Display an message if Product ID is not found
    If Not productIDValidBool Then
        MsgBox "Product ID does not exist in the database. Please check and try again.", vbExclamation, "Invalid Update Operation"
        Me.txtProductID.value = ""
        GoTo ExitHandler
    End If
 
    ' Trim all spaces
    Call trimAllSpaces

    ' Check if mandatory fields are filled
    Call mandatoryField
    If Not mandatoryFieldBool Then GoTo ExitHandler
    
    ' Ensure numeric validation before proceeding
    If Not isNumericTextboxValid(Me, Me.txtCost, "Cost") Then GoTo ExitHandler
    If Not isNumericTextboxValid(Me, Me.txtAmount, "Amount") Then GoTo ExitHandler
    If Not isNumericTextboxValid(Me, Me.txtFat, "Fat") Then GoTo ExitHandler
    If Not isNumericTextboxValid(Me, Me.txtSugar, "Sugar") Then GoTo ExitHandler
    If Not isNumericTextboxValid(Me, Me.txtSalt, "Salt") Then GoTo ExitHandler

    ' If any of these TextBox fields are empty, assign them a default value of 0
    Call addZero
    
    ' Update an existing record in Product Data
    wsProduct.Cells(rangeProduct.Row, 2).value = CStr(Me.txtProductName.value)
    wsProduct.Cells(rangeProduct.Row, 3).value = CStr(Me.txtBrand.value)
    wsProduct.Cells(rangeProduct.Row, 4).value = CDbl(Me.txtCost.value)
    wsProduct.Cells(rangeProduct.Row, 5).value = CDbl(Me.txtAmount.value)
    wsProduct.Cells(rangeProduct.Row, 6).value = CDbl(Me.txtFat.value)
    wsProduct.Cells(rangeProduct.Row, 7).value = CDbl(Me.txtSugar.value)
    wsProduct.Cells(rangeProduct.Row, 8).value = CDbl(Me.txtSalt.value)
    
    'Update recipe files after changing the ingredients of the product ID
    Application.ScreenUpdating = False
    Call CollectAffectedRecipes(productID, wsProduct, rangeProduct)
    Application.ScreenUpdating = True

    ' Cleanup
    Set wsProduct = Nothing
    Set rangeProduct = Nothing
    
    MsgBox "The product and all associated recipe files have been updated successfully.", vbInformation, "Update Successful"
    
    ' Clear form fields after operation
    Call clearUserForm
    
ExitHandler:
    Call ProtectAllSheets
End Sub

Private Sub btnNewProduct_Click() ' Adds a new ingredient record
    Dim lastRow As Long
    Dim wsRange As Range
    
    Call UnprotectAllSheets
    
    ' Check the Product ID
    Call productIDValid
    If Not productIDCheckBool Then GoTo ExitHandler

    ' Display an message if Product ID is found
    If productIDValidBool Then
        MsgBox "The entered Product ID already exists in the database. Please use a unique ID to add a new product / ingredient.", vbExclamation, "Duplicate Product ID"
        ' Me.txtProductID.value = ""
        GoTo ExitHandler
    End If
    
    ' Trim all spaces
    Call trimAllSpaces

    ' Check if mandatory fields are filled
    Call mandatoryField
    If Not mandatoryFieldBool Then GoTo ExitHandler

    ' Ensure numeric validation before proceeding
    If Not isNumericTextboxValid(Me, Me.txtCost, "Cost") Then GoTo ExitHandler
    If Not isNumericTextboxValid(Me, Me.txtAmount, "Amount") Then GoTo ExitHandler
    If Not isNumericTextboxValid(Me, Me.txtFat, "Fat") Then GoTo ExitHandler
    If Not isNumericTextboxValid(Me, Me.txtSugar, "Sugar") Then GoTo ExitHandler
    If Not isNumericTextboxValid(Me, Me.txtSalt, "Salt") Then GoTo ExitHandler

    ' If any of these TextBox fields are empty, assign them a default value of 0
    Call addZero

    ' Add a new ingredient to the database
    lastRow = wsProduct.Cells(wsProduct.Rows.Count, 1).End(xlUp).Row + 1
    wsProduct.Cells(lastRow, 1).value = productID
    wsProduct.Cells(lastRow, 2).value = CStr(Me.txtProductName.value)
    wsProduct.Cells(lastRow, 3).value = CStr(Me.txtBrand.value)
    wsProduct.Cells(lastRow, 4).value = CDbl(Me.txtCost.value)
    wsProduct.Cells(lastRow, 5).value = CDbl(Me.txtAmount.value)
    wsProduct.Cells(lastRow, 6).value = CDbl(Me.txtFat.value)
    wsProduct.Cells(lastRow, 7).value = CDbl(Me.txtSugar.value)
    wsProduct.Cells(lastRow, 8).value = CDbl(Me.txtSalt.value)
    
    ' Add border
    Set wsRange = wsProduct.Range("A" & lastRow & ":I" & lastRow)
    Call addBorder(wsProduct, wsRange)
    
    ' Cleanup
    Set wsProduct = Nothing
    Set rangeProduct = Nothing
    Set wsRange = Nothing
    
    MsgBox "The new ingredient has been added successfully.", vbInformation, "Addition Successful"

    ' Clear form fields after operation
     Call clearUserForm

ExitHandler:
    Call ProtectAllSheets
End Sub

Private Sub btnDeleteProduct_Click() ' Deletes an existing ingredient record
    Dim confirmation As Variant
    Dim lastRow As Long
    Dim wsRange As Range
    
    Call UnprotectAllSheets
    
    ' Recipe path validation
    Call recipePathValidation
    If Not recipePathValidBool Then GoTo ExitHandler
    
    ' Validate the Product ID
    Call productIDValid
    If Not productIDCheckBool Then GoTo ExitHandler

    ' Check if mandatory fields are filled
    Call mandatoryField
    If Not mandatoryFieldBool Then GoTo ExitHandler

    ' Ensure that the Product ID exists in the database
    If Not productIDValidBool Then
        MsgBox "Product ID does not exist in the database. Please check and try again.", vbExclamation, "Invalid Operation"
        Me.txtProductID.value = ""
        GoTo ExitHandler
    End If

    ' Confirm deletion from the user
    confirmation = MsgBox("Are you sure you want to delete the ingredient with Product ID: " & productID & "?", vbYesNo + vbQuestion, "Confirm Deletion")
    If confirmation = vbYes Then
        
        'Update recipe files after deleting the product ID
        Application.ScreenUpdating = False
        Call RemoveProductIDFromRecipeIndex(productID, wsProduct, rangeProduct) ' Remove product ID from recipe index
        Call CollectAffectedRecipes(productID, wsProduct, rangeProduct) 'Update recipe files
        Application.ScreenUpdating = True
        
        wsProduct.Rows(rangeProduct.Row).Delete 'Remove the product id from product page

        ' Add border
        lastRow = wsProduct.Cells(wsProduct.Rows.Count, 1).End(xlUp).Row
        Set wsRange = wsProduct.Range("A" & lastRow & ":I" & lastRow)
        Call addBorder(wsProduct, wsRange)
    
        ' Cleanup
        Set wsProduct = Nothing
        Set rangeProduct = Nothing
        Set wsRange = Nothing
    
        MsgBox "The product has been deleted, and all related recipe files have been updated successfully.", vbInformation, "Deletion & Update Successful"
        
        ' Clear form fields after deletion
        Call clearUserForm
    End If
    
ExitHandler:
    Call ProtectAllSheets
End Sub

Private Sub btnClearForm_Click() ' Clear the UserForm
    clearUserForm
End Sub

Private Sub btnClose_Click() ' Close the UserForm
    Unload Me
End Sub

Private Sub clearUserForm() ' Clears all input fields in the form
    Me.txtProductID.value = ""
    Me.txtProductName.value = ""
    Me.txtBrand.value = ""
    Me.txtCost.value = ""
    Me.txtAmount.value = ""
    Me.txtFat.value = ""
    Me.txtSugar.value = ""
    Me.txtSalt.value = ""
End Sub

Private Sub mandatoryField()
    mandatoryFieldBool = True
    
    ' Ensure the product name is provided
    If Me.txtProductName.value = "" Then
        mandatoryFieldBool = False
        MsgBox "Please enter product details before proceeding.", vbExclamation, "Missing Data"
        Exit Sub
    End If
    
    ' Ensure the Brand / Supplier is provided
    If Me.txtBrand.value = "" Then
        mandatoryFieldBool = False
        MsgBox "Please enter brand / supplier details before proceeding.", vbExclamation, "Missing Data"
        Exit Sub
    End If

    ' Ensure the Cost is provided
    If Me.txtCost.value = "" Then
        mandatoryFieldBool = False
        MsgBox "Please enter cost details before proceeding.", vbExclamation, "Missing Data"
        Exit Sub
    End If

    ' Ensure the Amount is provided
    If Me.txtAmount.value = "" Then
        mandatoryFieldBool = False
        MsgBox "Please enter amount details before proceeding.", vbExclamation, "Missing Data"
        Exit Sub
    End If
    
End Sub

Private Sub trimAllSpaces()
    ' Trim all spaces (Text)
    Me.txtProductName.value = Trim(Me.txtProductName.value)
    Me.txtBrand.value = Trim(Me.txtBrand.value)
    ' Remove extra spaces within the text fields and replace them with a single space
    Me.txtProductName.value = ReduceSpaces(Me.txtProductName.value)
    Me.txtBrand.value = ReduceSpaces(Me.txtBrand.value)

    ' Trim and remove all spaces (Numeric)
    Me.txtCost.value = Replace(Trim(Me.txtCost.value), " ", "")
    Me.txtAmount.value = Replace(Trim(Me.txtAmount.value), " ", "")
    Me.txtFat.value = Replace(Trim(Me.txtFat.value), " ", "")
    Me.txtSugar.value = Replace(Trim(Me.txtSugar.value), " ", "")
    Me.txtSalt.value = Replace(Trim(Me.txtSalt.value), " ", "")
    
End Sub

Private Sub addZero()
    ' If any of these TextBox fields are empty, assign them a default value of 0
    If Me.txtCost.value = "" Then Me.txtCost.value = 0
    If Me.txtAmount.value = "" Then Me.txtAmount.value = 0
    If Me.txtFat.value = "" Then Me.txtFat.value = 0
    If Me.txtSugar.value = "" Then Me.txtSugar.value = 0
    If Me.txtSalt.value = "" Then Me.txtSalt.value = 0
End Sub

