VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOpenRecipeFile 
   Caption         =   "Recipe File Manager"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280
   OleObjectBlob   =   "frmOpenRecipeFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOpenRecipeFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sortColumn As Integer
Dim sortAscending As Boolean

Private Sub btnOpen_Click()
    Dim selectedRecipe As String
    Dim selectedID As String
    Dim recipePath As String
    Dim filePath As String

    ' Ensure a selection is made
    If Me.lstRecipeItems.selectedItem Is Nothing Then
        MsgBox "Please select a recipe!", vbExclamation, "No Selection"
        Exit Sub
    End If

    ' Retrieve selected Recipe Name and Recipe ID
    selectedRecipe = Me.lstRecipeItems.selectedItem.SubItems(1) ' Recipe Name
    selectedID = Me.lstRecipeItems.selectedItem.SubItems(2) ' Recipe ID

    ' Recipe path validation
    Call recipePathValidation
    If Not recipePathValidBool Then Exit Sub

    ' Construct the full file path
    recipePath = ThisWorkbook.Path & "\Recipes\"
    filePath = recipePath & selectedRecipe & "_" & selectedID & ".xlsx"

    ' Open the selected recipe file
    If Dir(filePath) <> "" Then
        Workbooks.Open filePath
    Else
        MsgBox "File not found: " & filePath, vbCritical
    End If

    Me.lstRecipeItems.selectedItem = Nothing ' Clear selection
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub lstRecipeItems_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    ' Clear selection when the user clicks on an empty area of the ListView
    Me.lstRecipeItems.selectedItem = Nothing
End Sub

Private Sub frmRecipeFrame_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Clear selection when the user clicks on an empty area of the ListView
    Me.lstRecipeItems.selectedItem = Nothing
End Sub

' Trigger btnOpen_Click when an item is double-clicked
Private Sub lstRecipeItems_DblClick()
    ' Ensure an item is selected before executing the open procedure
    If Not Me.lstRecipeItems.selectedItem Is Nothing Then
        Call btnOpen_Click ' Call the Open button procedure
    End If
End Sub

' Handle column click event for sorting
Private Sub lstRecipeItems_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    ' Ignore sorting if "No." column is clicked
    If ColumnHeader.Index = 1 Then Exit Sub ' No. column is always fixed

    ' Toggle sorting order if the same column is clicked again
    If ColumnHeader.Index - 1 = sortColumn Then
        sortAscending = Not sortAscending
    Else
        sortColumn = ColumnHeader.Index - 1
        sortAscending = True
    End If

    ' Perform sorting
    SortListView
End Sub

Private Sub SortListView()
    Dim i As Integer, j As Integer
    Dim tempNo As String
    Dim tempText As String
    Dim tempID As String
    Dim isNumber As Boolean

    ' Check if sorting by Recipe ID (Column 2) should be numeric
    isNumber = (sortColumn = 2)

    ' Bubble Sort Algorithm for sorting ListView items (excluding No. column)
    For i = 1 To Me.lstRecipeItems.ListItems.Count - 1
        For j = i + 1 To Me.lstRecipeItems.ListItems.Count
            Dim swap As Boolean
            swap = False

            ' Compare values (ignoring No. column)
            If isNumber Then
                ' Numeric comparison for Recipe ID
                If sortAscending Then
                    swap = Val(Me.lstRecipeItems.ListItems(i).SubItems(2)) > Val(Me.lstRecipeItems.ListItems(j).SubItems(2))
                Else
                    swap = Val(Me.lstRecipeItems.ListItems(i).SubItems(2)) < Val(Me.lstRecipeItems.ListItems(j).SubItems(2))
                End If
            Else
                ' Text comparison for Recipe Name
                If sortAscending Then
                    swap = StrComp(Me.lstRecipeItems.ListItems(i).SubItems(1), Me.lstRecipeItems.ListItems(j).SubItems(1), vbTextCompare) > 0
                Else
                    swap = StrComp(Me.lstRecipeItems.ListItems(i).SubItems(1), Me.lstRecipeItems.ListItems(j).SubItems(1), vbTextCompare) < 0
                End If
            End If

            ' Swap items if needed (No. column remains unchanged)
            If swap Then
                tempNo = Me.lstRecipeItems.ListItems(i).Text ' Store No. value
                tempText = Me.lstRecipeItems.ListItems(i).SubItems(1)
                tempID = Me.lstRecipeItems.ListItems(i).SubItems(2)

                ' Swap Recipe Name & Recipe ID (No. column stays the same)
                Me.lstRecipeItems.ListItems(i).SubItems(1) = Me.lstRecipeItems.ListItems(j).SubItems(1)
                Me.lstRecipeItems.ListItems(i).SubItems(2) = Me.lstRecipeItems.ListItems(j).SubItems(2)

                Me.lstRecipeItems.ListItems(j).SubItems(1) = tempText
                Me.lstRecipeItems.ListItems(j).SubItems(2) = tempID
            End If
        Next j
    Next i
End Sub

' Initialize the ListView and populate it with recipe files
Private Sub UserForm_Initialize()
    Dim colHeader As ColumnHeader
    Dim recipePath As String
    Dim fileName As String
    Dim itemX As listItem
    Dim recipeParts As Variant
    Dim fileCounter As Integer

    ' Define the path to the Recipes folder
    recipePath = ThisWorkbook.Path & "\Recipes\"

    ' Recipe path validation
    Call recipePathValidation
    frmOpenRecipeFileBool = True
    If Not recipePathValidBool Then
        frmOpenRecipeFileBool = False
        Exit Sub
    End If

    ' Prevent editing of the first column (Product ID)
    Me.lstRecipeItems.LabelEdit = lvwManual
    
    ' Configure ListView properties
    With Me.lstRecipeItems
        .ColumnHeaders.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .Sorted = False ' Sorting will be handled manually

        ' Add column headers
        Set colHeader = .ColumnHeaders.Add(, , "No.", 30) ' Fixed column
        Set colHeader = .ColumnHeaders.Add(, , "Recipe Name", 200)
        Set colHeader = .ColumnHeaders.Add(, , "Recipe ID", 80)
    End With

    ' Retrieve Excel files from the folder
    fileName = Dir(recipePath & "*.xls*")
    fileCounter = 1

    Do While fileName <> ""
        ' Remove the file extension
        fileName = Left(fileName, InStrRev(fileName, ".") - 1)

        ' Split the filename by underscore to get Recipe Name and Recipe ID
        recipeParts = Split(fileName, "_")

        ' Ensure the filename contains at least two parts (Recipe Name & Recipe ID)
        If UBound(recipeParts) >= 1 Then
            ' Add an entry to the ListView
            Set itemX = Me.lstRecipeItems.ListItems.Add(, , fileCounter) ' No
            itemX.SubItems(1) = recipeParts(0) ' Recipe Name
            itemX.SubItems(2) = recipeParts(1) ' Recipe ID

            fileCounter = fileCounter + 1
        End If

        ' Get the next file
        fileName = Dir
    Loop

    ' Default sorting column (Recipe Name) on load
    ' sortColumn = 1 ' Recipe Name column
    sortColumn = 2 ' Recipe ID column
    sortAscending = True
    SortListView ' Call sorting function to sort on initialization
End Sub
