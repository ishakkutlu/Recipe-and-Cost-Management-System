Attribute VB_Name = "mdlFunction"
Option Explicit

Function isNumericTextboxValid(frm As Object, txt As MSForms.TextBox, fieldLabel As String) As Boolean
    ' This function validates if the given TextBox contains only numeric values such as 1.05, and 1,000.25.
    ' It also ensures that decimal and thousand separators match Excel's system settings.

    Dim decimalSeparator As String
    Dim thousandSeparator As String
    Dim value As String
    Dim numParts() As String
    Dim integerPart As String
    Dim decimalPart As String
    Dim thousandGroups() As String
    Dim i As Integer
    Dim thousandCount As Integer
    Dim decimalCount As Integer

    ' Get system decimal and thousand separators
    decimalSeparator = Application.International(xlDecimalSeparator)
    thousandSeparator = Application.International(xlThousandsSeparator)

    ' Get the textbox value (remove spaces)
    value = Replace(Trim(txt.value), " ", "")

    ' Assume valid by default
    isNumericTextboxValid = True

    ' If empty, skip validation
    If value = "" Then Exit Function

    ' Count occurrences of separators
    thousandCount = Len(value) - Len(Replace(value, thousandSeparator, ""))
    decimalCount = Len(value) - Len(Replace(value, decimalSeparator, ""))

    ' More than one decimal separator is invalid
    If decimalCount > 1 Then
        MsgBox fieldLabel & " value must be numeric only! (Too many decimal separators)", vbExclamation, "Invalid Input"
        isNumericTextboxValid = False
        Exit Function
    End If

    ' Thousand separator must not appear after decimal separator
    If InStr(value, decimalSeparator) > 0 And InStr(value, thousandSeparator) > InStr(value, decimalSeparator) Then
        MsgBox fieldLabel & " value must be numeric only! (Thousand separator after decimal separator)", vbExclamation, "Invalid Input"
        isNumericTextboxValid = False
        Exit Function
    End If

    ' Check if value contains only valid characters (digits and separators)
    If value Like "*[!0-9" & decimalSeparator & thousandSeparator & "]*" Then
        MsgBox fieldLabel & " value must be numeric only!", vbExclamation, "Invalid Input"
        isNumericTextboxValid = False
        Exit Function
    End If

    ' Value must NOT end with a decimal separator or thousand separator
    If Right(value, 1) = decimalSeparator Or Right(value, 1) = thousandSeparator Then
        MsgBox fieldLabel & " value must be numeric only! (Cannot end with separator)", vbExclamation, "Invalid Input"
        isNumericTextboxValid = False
        Exit Function
    End If

    ' Split by decimal separator
    numParts = Split(value, decimalSeparator)

    ' Get integer and decimal parts safely
    integerPart = numParts(0)
    If UBound(numParts) = 1 Then
        decimalPart = numParts(1)
    Else
        decimalPart = ""
    End If

    ' Validate thousand separator usage in the integer part
    If InStr(integerPart, thousandSeparator) > 0 Then
        thousandGroups = Split(integerPart, thousandSeparator)
        
        ' The first group can have 1-3 digits, but all others must have exactly 3
        If Len(thousandGroups(0)) < 1 Or Len(thousandGroups(0)) > 3 Then
            MsgBox fieldLabel & " value must be numeric only! (Invalid thousands separator grouping)", vbExclamation, "Invalid Input"
            isNumericTextboxValid = False
            Exit Function
        End If

        ' Ensure that all other groups are exactly 3 digits
        For i = 1 To UBound(thousandGroups)
            If Len(thousandGroups(i)) <> 3 Then
                MsgBox fieldLabel & " value must be numeric only! (Invalid thousands grouping)", vbExclamation, "Invalid Input"
                isNumericTextboxValid = False
                Exit Function
            End If
        Next i
    End If

    ' Ensure decimal part (if exists) is only numeric
    If decimalPart <> "" And Not IsNumeric(decimalPart) Then
        MsgBox fieldLabel & " value must be numeric only! (Invalid decimal format)", vbExclamation, "Invalid Input"
        isNumericTextboxValid = False
        Exit Function
    End If
End Function

Function isOnlyNumeric(inputValue As String) As Boolean
    ' This function checks if the inputValue contains only numeric characters (0-9).
    
    Dim i As Integer
    
    ' Assume valid by default
    isOnlyNumeric = True

    ' If empty, skip validation
    If inputValue = "" Then Exit Function

    ' Check each character in the input string
    For i = 1 To Len(inputValue)
        If Mid(inputValue, i, 1) Like "[!0-9]" Then
            ' MsgBox "Only numeric characters (0-9) are allowed!", vbExclamation, "Invalid Input"
            isOnlyNumeric = False
            Exit Function
        End If
    Next i
End Function

Function ReduceSpaces(ByVal inputText As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Pattern = "\s+"  ' One or more spaces
        .Global = True     ' Replace all occurrences
        ReduceSpaces = .Replace(inputText, " ")
    End With
End Function




