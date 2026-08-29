Attribute VB_Name = "Test"
Option Explicit

'namespace=xvba_unit_test/Test

'/*
'FAS unit checks. XVBA includes this folder in Production exports.
'Run test from the XVBA macro list after a Production export.
'*/
Public Sub test()
    FAS_TestShortcutParsing
    FAS_TestShortcutDefinitions
    FAS_TestCommandCatalog
    FAS_TestCommandParsingContracts
    FAS_TestFormatCycleContracts
    FAS_TestIFSContracts
    FAS_TestGroupByContracts
    FAS_TestViewAndUndoContracts

    MsgBox "FAS xVBA unit tests passed.", vbInformation, "FAS Tests"
End Sub

Public Sub FAS_TestGroupByContracts()
    Dim rowFieldsText As String
    Dim valueFieldsText As String
    Dim functionName As String

    FAS_TestAssertTrue FAS_ParseGroupByFormula("=GROUPBY(V49:X139; AG49:AI139; SUM; 3;-3;-4)", _
        rowFieldsText, valueFieldsText, functionName), "GROUPBY semicolon parser"
    FAS_TestAssertEqual "V49:X139", rowFieldsText, "GROUPBY row fields"
    FAS_TestAssertEqual "AG49:AI139", valueFieldsText, "GROUPBY value fields"
    FAS_TestAssertEqual "SUM", functionName, "GROUPBY aggregation"
    FAS_TestAssertTrue FAS_ParseGroupByFormula("=GROUPBY(A2:B10,C2:D10,COUNT,3,-3)", _
        rowFieldsText, valueFieldsText, functionName), "GROUPBY comma parser"
    FAS_TestAssertEqual "COUNT", functionName, "GROUPBY count aggregation"
    FAS_TestAssertTrue FAS_IsKnownCommand("groupby"), "Group By command registry"
    FAS_TestAssertTrue FAS_IsKnownCommand("groupbycount"), "Group By count command registry"
    FAS_TestAssertTrue FAS_CountCommandCatalogMatches("groupby") > 0, "Group By catalog entry"
End Sub

Public Sub FAS_TestCommandParsingContracts()
    Dim queryText As Variant
    Dim groupKey As String
    Dim slotIndex As Long

    FAS_TestAssertTrue FAS_IsCommandHelpCommand("help gridlines", queryText), "Command help prefix"
    FAS_TestAssertEqual "gridlines", CStr(queryText), "Command help query extraction"
    FAS_TestAssertTrue FAS_IsKnownCommand("round 2"), "Parameterized formula command"
    FAS_TestAssertTrue FAS_ParseJumpCommand("tb2", groupKey, slotIndex), "Parameterized sheet jump"
    FAS_TestAssertEqual "TB", groupKey, "Sheet jump group"
    FAS_TestAssertEqual 2, slotIndex, "Sheet jump slot"
End Sub

Public Sub FAS_TestShortcutParsing()
    Dim errorText As String
    Dim actualValue As String

    FAS_TestAssertEqual "Ctrl+Alt+Up", FAS_NormalizeShortcutDisplay(" alt + ctrl + up "), "Shortcut normalization"
    FAS_TestAssertEqual "^%{UP}", FAS_ShortcutDisplayToOnKey("Ctrl+Alt+Up"), "Arrow shortcut conversion"
    FAS_TestAssertEqual "+{F2}", FAS_ShortcutDisplayToOnKey("Shift+F2"), "Function shortcut conversion"
    FAS_TestAssertEqual vbNullString, FAS_ShortcutDisplayToOnKey(vbNullString), "Blank shortcut conversion"

    errorText = vbNullString
    actualValue = FAS_ShortcutDisplayToOnKey("Ctrl+Shift", errorText)
    FAS_TestAssertTrue Len(actualValue) = 0 And Len(errorText) > 0, "Invalid shortcut is rejected"

    errorText = vbNullString
    actualValue = FAS_ShortcutDisplayToOnKey("J", errorText)
    FAS_TestAssertTrue Len(actualValue) = 0 And Len(errorText) > 0, "Unmodified shortcut is rejected"
End Sub

Public Sub FAS_TestShortcutDefinitions()
    Dim definitions As Collection
    Dim item As Variant

    Set definitions = FAS_GetShortcutDefinitions()
    FAS_TestAssertEqual 10, definitions.Count, "Shortcut definition count"

    For Each item In definitions
        FAS_TestAssertTrue Len(CStr(item(0))) > 0, "Shortcut action id"
        FAS_TestAssertTrue Len(CStr(item(1))) > 0, "Shortcut display name"
        FAS_TestAssertTrue Len(CStr(item(3))) > 0, "Shortcut macro name"
        FAS_TestAssertTrue Len(FAS_GetDefaultShortcut(CStr(item(0)))) > 0, "Default shortcut exists"
    Next item

    FAS_TestAssertEqual "Ctrl+Shift+T", FAS_GetDefaultShortcut("tb"), "TB default shortcut"
    FAS_TestAssertEqual "Ctrl+Shift+1", FAS_GetDefaultShortcut("numberformat"), "Number format default shortcut"
    FAS_TestAssertEqual "Ctrl+Shift+5", FAS_GetDefaultShortcut("percentformat"), "Percent format default shortcut"
    FAS_TestAssertEqual "Ctrl+Shift+4", FAS_GetDefaultShortcut("currencyformat"), "Currency format default shortcut"
    FAS_TestAssertTrue FAS_IsKnownShortcutAction("palette"), "Known shortcut action"
End Sub

Public Sub FAS_TestFormatCycleContracts()
    FAS_TestAssertTrue FAS_IsFormatCycleCommand("format number"), "Number format cycle command"
    FAS_TestAssertTrue FAS_IsFormatCycleCommand("format percent"), "Percent format cycle command"
    FAS_TestAssertTrue FAS_IsFormatCycleCommand("format currency"), "Currency format cycle command"
    FAS_TestAssertTrue FAS_IsKnownCommand("format number"), "Number format command registry"
    FAS_TestAssertTrue FAS_CountCommandCatalogMatches("format percent") > 0, "Percent format catalog entry"
    FAS_TestAssertTrue FAS_CountCommandCatalogMatches("format settings") > 0, "Format settings catalog entry"
End Sub

Public Sub FAS_TestCommandCatalog()
    FAS_TestAssertTrue FAS_CountCommandCatalogMatches("linkpaste") > 0, "Existing command catalog entry"
    FAS_TestAssertTrue FAS_CountCommandCatalogMatches("shortcuts") > 0, "Shortcut settings catalog entry"
    FAS_TestAssertTrue FAS_CountCommandCatalogMatches("capex") > 0, "CAPEX summary catalog entry"
    FAS_TestAssertTrue FAS_IsKnownShortcutCommand("shortcuts reset"), "Shortcut reset command"
    FAS_TestAssertTrue FAS_IsKnownCommand("shortcuts"), "Shortcut command registry"
    FAS_TestAssertTrue FAS_IsKnownCommand("capex"), "CAPEX command registry"
    FAS_TestAssertEqual "acquisition cost", FAS_NormalizeCapexHeader(" Acquisition  Cost "), "CAPEX header normalization"
    FAS_TestAssertEqual "2024", FAS_ExtractCapexYear("CAPEX - Year 2024"), "CAPEX title year"
    FAS_TestAssertEqual "class i", FAS_NormalizeCapexHeader(" Class I "), "CAPEX class header normalization"
    FAS_TestAssertEqual "2025", FAS_ExtractCapexYear("CAPEX 2025 versus 2024"), "CAPEX first title year"
End Sub

Public Sub FAS_TestIFSContracts()
    Dim functionName As String
    Dim argumentCount As Long
    Dim convertedFormula As String

    FAS_TestAssertTrue FAS_ParseIFSFormula("=SUMIFS($E$2:$E$10,$A$2:$A$10,$G2)", functionName, argumentCount), "SUMIFS parser"
    FAS_TestAssertEqual "SUMIFS", functionName, "SUMIFS function name"
    FAS_TestAssertEqual 3, argumentCount, "SUMIFS argument count"
    FAS_TestAssertTrue FAS_ConvertIFSFormula("=SUMIFS($E$2:$E$10,$A$2:$A$10,$G2)", "COUNTIFS", vbNullString, convertedFormula), "Measure to COUNTIFS conversion"
    FAS_TestAssertEqual "=COUNTIFS($A$2:$A$10,$G2)", convertedFormula, "Measure removed for COUNTIFS"
    FAS_TestAssertTrue FAS_ConvertIFSFormula("=COUNTIFS($A$2:$A$10,$G2,$B$2:$B$10,$H2)", "AVERAGEIFS", "Sheet1!$E$2:$E$10", convertedFormula), "COUNTIFS measure conversion"
    FAS_TestAssertEqual "=AVERAGEIFS(Sheet1!$E$2:$E$10,$A$2:$A$10,$G2,$B$2:$B$10,$H2)", convertedFormula, "COUNTIFS measure inserted once"
    FAS_TestAssertTrue Not FAS_ParseIFSFormula("=SUBTOTAL(9,$E$2:$E$10)", functionName, argumentCount), "SUBTOTAL is not a direct IFS formula"
    FAS_TestAssertTrue FAS_IsKnownCommand("ifs sum"), "IFS builder command registry"
    FAS_TestAssertTrue FAS_IsKnownCommand("switch averageifs"), "IFS switch command registry"
End Sub

Public Sub FAS_TestViewAndUndoContracts()
    FAS_TestAssertTrue FAS_DEFAULT_HIDE_GRIDLINES, "Gridlines are hidden by default"
    FAS_TestAssertTrue FAS_IsKnownCommand("hidegridlines"), "Hide-gridlines command registry"
    FAS_TestAssertTrue FAS_CountCommandCatalogMatches("gridlines") >= 2, "Gridlines command catalog entries"
    FAS_TestAssertTrue FAS_DEFAULT_OUTLINE_ABOVE_LEFT, "Outline direction is above and left by default"
    FAS_TestAssertTrue FAS_IsKnownCommand("outlineaboveleft"), "Outline direction command registry"
    FAS_TestAssertTrue FAS_CountCommandCatalogMatches("outline") >= 2, "Outline direction command catalog entries"
    FAS_TestAssertTrue FAS_UndoMaxCellCountForTests() > 0, "Undo snapshot capacity"
End Sub

Private Sub FAS_TestAssertEqual(ByVal expectedValue As Variant, ByVal actualValue As Variant, ByVal testName As String)
    If CStr(expectedValue) <> CStr(actualValue) Then
        Err.Raise vbObjectError + 6200, "FAS xVBA unit test", testName & " failed. Expected '" & CStr(expectedValue) & "', got '" & CStr(actualValue) & "'."
    End If
End Sub

Private Sub FAS_TestAssertTrue(ByVal conditionValue As Boolean, ByVal testName As String)
    If Not conditionValue Then
        Err.Raise vbObjectError + 6201, "FAS xVBA unit test", testName & " failed."
    End If
End Sub
