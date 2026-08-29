'This Script copy files from VSCod to Excel
'@create: Alberto Eduardo
'@Since: 18/04/2020
'@Version: 1.0


Dim  CURRENT_DIR, LOG_FILE_PATH,TMP_FOLDER, FSO,VBA_TMP_FOLDER
Set FSO = CreateObject("Scripting.FileSystemObject")


Const XML_FILE_NAME = "EXP-SDRRMACRO-EXCYURI-XVARGEMX-A40149D.xml"
Const XVBA_TEMP_FOLDER_NAME = "\xvba_files"
Const WORKSPACE_LOG_FOLDER_NAME = "logs"
Const WORKSPACE_SERVER_LOG_FILE_NAME  = "excel_server.log"
Const WORKSPACE_IMPORT_LOG_FILE_NAME  = "excel_import.log"
Const TEMP_EXPORT_FILE_LOG_NAME = "xvba_exported_file.log"
Const TEMP_CODE_DEBUG_PARSED_FOLDER = "xvba_code_debug_parsed"
Const WORKSPACE_VBA_MODULES_FOLDER_NAME = "xvba_modules"
Const WORKSPACE_VBA_TEST_FOLDER_NAME = "xvba_unit_test"

CURRENT_DIR    = FSO.GetParentFolderName(WScript.ScriptFullName)
TMP_FOLDER = FSO.GetSpecialFolder(2)
VBA_TMP_FOLDER = TMP_FOLDER & XVBA_TEMP_FOLDER_NAME

'VSCode enviroment Variables
Dim VSCODE_ENV
On Error Resume Next
Set VSCODE_ENV = getActionFromXML()
Call ErrorHandler("BootStrap")

LOG_FILE_PATH = VSCODE_ENV("WorkspacePath") & "\" & WORKSPACE_LOG_FOLDER_NAME & "\" & WORKSPACE_SERVER_LOG_FILE_NAME
LOG_FILE_EXPORT =  VSCODE_ENV("WorkspacePath") & "\" & WORKSPACE_LOG_FOLDER_NAME & "\" & WORKSPACE_IMPORT_LOG_FILE_NAME


If VSCODE_ENV("Command") = "RUN_AUTO" Then
	On Error Resume Next
	Call AutoloadModules(VSCODE_ENV("Prop"))
	Call ErrorHandler("AutoloadModules")
ElseIf VSCODE_ENV("Command") = "RUN_CHANGED" Then
	On Error Resume Next
	Call ImportChangedModules(VSCODE_ENV("ChangedFiles"))
	Call ErrorHandler("ImportChangedModules")
Else
	On Error Resume Next
	Call ExportModules
	Call ErrorHandler("ExportModules")
End If


Public Function ErrorHandler(message)
	If Err.Number <> 0 Then
		WScript.Echo "Error in " & message & " : " & Err.Description
		Call LogMessages(0, LOG_FILE_PATH, "Error in" & message)
		Err.Clear
	End If
	
End Function


'Export modules to file


Public Function ExportModules()
	' As Boolean
	Dim bExport
	
	'As String
	Dim szFileName 
	'As VBIDE.VBComponent
	Dim cmpComponent 
	
	'Create temp folder for vba files
 	Call TestForDir(VBA_TMP_FOLDER)
	
	Dim book,pathWorkbook
	pathWorkbook = VSCODE_ENV("WorkspacePath") & "\" & VSCODE_ENV("WorkbookName")

	Set book =  GetObject(pathWorkbook).Application.Workbooks(VSCODE_ENV("WorkbookName"))
	
	Call LogMessages(1,LOG_FILE_EXPORT, "Export VBA Files To VSCode Start")
	
	For Each cmpComponent In book.VBProject.VBComponents
		
		bExport = True
		szFileName = cmpComponent.Name
		
		''' Concatenate the correct filename for export.
		Select Case cmpComponent.Type
			Case 2
			szFileName = szFileName & ".cls"
			Call LogMessages(1,LOG_FILE_EXPORT, "EXPORTING CLASS MODULES:" & szFileName & " Type: " & 2 & " - " & cmpComponent.Name)
			Case 3
			szFileName = szFileName & ".frm"
			Call LogMessages(1,LOG_FILE_EXPORT, "EXPORTING FORM MODULES:" & szFileName 	& "Type: " & 3 & " - " & cmpComponent.Name)
			Case 1
			szFileName = szFileName & ".bas"
			Call LogMessages(1,LOG_FILE_EXPORT, "EXPORTING MODULES:" & szFileName & " Type: " & 1 & " - " & cmpComponent.Name)
			Case 100
			szFileName = szFileName & ".cls"
			Call LogMessages(1,LOG_FILE_EXPORT, "EXPORTING CLASS MODULES:" & szFileName & " Type: " & 100 & " - " & cmpComponent.Name)
		End Select
		
		If bExport Then
			' Export the component to a text file.
			cmpComponent.Export VBA_TMP_FOLDER & "\" & szFileName
			
		End If
		
	Next 
	
	Call LogMessages(1,LOG_FILE_EXPORT, "Export VBA Files To VSCode Finish")
	Call finishExportModules()

	
End Function


'Create Log file when excel finish create all vba files on user folder
'Change 20/10/2020 - folder path to os temp folder
Public Function finishExportModules()
	
	'As String
	Dim szExportPath 
	
	
	szExportPath = TMP_FOLDER & "\" & TEMP_EXPORT_FILE_LOG_NAME
	
	'As FileSystemObject
	Dim fso 
	Set fso = CreateObject("Scripting.FileSystemObject")
	'As TextStream
	Dim fileStream 
	
	' Here the actual file is created and opened for write access
	Set fileStream = fso.CreateTextFile(szExportPath)
	
	' Write something to the file
	fileStream.WriteLine Now & " - Exported"
	
	' Close it, so it is not locked anymore
	fileStream.Close
	
End Function



'
'
'
Public Function TestForDir(strDir)
	
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(strDir) Then
		fso.CreateFolder (strDir)
	End If
	Set fso = Nothing
	
End Function


Public Function AutoloadModules(prop)
	
	
	
	'As Scripting.FileSystemObject
	Dim objFSO 
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	' Path to user vba files
	Dim szImportPath 
	
	'Check What mode user want to export files to excel
	'Locate the files folder
	If prop = "Dev (With Line Number)" Then
	   'Locate on OS TEMP Folder - Files Parsed
	  szImportPath = TMP_FOLDER & "\" & TEMP_CODE_DEBUG_PARSED_FOLDER
	Else
  	 '' Locate on VSCODE Workspace
    	szImportPath = VSCODE_ENV("WorkspacePath") & "\" & VSCODE_ENV("FolderVBA")
	End If

	
	'Path to XVBA Third Modules Files
	Dim xvbaModulesPath
	'' XVBA Modules Path
	xvbaModulesPath = VSCODE_ENV("WorkspacePath") & "\" & WORKSPACE_VBA_MODULES_FOLDER_NAME
	

	'Path to XVBA Unit Test Files
	Dim xvbaUnitTestPath 
	'' XVBA Unit Test files Path
	xvbaUnitTestPath = VSCODE_ENV("WorkspacePath") & "\" & WORKSPACE_VBA_TEST_FOLDER_NAME

	
	Dim book
	Set book =  GetObject(VSCODE_ENV("WorkspacePath") & "\" & VSCODE_ENV("WorkbookName")).Application.Workbooks(VSCODE_ENV("WorkbookName"))

	' Freeze Excel while importing all VBA modules (prevents redraw and recalculation on every Import call)
	Dim xlApp
	Set xlApp = book.Application
	xlApp.ScreenUpdating = False
	xlApp.Calculation    = -4135  ' xlCalculationManual
	xlApp.EnableEvents   = False

	'Delete all modules/Userforms from the ActiveWorkbook
	Call DeleteVBAModulesAndUserForms(book)
	''' Import all user VBA files
	Call ImportFiles(objFSO.GetFolder(szImportPath), book)
	''' Import all Thir Packages VBA files from xvba_modules folder
	Call ImportFiles(objFSO.GetFolder(xvbaModulesPath), book)
	If prop = "Production" Then
	 ''' Import all XVBA Unit Test Files from xvba_unit_test folder in Production
	 Call ImportFiles(objFSO.GetFolder(xvbaUnitTestPath), book)
	End If

	Call LogMessages(1,LOG_FILE_PATH, "Import VBA Files Finish")

	' Restore Excel state � do NOT call xlApp.Calculate;
	' Excel will recalculate naturally when it next needs to.
	xlApp.EnableEvents   = True
	xlApp.Calculation    = -4105  ' xlCalculationAutomatic
	xlApp.ScreenUpdating = True
	Set xlApp = Nothing


End Function





'/*
'This function read a XML file and get
'VSCode Workspace path and the controller to run in Excel
'@create: Alberto Eduardo
'@Since: 18/04/2020
'@Version: 0.1
'@Return getConfig: Scripting.Dictionary
'*/
Public Function getActionFromXML() 
	

	'Declare Config Dictionary
	Dim ConfigDict
	Set ConfigDict = CreateObject("Scripting.Dictionary")
	
	'Declare XML Object
	Dim xmlObj 
	Set xmlObj = CreateObject("MSXML2.DOMDocument")
	
	Dim xmlPath
	xmlPath = CURRENT_DIR & "\" & XML_FILE_NAME
	'Set XML params
	xmlObj.async = True
	xmlObj.validateOnParse = False
	xmlObj.Load (xmlPath)

	'Set variable for get document Root Node
	Dim docNodes 
	'Get Root Node "Xvba"
	Set docNodes = xmlObj.SelectNodes("//Xvba")
	'Set varibles for parse document nodes
	Dim rootNode 
	Dim configNode 
	Dim controllerNode 
	
	For Each rootNode In docNodes
		For Each configNode In rootNode.childNodes
			ConfigDict.Add "WorkspacePath", configNode.SelectSingleNode("WorkspacePath").text
			ConfigDict.Add "WorkbookName", configNode.SelectSingleNode("WorkbookName").text
			ConfigDict.Add "State", configNode.SelectSingleNode("State").text
			ConfigDict.Add "FolderVBA", configNode.SelectSingleNode("FolderVBA").text
			ConfigDict.Add "Command", configNode.SelectSingleNode("Command").text
			ConfigDict.Add "Prop", configNode.SelectSingleNode("Prop").text
			' ChangedFiles node is optional (only present for RUN_CHANGED command)
			Dim changedFilesNode
			Set changedFilesNode = configNode.SelectSingleNode("ChangedFiles")
			If Not changedFilesNode Is Nothing Then
				ConfigDict.Add "ChangedFiles", changedFilesNode.text
			Else
				ConfigDict.Add "ChangedFiles", ""
			End If
			Set changedFilesNode = Nothing
		Next
	Next

	'Return Config Dictionary Values
	Set getActionFromXML = ConfigDict
	
End Function


'
'Import only the VBA modules whose files changed since last export
'Receives a pipe-separated list of workspace-relative file paths
'
Public Function ImportChangedModules(changedFiles)

	If changedFiles = "" Then
		Call LogMessages(1, LOG_FILE_PATH, "ImportChangedModules: No changed files received")
		Exit Function
	End If

	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	Dim book
	Set book = GetObject(VSCODE_ENV("WorkspacePath") & "\" & VSCODE_ENV("WorkbookName")).Application.Workbooks(VSCODE_ENV("WorkbookName"))

	' Freeze Excel while replacing individual modules (prevents redraw/recalc on each change)
	Dim xlAppCh
	Set xlAppCh = book.Application
	xlAppCh.ScreenUpdating = False
	xlAppCh.Calculation    = -4135  ' xlCalculationManual
	xlAppCh.EnableEvents   = False

	Dim cmpComponents
	Set cmpComponents = book.VBProject.VBComponents

	Dim regExName
	Set regExName = CreateObject("VBScript.RegExp")
	regExName.Global = False
	regExName.IgnoreCase = True
	regExName.pattern = "(Attribute)\s+(VB_Name)\s+(\=)\s+\""(\w+)"

	Dim files
	files = Split(changedFiles, "|")

	Dim i
	For i = 0 To UBound(files)
		Dim relPath
		relPath = Trim(files(i))
		If relPath <> "" Then
			Dim fullPath
			fullPath = VSCODE_ENV("WorkspacePath") & "\" & relPath

			If objFSO.FileExists(fullPath) Then
				Dim fileExtension
				fileExtension = LCase(objFSO.GetExtensionName(fullPath))

				If fileExtension = "bas" Then
					' Standard modules: remove old one and re-import
					Dim basName
					basName = objFSO.GetBaseName(fullPath)
					On Error Resume Next
					Dim basComp
					Set basComp = cmpComponents(basName)
					If Not basComp Is Nothing Then
						cmpComponents.Remove basComp
					End If
					Set basComp = Nothing
					On Error GoTo 0
					On Error Resume Next
					cmpComponents.Import fullPath
					Call ErrorHandler("ImportChangedModules .bas " & basName)
					On Error GoTo 0
					Call LogMessages(1, LOG_FILE_PATH, "MODULE IMPORTED: " & basName)

				ElseIf fileExtension = "frm" Then
					' UserForms: remove old one and re-import
					Dim frmName
					frmName = objFSO.GetBaseName(fullPath)
					On Error Resume Next
					Dim frmComp
					Set frmComp = cmpComponents(frmName)
					If Not frmComp Is Nothing Then
						cmpComponents.Remove frmComp
					End If
					Set frmComp = Nothing
					On Error GoTo 0
					On Error Resume Next
					Call ImportAndModifyForm(fullPath, book, objFSO.OpenTextFile(fullPath).ReadAll)
					Call ErrorHandler("ImportChangedModules .frm " & frmName)
					On Error GoTo 0
					Call LogMessages(1, LOG_FILE_PATH, "FORM IMPORTED: " & frmName)

				ElseIf fileExtension = "cls" Then
					' Class modules AND sheet/workbook objects (type 100)
					' Must inspect VB_Name attribute and replace code in-place for type 100
					Dim sourceFile
					Set sourceFile = objFSO.OpenTextFile(fullPath)
					Dim fileContent
					fileContent = sourceFile.ReadAll
					sourceFile.Close
					Set sourceFile = Nothing

					Dim matched
					Set matched = regExName.Execute(fileContent)

					If matched.Count >= 1 Then
						Dim classNameMatched
						classNameMatched = matched(0).SubMatches(3)

						On Error Resume Next
						Dim clsComp
						Set clsComp = cmpComponents.Item(classNameMatched)
						On Error GoTo 0

						If Not IsEmpty(clsComp) And Not clsComp Is Nothing Then
							' Component already exists (includes Sheet / ThisWorkbook type 100):
							' Replace code lines in-place � never call Remove+Import for type 100
							Call addCodeInExcelClassObject(clsComp, fileContent)
							Call LogMessages(1, LOG_FILE_PATH, "CLASS UPDATED: " & classNameMatched & " FROM FILE: " & fullPath)
						Else
							' New class that does not yet exist in the workbook
							On Error Resume Next
							cmpComponents.Import fullPath
							Call ErrorHandler("ImportChangedModules .cls new " & classNameMatched)
							On Error GoTo 0
							Call LogMessages(1, LOG_FILE_PATH, "CLASS IMPORTED (new): " & classNameMatched & " FROM FILE: " & fullPath)
						End If
						Set clsComp = Nothing
					Else
						' Could not parse VB_Name � fall back to plain import
						On Error Resume Next
						cmpComponents.Import fullPath
						Call ErrorHandler("ImportChangedModules .cls fallback " & fullPath)
						On Error GoTo 0
						Call LogMessages(0, LOG_FILE_PATH, "CLASS fallback import: " & fullPath)
					End If
					Set matched = Nothing
				End If
			Else
				Call LogMessages(0, LOG_FILE_PATH, "ImportChangedModules: File not found: " & fullPath)
			End If
		End If
	Next

	Set regExName = Nothing

	' Restore Excel state � do NOT call xlAppCh.Calculate;
	' Excel will recalculate naturally when it next needs to.
	xlAppCh.EnableEvents   = True
	xlAppCh.Calculation    = -4105  ' xlCalculationAutomatic
	xlAppCh.ScreenUpdating = True
	Set xlAppCh = Nothing

	Call LogMessages(1, LOG_FILE_PATH, "ImportChangedModules Finish")

End Function


'/*
'Delete All VBA modules
'
'@param[Excel.Workbook] book
'*/
Function DeleteVBAModulesAndUserForms(book) 
	
	'As VBIDE.VBProject
	Dim VBProj 
	Set VBProj = book.VBProject
	
	'As VBIDE.VBComponent
	Dim VBComp 
	
	For Each VBComp In VBProj.VBComponents
		If VBComp.Type = 100 Then
			'Thisworkbook or worksheet module
			'We do nothing
		Else
			VBProj.VBComponents.Remove VBComp
		End If
	Next 
	
End Function



'
'Import All VBA modules from files i VSCode workspace
'
'@param[Excel.Workbook] book
'
Public Sub ImportFiles(folder, book )
	
   'As VBIDE.VBComponents
	Dim cmpComponents
	Set cmpComponents = book.VBProject.VBComponents
	
	Dim objFSO	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Dim fileExtension
	Dim fileContent
	Dim Matched
	Dim vbComponent

    Dim regExRow 
    Set regExRow = CreateObject("VBScript.RegExp")
    regExRow.Global = False
	regExRow.IgnoreCase  = True
    regExRow.pattern = "(Attribute)\s+(VB_Name)\s+(\=)\s+\""(\w+)"

    Call LogMessages(1,LOG_FILE_PATH, "Importing VBA Files from: " & folder)
	
	'A class module
 	  vbext_ct_ClassModule = 2
     'A document module like ThisWorkbook or Sheet1
	  vbext_ct_Document = 100
	 'A standard code module
      vbext_ct_ClassModule = 2
	 'A UserForm module
      vbext_ct_MSForm = 3


	For Each objFile In objFSO.GetFolder(folder).Files
		
		fileExtension = objFSO.GetExtensionName(objFile.Name)
		filePath = objFile.Path
		Set sourceFile  = objFSO.OpenTextFile(objFile.Path)
		fileContent = sourceFile.ReadAll
	
		

	
		If  (fileExtension = "bas" ) Then
		
   	     	Call cmpComponents.Import(filePath)
		  	Call LogMessages(1,LOG_FILE_PATH, "MODULE IMPORTED: " & objFile.Name & " FROM FILE: " & filePath & " TO EXCEL")

		ElseIf(fileExtension = "frm")then
		
			
		  Call ImportAndModifyForm(filePath, book,fileContent)
          Call LogMessages(1,LOG_FILE_PATH, "FORM IMPORTED: " & objFile.Name & " FROM FILE: " & filePath & " TO EXCEL")
		
		ElseIf (fileExtension = "cls" ) Then
		
			Set Matched = regExRow.Execute(fileContent)
	        classNameMached = Matched(0).SubMatches(3)

			On Error Resume Next
			Set vbComponent = cmpComponents.Item(classNameMached)
			On Error GoTo 0

			If Not IsEmpty(vbComponent) Then	
				
				Call addCodeInExcelClassObject(vbComponent,fileContent)
				Call LogMessages(1,LOG_FILE_PATH, "CLASS UPDATED: " & classNameMached & " FROM FILE: " & filePath & "" )
			Else
				Call LogMessages(1,LOG_FILE_PATH, "CLASS IMPORTED: " & classNameMached & " FROM FILE: " & filePath & "")		
				Call cmpComponents.Import (filePath)
			End If


		End If

		sourceFile.Close
	Next 
	
	On Error Resume Next

	For Each Subfolder In folder.SubFolders
		Call ImportFiles(Subfolder, book)
	Next
	

	Set folder = Nothing
	Set objFile = Nothing
	Set Subfolder = Nothing
	Set sourceFile = Nothing
	Set Matched = Nothing
	Set regExRow = Nothing
	
End Sub

Function ImportAndModifyForm(filePath, book, fileContent)
    ' Import the form into the VBProject
    Dim cmpComponent
    Set cmpComponent = book.VBProject.VBComponents.Import(filePath)

    ' Clear the existing code from the module
    Dim totalLines
    totalLines = cmpComponent.CodeModule.CountOfLines
    cmpComponent.CodeModule.DeleteLines 1, totalLines

    ' Define the lines to be excluded
    Dim excludeStart, excludeEnd
    excludeStart = "VERSION"
    excludeEnd = "Attribute VB_Exposed = False"

    ' Remove the excluded content from the file content
    Dim startIdx, endIdx
    startIdx = InStr(1, fileContent, excludeStart, vbTextCompare)
    endIdx = InStrRev(fileContent, excludeEnd, -1, vbTextCompare)

    If startIdx > 0 And endIdx > 0 And endIdx > startIdx Then
        Dim modifiedContent
        modifiedContent = Left(fileContent, startIdx - 1) & Mid(fileContent, endIdx + Len(excludeEnd) + 2)

        ' Remove leading blank lines
        Dim regExRemoveLeading
        Set regExRemoveLeading = CreateObject("VBScript.RegExp")
        regExRemoveLeading.Pattern = "^\s*\r?\n?"
        modifiedContent = regExRemoveLeading.Replace(modifiedContent, "")

		' Add two line breaks at the beginning
        modifiedContent = vbCrLf & vbCrLf & modifiedContent

        ' Remove trailing blank lines
        Dim regExRemoveTrailing
        Set regExRemoveTrailing = CreateObject("VBScript.RegExp")
        regExRemoveTrailing.Pattern = "\r?\n?\s*$"
        modifiedContent = regExRemoveTrailing.Replace(modifiedContent, "")

        ' Insert the new code into the code module
        If Len(modifiedContent) > 0 Then
            cmpComponent.CodeModule.InsertLines 1, modifiedContent
        End If
    End If

    ' Cleanup
    Set cmpComponent = Nothing
End Function


Public Function addCodeInExcelClassObject(cmpComponent, fileContent)

	' Declare variables.
	Dim regExRow
	Set regExRow = CreateObject("VBScript.RegExp")
	regExRow.Global = False
	regExRow.IgnoreCase  = True
	regExRow.pattern = "^(s*VERSION)[\s\S]+(VB_Exposed\s=\s*)(True|False)"

	fileContentedCleared = regExRow.Replace(fileContent, "")


    ' Split the fileContentedCleared into lines.
    Dim lines
    lines = Split(fileContentedCleared, vbCrLf)

    ' Remove the first and last blank lines if they exist.
    If UBound(lines) >= 0 Then
        If Len(lines(0)) = 0 Then
            lines = RemoveElementFromArray(lines, 0)
        End If

        If Len(lines(UBound(lines))) = 0 Then
            lines = RemoveElementFromArray(lines, UBound(lines))
        End If
    End If

    ' Delete the existing code from the code module.
    totalLines = cmpComponent.CodeModule.CountOfLines
    cmpComponent.CodeModule.DeleteLines 1, totalLines

    ' Insert the new code into the code module.
    If UBound(lines) >= 0 Then
        cmpComponent.CodeModule.InsertLines 1, Join(lines, vbCrLf)
    End If

    ' Clean up objects.
    Set cmpComponent = Nothing
    Set regExRow = Nothing


End Function


Function RemoveElementFromArray(arr, index)
    Dim tempArr
    Dim i

    ReDim tempArr(UBound(arr) - 1)
    For i = 0 To UBound(arr) - 1
        If i < index Then
            tempArr(i) = arr(i)
        Else
            tempArr(i) = arr(i + 1)
        End If
    Next

    RemoveElementFromArray = tempArr
End Function



'Log Messages
'@author Alberto Eduardo
'@since 21/04/2020
Public Function LogMessages(messageType, filePath , message )
	

	Dim PREFIX ,data
	PREFIX = Now & " - "
	
	Dim objFileToWrite
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(filePath,8,True)
	
	Select Case messageType
		
		Case 0 
		data = PREFIX & "Error:" & message
		Case 1 
		data = PREFIX & "DEBUG:" & message
		Case Else 
		data = PREFIX & "INFO:" & message
	End Select
	objFileToWrite.WriteLine(data)
	
	objFileToWrite.Close
	Set objFileToWrite = Nothing
	
End Function



