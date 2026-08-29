VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "Xlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'namespace=xvba_modules\Xlog



'/*
'
'Log Messages
'@author Alberto Eduardo
'@since 21/04/2020
'
'*/

Private param As String

Private Sub class_initialize()
  param = Application.ThisWorkbook.Path  & "\app.log"
End Sub


'/*
'
'Set Log Message
'
'@param[Integer] messageType 
'@param["String"] messageText
'@param[String] filePath
'*/
Public Function message(messageType As Integer, messageText As String,Optional filePath)
 
  If IsMissing(filePath) Then
    filePath = param
  End if 
  Dim FileNum As Integer
  
  FileNum = FreeFile
  
  Dim PREFIX As String
  
  PREFIX = Now & " - "
  
  
  Open filePath For Append As #FileNum
  
  Select Case messageType
   
    Case 0 'Error Message
    Print #FileNum, PREFIX & "Error:" & messageText
    Case 1 'Success
    Print #FileNum, PREFIX & "DEBUG:" & messageText
    Case Else 'No Type Set
    Print #FileNum, PREFIX & "INFO:" & messageText
  End Select
     
  Close #FileNum
  
  
  
End Function
  
  
  
  
Public Function ErrorHandller(filePath As String, macro As String)

  Dim messageLog As String
  messageLog = Err.Description & "  [" & Err.Number & "] Macro:" & macro

  ' Xlog.message is the logger exposed by this XVBA package.  The old
  ' LogMessages call belonged to the XVBA server script and is not a VBA
  ' procedure, so it causes a compile error in the imported class.
  Call Xlog.message(0, messageLog, filePath)

  ' Preserve the package's original fail-safe behavior after logging.
  ThisWorkbook.Close SaveChanges:=False

End Function
  
