Attribute VB_Name = "ModuleSlideLibrary"
'MIT License

'Copyright (c) 2021 - 2026 iappyx

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Public Function GetSlideLibraryFilePath(Optional ByVal ShowSettingsIfMissing As Boolean = True) As String
    Dim LibraryFilePath As String
    LibraryFilePath = Trim$(GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", ""))

    If LibraryFilePath = "" Then
        If ShowSettingsIfMissing Then
            MsgBox "No slide library file set. Please set the file in Instrumenta settings."
            SettingsForm.Show
        End If
        Exit Function
    End If

    If Dir$(LibraryFilePath) = "" Then
        MsgBox "The slide library file could not be found:" & vbCrLf & vbCrLf & LibraryFilePath & vbCrLf & vbCrLf & "Please choose the slide library file again in Instrumenta settings."
        If ShowSettingsIfMissing Then SettingsForm.Show
        Exit Function
    End If

    GetSlideLibraryFilePath = LibraryFilePath
End Function

Public Function OpenSlideLibraryPresentation(Optional ByVal ShowWindow As Boolean = False, Optional ByVal OpenReadOnly As Boolean = True) As PowerPoint.Presentation
    Dim LibraryFilePath As String
    LibraryFilePath = GetSlideLibraryFilePath(True)
    If LibraryFilePath = "" Then Exit Function

    On Error GoTo OpenFailed

    #If Mac Then
        Set OpenSlideLibraryPresentation = PowerPoint.Presentations.Open(LibraryFilePath)
    #Else
        Dim ReadOnlyState As MsoTriState
        Dim WindowState As MsoTriState

        If OpenReadOnly Then
            ReadOnlyState = msoTrue
        Else
            ReadOnlyState = msoFalse
        End If

        If ShowWindow Then
            WindowState = msoTrue
        Else
            WindowState = msoFalse
        End If

        Set OpenSlideLibraryPresentation = PowerPoint.Presentations.Open(LibraryFilePath, ReadOnlyState, msoFalse, WindowState)
    #End If

    Exit Function

OpenFailed:
    MsgBox "Instrumenta could not open the slide library file:" & vbCrLf & vbCrLf & LibraryFilePath & vbCrLf & vbCrLf & "PowerPoint said: " & Err.Description & vbCrLf & vbCrLf & "Please check that the file exists, is not blocked, and is not already open in a conflicting state."
End Function

Sub OpenSlideLibraryFile()

    Dim LibraryPresentation       As PowerPoint.Presentation
    Set LibraryPresentation = OpenSlideLibraryPresentation(True, False)


End Sub


Sub AddSelectedSlidesToLibraryFile()

Dim LibraryPresentation       As PowerPoint.Presentation
Set MyDocument = Application.ActiveWindow

If GetSlideLibraryFilePath(True) <> "" Then
    
    If Application.ActiveWindow.Selection.SlideRange.count > 0 Then
        
        Set LibraryPresentation = OpenSlideLibraryPresentation(True, False)
        If LibraryPresentation Is Nothing Then Exit Sub
        
        MyDocument.Selection.SlideRange.Copy
        
        LibraryPresentation.Application.CommandBars.ExecuteMso ("PasteSourceFormatting")
                
        LibraryPresentation.SaveCopyAs GetSetting("Instrumenta", "SlideLibrary", "SlideLibraryFile", "")
        LibraryPresentation.Close
        
        Set LibraryPresentation = Nothing
        
    Else
        
        MsgBox "No slides selected."
        
    End If
    
End If


End Sub
