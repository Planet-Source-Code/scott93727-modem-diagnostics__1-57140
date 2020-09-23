Attribute VB_Name = "Module1"
' Public variables
Public Echo As Boolean        ' Echo On/Off flag.
Public CancelSend As Integer  ' Flag to stop sending a text file.

Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'*** Global module for MDI Notepad sample application.  ***
'**********************************************************
Option Explicit

' User-defined type to store information about child forms
Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type

Public FState As FormState              ' Array of user-defined types
Public gFindString As String            ' Holds the search text.
Public gFindCase As Integer             ' Key for case sensitive search
Public gFindDirection As Integer        ' Key for search direction.
Public gCurPos As Integer               ' Holds the cursor location.
Public gFirstTime As Integer            ' Key for start position.
Public Const ThisApp = "ModemMaster"        ' Registry App constant.
Public Const ThisKey = "Recent Files"   ' Registry Key constant.
Sub FileNew()
    Dim intResponse As Integer
    
    ' If the file has changed, save it
    If FState.Dirty = True Then
        intResponse = FileSave
        If intResponse = False Then Exit Sub
    End If
    ' Clear the textbox and update the caption.
    frmTerminal.txtTerm.Text = ""
    frmTerminal.Caption = "Untitled"
End Sub
Function FileSave() As Integer
    Dim strFilename As String

    If frmTerminal.Caption = "Untitled" Then
        ' The file hasn't been saved yet.
        ' Get the filename, and then call the save procedure, GetFileName.
        strFilename = GetFileName(strFilename)
    Else
        ' The form's Caption contains the name of the open file.
        strFilename = Right(frmTerminal.Caption, Len(frmTerminal.Caption) - 14)
    End If
    ' Call the save procedure. If Filename = Empty, then
    ' the user chose Cancel in the Save As dialog box; otherwise,
    ' save the file.
    If strFilename <> "" Then
        SaveFileAs strFilename
        FileSave = True
    Else
        FileSave = False
    End If
End Function
Sub FindIt()
    Dim intStart As Integer
    Dim intPos As Integer
    Dim strFindString As String
    Dim strSourceString As String
    Dim strMsg As String
    Dim intResponse As Integer
    Dim intOffset As Integer
    
    ' Set offset variable based on cursor position.
    If (gCurPos = frmTerminal.txtTerm.SelStart) Then
        intOffset = 1
    Else
        intOffset = 0
    End If

    ' Read the public variable for start position.
    If gFirstTime Then intOffset = 0
    ' Assign a value to the start value.
    intStart = frmTerminal.txtTerm.SelStart + intOffset
        
    ' If not case sensitive, convert the string to upper case
    If gFindCase Then
        strFindString = gFindString
        strSourceString = frmTerminal.txtTerm.Text
    Else
        strFindString = UCase(gFindString)
        strSourceString = UCase(frmTerminal.txtTerm.Text)
    End If
            
    ' Search for the string.
    If gFindDirection = 1 Then
        intPos = InStr(intStart + 1, strSourceString, strFindString)
    Else
        For intPos = intStart - 1 To 0 Step -1
            If intPos = 0 Then Exit For
            If Mid(strSourceString, intPos, Len(strFindString)) = strFindString Then Exit For
        Next
    End If

    ' If the string is found...
    If intPos Then
        frmTerminal.txtTerm.SelStart = intPos - 1
        frmTerminal.txtTerm.SelLength = Len(strFindString)
    Else
        strMsg = "Cannot find " & Chr(34) & gFindString & Chr(34)
        intResponse = MsgBox(strMsg, 0, App.Title)
    End If
    
    ' Reset the public variables
    gCurPos = frmTerminal.txtTerm.SelStart
    gFirstTime = False
End Sub
Sub GetRecentFiles()
    ' This procedure demonstrates the use of the GetAllSettings function,
    ' which returns an array of values from the Windows registry. In this
    ' case, the registry contains the files most recently opened.  Use the
    ' SaveSetting statement to write the names of the most recent files.
    ' That statement is used in the WriteRecentFiles procedure.
    Dim i As Integer
    Dim varFiles As Variant ' Varible to store the returned array.
    
    ' Get recent files from the registry using the GetAllSettings statement.
    ' ThisApp and ThisKey are constants defined in this module.
    If GetSetting(ThisApp, ThisKey, "RecentFile1") = Empty Then Exit Sub
    
    varFiles = GetAllSettings(ThisApp, ThisKey)
    
    For i = 0 To UBound(varFiles, 1)
        frmTerminal.mnuRecentFile(0).Visible = True
        frmTerminal.mnuRecentFile(i + 1).Caption = varFiles(i, 1)
        frmTerminal.mnuRecentFile(i + 1).Visible = True
    Next i
End Sub
Sub ResizeNote()
    ' Expand text box to fill the form's internal area.
    If frmTerminal.tbrToolBar.Visible Then
        frmTerminal.txtTerm.Height = frmTerminal.ScaleHeight - frmTerminal.tbrToolBar.Height
        frmTerminal.txtTerm.Width = frmTerminal.ScaleWidth
        frmTerminal.txtTerm.Top = frmTerminal.tbrToolBar.Height
    Else
        frmTerminal.txtTerm.Height = frmTerminal.ScaleHeight
        frmTerminal.txtTerm.Width = frmTerminal.ScaleWidth
        frmTerminal.txtTerm.Top = 0
    End If
End Sub
Sub WriteRecentFiles(OpenFileName)
    ' This procedure uses the SaveSettings statement to write the names of
    ' recently opened files to the System registry. The SaveSetting
    ' statement requires three parameters. Two of the parameters are
    ' stored as constants and are defined in this module.  The GetAllSettings
    ' function is used in the GetRecentFiles procedure to retrieve the
    ' file names stored in this procedure.
    
    Dim i As Integer
    Dim strFile As String
    Dim strKey As String

    ' Copy RecentFile1 to RecentFile2, and so on.
    For i = 3 To 1 Step -1
        strKey = "RecentFile" & i
        strFile = GetSetting(ThisApp, ThisKey, strKey)
        If strFile <> "" Then
            strKey = "RecentFile" & (i + 1)
            SaveSetting ThisApp, ThisKey, strKey, strFile
        End If
    Next i
  
    ' Write the open file to first recent file.
    SaveSetting ThisApp, ThisKey, "RecentFile1", OpenFileName
End Sub
'*** Standard module with procedures for working with   ***
'*** files. Part of the SDI Notepad sample application. ***
'**********************************************************
Sub FileOpenProc()
    Dim intRetVal
    Dim intResponse As Integer
    Dim strOpenFileName As String
    
    ' If the file has changed, save it
    If FState.Dirty = True Then
        intResponse = FileSave
        If intResponse = False Then Exit Sub
    End If
    On Error Resume Next
    
    frmTerminal.OpenLog.filename = ""
    frmTerminal.OpenLog.ShowOpen
    If Err <> 32755 Then    ' User chose Cancel.
        strOpenFileName = frmTerminal.OpenLog.filename
        ' If the file is larger than 65K, it can't
        ' be opened, so cancel the operation.
        If FileLen(strOpenFileName) > 65000 Then
            MsgBox "The file is too large to open."
            Exit Sub
        End If
        
        OpenFile (strOpenFileName)
        UpdateFileMenu (strOpenFileName)
    End If
End Sub
Function GetFileName(filename As Variant)
    ' Display a Save As dialog box and return a filename.
    ' If the user chooses Cancel, return an empty string.
    On Error Resume Next
    frmTerminal.OpenLog.filename = filename
    frmTerminal.OpenLog.ShowSave
    If Err <> 32755 Then    ' User chose Cancel.
        GetFileName = frmTerminal.OpenLog.filename
    Else
        GetFileName = ""
    End If
End Function
Function OnRecentFilesList(filename) As Integer
    Dim i         ' Counter variable.

    For i = 1 To 5
        If frmTerminal.mnuRecentFile(i).Caption = filename Then
            OnRecentFilesList = True
            Exit Function
        End If
    Next i
    OnRecentFilesList = False
End Function
Sub OpenFile(filename)
    Dim fIndex As Integer
    
    On Error Resume Next
    ' Open the selected file.
    Open filename For Input As #1
    If Err Then
        MsgBox "Can't open file: " + filename
        Exit Sub
    End If
    ' Change the mouse pointer to an hourglass.
    Screen.MousePointer = 11
    
    ' Change the form's caption and display the new text.
    frmTerminal.Caption = "SDI NotePad - " & UCase(filename)
    frmTerminal.txtTerm.Text = Input(LOF(1), 1)
    FState.Dirty = False
    Close #1
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
End Sub
Sub SaveFileAs(filename)
    On Error Resume Next
    Dim strContents As String

    ' Open the file.
    Open filename For Output As #1
    ' Place the contents of the notepad into a variable.
    strContents = frmTerminal.txtTerm.Text
    ' Display the hourglass mouse pointer.
    Screen.MousePointer = 11
    ' Write the variable contents to a saved file.
    Print #1, strContents
    Close #1
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
    ' Set the form's caption.
    If Err Then
        MsgBox Error, 48, App.Title
    Else
        frmTerminal.Caption = filename
        ' Reset the dirty flag.
        FState.Dirty = False
    End If
End Sub
Sub UpdateFileMenu(filename)
        Dim intRetVal As Integer
        ' Check if the open filename is already in the File menu control array.
        intRetVal = OnRecentFilesList(filename)
        If Not intRetVal Then
            ' Write open filename to the registry.
            WriteRecentFiles (filename)
        End If
        ' Update the list of the most recently opened files in the File menu control array.
        GetRecentFiles
End Sub
 Sub Savefilegui()
    Dim strSaveFileName As String
    Dim strDefaultName As String
    
    ' Assign the form caption to the variable.
    strDefaultName = Right$(frmTerminal.Caption, Len(frmTerminal.Caption) - 14)
    If frmTerminal.Caption = "Untitled" Then
        ' The file hasn't been saved yet.
        ' Get the filename, and then call the save procedure, strSaveFileName.
        
        strSaveFileName = GetFileName("Untitled.txt")
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
        ' Update the list of recently opened files in the File menu control array.
        UpdateFileMenu (strSaveFileName)
    Else
        ' The form's Caption contains the name of the open file.
        strSaveFileName = GetFileName(strDefaultName)
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
        ' Update the list of recently opened files in the File menu control array.
        UpdateFileMenu (strSaveFileName)
    End If
End Sub

