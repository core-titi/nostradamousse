'This function is from http://jmerrell.com/2011/05/21/outlook-macros-move-email/
'Author : Jim Merrell
Function GetFolder(ByVal FolderPath As String) As Outlook.folder
    Dim TestFolder As Outlook.folder
    Dim FoldersArray As Variant
    Dim i As Integer
    
    On Error GoTo GetFolder_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set TestFolder = Application.Session.Folders.Item(FoldersArray(0))
    
    If Not TestFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = TestFolder.Folders
            Set TestFolder = SubFolders.Item(FoldersArray(i))
            
            If TestFolder Is Nothing Then
                Set GetFolder = Nothing
            End If
        Next
    End If
    
    'Return the TestFolder
    Set GetFolder = TestFolder
    Exit Function

GetFolder_Error:
    Set GetFolder = Nothing
    Exit Function
End Function
'end of Jim Merrell code


Sub GetFoldersList(fList As String)
    On Error GoTo On_Error
     
    'Dim listInString As String
    Dim folder As Outlook.folder
        
    'get all mailbox storages
    For Each folder In Application.Session.Folders
        Call RecurseFolders(folder, fList)
    Next
     
Exiting:
    Exit Sub
On_Error:
    MsgBox "error=" & Err.Number & " " & Err.Description
End Sub

Sub RecurseFolders(CurrentFolder As Outlook.folder, list As String)
    Dim SubFolder As Outlook.folder

    'just use ,,, as delimiter
    list = list & ",,," & CurrentFolder.FolderPath
     
    'recursive process to get children folders
    For Each SubFolder In CurrentFolder.Folders
        Call RecurseFolders(SubFolder, list)
    Next SubFolder
End Sub

Sub UpdateFolderList()
    On Error GoTo On_Error
    
    'folder list in string
    Dim flString As String
    
    'folder list in array
    Dim fList() As String
    
    'get folder list in string
    Call GetFoldersList(flString)
    
    'transforme in array
    fList() = Split(flString, ",,,")
    
    'empty the folder list input
    folderList.Clear
    
    'find the items with search string
    For Each it In fList
    
        If CaseSensible.Value = False Then 'ignore case
        
            If InStr(1, LCase(it), LCase(searchbox.Value)) <> 0 Then
                folderList.AddItem it
            End If
            
        Else 'with case sensible
        
            If InStr(1, it, searchbox.Value) <> 0 Then
                folderList.AddItem it
            End If
            
        End If
        
    Next
    
    'unique folder item with be put automatically in searchbox
    If folderList.ListCount = 1 Then
        folderList.Selected(0) = True
        searchbox.Value = folderList.Value
    End If
    
Exiting:
    Exit Sub
On_Error:
    MsgBox "error=" & Err.Number & " " & Err.Description
End Sub



Private Sub btMoveToFolder_Click()

    'no mail item selected
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox ("No item selected")
        Exit Sub
    End If
    
    'init base object
    Dim olExpl As Outlook.Explorer
    Dim olSel As Outlook.Selection
    Dim mailItem As Outlook.mailItem
    Dim destFolder As Outlook.folder
    
    'set object instance
    Set olExpl = Application.ActiveExplorer
    Set olSel = olExpl.Selection
    Set destFolder = GetFolder(searchbox.Value)
    
    'get item 1-by-1
    For i = 1 To olSel.Count
        
        'get mail item
        Set mailItem = olSel.Item(i)
        
        'move mail item in destination folder
        mailItem.Move destFolder
        
    Next i
    
    'end of the program
    End
    
End Sub

Private Sub folderList_Click()

    'put selected value in search text box
    searchbox.Value = folderList.Value
    'focus on search box to read the text in the right side
    searchbox.SetFocus
    
End Sub


Private Sub searchbox_Change()

    'clear listbox items if searchbox is empty
    If searchbox.Value = "" Then
        folderList.Clear
    Else
        UpdateFolderList
    End If

End Sub

