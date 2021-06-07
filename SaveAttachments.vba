Public strFolderpath As String

Public Sub SaveToFolder()
                    '''''''''''''''''''''''''''''''''''''''
                    ' MAKE SURE TO CREATE A FOLDER FIRST
                    ' THEN UPDATE THE PATH
                    ' READ.ME WILL HAVE FURTHER INSTRUCTIONS
                    '''''''''''''''''''''''''''''''''''''''
strFolderpath = "PATH TO FOLDER GOES HERE”
SaveAttachments
End Sub

Private Sub SaveAttachments()
Dim objOL As Outlook.Application
Dim objMsg As Object
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String

    On Error Resume Next
    Set objOL = Application

    Set objSelection = objOL.ActiveExplorer.Selection

    ' Check each selected item for attachments.
    For Each objMsg In objSelection

    Set objAttachments = objMsg.Attachments
    lngCount = objAttachments.Count
        
    If lngCount > 0 Then
    
    ' Use a count down loop for removing items
    ' from the collection. Otherwse, the loop counter gets
    ' confused and only %2 is removed.
    
    For i = lngCount To 1 Step -1
    
    ' Get the file name.
    strFile = objAttachments.Item(i).FileName
    
    ' Combine with the path to the folder.
    strFile = strFolderpath & strFile
    
    ' Save the attachment as a file.
    objAttachments.Item(i).SaveAsFile strFile
    
    Next i
    End If
    
    Next
    
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
End Sub

