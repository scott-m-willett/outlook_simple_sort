Public Sub MoveSenderByName()

    On Error Resume Next

    ' Various objects required to work with emails and folders
    Dim Explorer As Outlook.Explorer
    Dim Namespace As Outlook.Namespace
    Dim myFolders As Folder
    Set Explorer = Application.ActiveExplorer
    Set Namespace = Application.GetNamespace("MAPI")
    
    ' Set the default folder where sender folders exist
    Set myFolders = Namespace.GetDefaultFolder(olFolderInbox).Folders("Reference").Folders("By Sender")

    ' For each item selected in the inbox, format their name, create a folder if it doesn't exist, and move the mailitem into it
    For Each Item In Explorer.Selection
        Sender = Item.SenderName
        If (InStr(1, Sender, " ")) Then
            FirstName = Left(Sender, InStr(1, Sender, " ") - 1)
            LastName = Mid(Sender, InStr(1, Sender, " ") + 1, Len(Sender))
            FolderFormat = LastName & ", " & FirstName
            myFolders.Folders.Add (FolderFormat)
            Item.Move (myFolders.Folders(FolderFormat))
        End If
    Next
    
End Sub

Public Sub MoveToGroupFolder()

    On Error Resume Next

    ' Various objects required to work with emails and folders
    Dim Explorer As Outlook.Explorer
    Dim Namespace As Outlook.Namespace
    Dim myFolders As Folder
    Set Explorer = Application.ActiveExplorer
    Set Namespace = Application.GetNamespace("MAPI")
    
    ' Set the default folder where sender folders exist
    Set myFolders = Namespace.GetDefaultFolder(olFolderInbox).Folders("Reference").Folders("By Group")

    Group = InputBox("Input group name")

    ' Ask for a name, create the folder if it doesn't exist based of this, then place into that name group
    For Each Item In Explorer.Selection
        
        myFolders.Folders.Add (Group)
        Item.Move (myFolders.Folders(Group))
    Next
    
End Sub

Public Sub MoveToPurposeFolder()

    On Error Resume Next

    ' Various objects required to work with emails and folders
    Dim Explorer As Outlook.Explorer
    Dim Namespace As Outlook.Namespace
    Dim myFolders As Folder
    Set Explorer = Application.ActiveExplorer
    Set Namespace = Application.GetNamespace("MAPI")
    
    ' Set the default folder where sender folders exist
    Set myFolders = Namespace.GetDefaultFolder(olFolderInbox).Folders("Reference").Folders("By Purpose")

    Purpose = InputBox("Input purpose of the email")

    ' Ask for a name, create the folder if it doesn't exist based of this, then place into that folder
    For Each Item In Explorer.Selection
        myFolders.Folders.Add (Purpose)
        Item.Move (myFolders.Folders(Purpose))
    Next
    
End Sub