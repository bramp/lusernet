Attribute VB_Name = "modFrmMain"
Option Explicit

Public Sub addShare(ByVal newShare As String)
    If newShare <> "" And FolderExists(newShare) Then

        frmMain.StatusChange "Scanning '" & newShare & "' for new files"

        Dim path As String
        path = newShare
        Dim folders() As String
        Dim folderName As String
        Dim shareID As Long
        If Right(newShare, 1) = "\" Then newShare = Left(newShare, Len(newShare) - 1) 'Removes trailing slash
        folders = Split(newShare, "\")
        folderName = folders(UBound(folders))
        If folderName = "" Then folderName = folders(UBound(folders) - 1)

        newShare = Replace(newShare, "'", "''")
        Set rs = conn.Execute("SELECT folderID FROM tblFolders WHERE folderName = '" & newShare & "' AND rootID=-1")
        'Check if share is in database
        If rs.EOF Or rs.BOF Then
            'Create Share
            Set rs = conn.Execute("INSERT INTO tblFolders (rootID, folderName, lastUpdate) VALUES (-1,'" & newShare & "', NOW())")
            Set rs = conn.Execute("SELECT folderID FROM tblFolders WHERE folderName = '" & newShare & "' AND rootID=-1")
            shareID = rs("folderID")
            rs.Close
        Else
            shareID = rs("folderID")
            rs.Close
            'Update Share
            Set rs = conn.Execute("UPDATE tblFolders SET lastUpdate=NOW() WHERE folderID=" & shareID)
        End If

        Set rs = Nothing

        'Now Iterate through files/folders
        GetFiles path, shareID
        GetFolders path, shareID

        frmMain.StatusChange "Done scan"
        frmMain.ShareRefreshTree
    Else
        frmMain.StatusChange "Unable to refresh '" & newShare & "', the folder was not found"
    End If
End Sub

'This resizes everything etc
Public Sub DoGUIStuff()

    With frmMain
          
        .status.Top = .ScaleHeight - .status.Height - 2
        .status.Left = 5
        .status.Width = .ScaleWidth - 8
        '.status.Height = .status.Height + 2
        .shpStatus.Top = .status.Top - 2
        .shpStatus.Left = .status.Left - 2
        .shpStatus.Width = .status.Width + 2
        .shpStatus.Height = .status.Height + 1

        .miniGraph.Height = .status.Height - 4
        .miniGraph.Top = .status.Top
        .miniGraph.Left = .status.Width - .miniGraph.Width + 2
        .miniGraph.BackColor = .shpStatus.FillColor

        .shpGraph.Left = .miniGraph.Left - 1
        .shpGraph.Top = .miniGraph.Top - 1
        .shpGraph.Width = .miniGraph.Width + 2
        .shpGraph.Height = .miniGraph.Height + 2

        .StatusChange "LUSerNet version " & App.Minor & "." & App.Revision
        .tmrMessage.Interval = 1
        .lblTitle.Caption = .lblTitle.Caption & " v" & App.Minor & "." & App.Revision
        MakeRounded frmMain

        PopulateSearch

        'Skin the app
        .Skinner.Repaint frmMain
    End With
End Sub

Public Sub LoadSettings()
    frmMain.txtDownloadLocation.Text = GetSetting("LUSerNet", "Main", "DownloadLocation", "")
    If Trim(frmMain.txtDownloadLocation.Text) = "" Then
        frmMain.txtDownloadLocation.Text = App.path
    End If
    
    frmMain.UpDown(0).Value = GetSetting("LUSerNet", "Main", "UploadLimit", 5)
    If frmMain.UpDown(0).Value <= 0 Then frmMain.UpDown(0).Value = 5
    frmMain.UpDown(1).Value = frmMain.UpDown(0).Value
    frmMain.txtUploadTotal.Text = frmMain.UpDown(0).Value
    frmMain.txtDownloadTotal.Text = frmMain.UpDown(0).Value
End Sub

Public Sub SaveSettings()
    SaveSetting "LUSerNet", "Main", "DownloadLocation", frmMain.txtDownloadLocation.Text
    SaveSetting "LUSerNet", "Main", "UploadLimit", frmMain.UpDown(0).Value
    SaveSetting "LUSerNet", "Main", "Java", "import java.lancs.*"
End Sub

'Populates Search combo
Public Sub PopulateSearch()
    frmMain.cmbSearch.AddItem "DivX Movies (*.avi)"
    frmMain.cmbSearch.AddItem "MPeg Movies (*.mpg)"
    frmMain.cmbSearch.AddItem "ASF Movies (*.asf)"
    frmMain.cmbSearch.AddItem "MP3 Music (*.mp3)"
    frmMain.cmbSearch.AddItem "WMA Music (*.wma)"
End Sub

