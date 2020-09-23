VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check For Update Form"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblAction 
      Alignment       =   2  'Center
      Caption         =   "Authenticating, please wait"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by Puddy Davidson - http://puddys-world.com

'This is my example of a supreme auto updater that handles the following:
    'can disable the program from future use, and explains why
    'checks for new versions and gives an option to download it, including update release notes
    'deletes the old version from users computer, and automatically opens the new version

'I give this to you all freely, and all i ask is for you to vote for it if you like it, comments would be nice also
'use of this example is given freely to all, you may modify it as you wish
'Again its the first time i have asked ppl to vote, so if you use it thats all i ask

'Example:
'I have compiled this exact source into version 1.5 and i have uploaded to my server called Update ExampleNEW.exe
'to test this auto updater as is with no changes, compile this project (version is set at 1.0)
'run this project, you will see there is a new version available (1.5 as is noted on my server)
'Enjoy guys :)


'declare some variables that we will use
Dim Size As Long, Remaining As Long, NowSize As Long
Dim ProgressReal As Integer, Chunk() As Byte
Dim FileName As String, updateText As String, contentsOf As String
Dim reason As String, newVer As String, curVer As String, updateString As String

Private Sub Form_Load()
On Error GoTo errHandler 'add simple error handling to catch any runtime error

        'run this to clean up after the updater ran (if it ran, doesnt matter)
    If Dir(App.Path & "\DeleteOLD.bat") <> "" Then 'exists
        Kill App.Path & "\DeleteOLD.bat" 'after the update, make sure the batch file is removed
    End If
        Me.Height = 990: Me.Width = 4260: lblAction.Caption = "Authenticating, please wait" 'enter the update form
        Me.Show 'show that the program is actually active
        
            'I will use my server to host these example files, I will keep them there for some time
            'You will need to make these files onto your own server in your apps
            
         'this is an extra, you can remove the disable feature if you like
        contentsOf = Inet1.OpenURL("http://puddys-programs.com/updateExample/exampleDisabled.txt")
    If InStr(contentsOf, "disabled") Then 'is disabled written in the txt doc?
        reason = Inet1.OpenURL("http://puddys-programs.com/updateExample/reasonDisabled.txt") 'so if it is disabled, give a reason why - you could add this to
                                                                                              'exampleDisabled.txt doc also to avoid a second connection
                                                                                              
        MsgBox ("Puddy has disabled this program" & vbCrLf & "Reason being:" & vbCrLf & vbCrLf & reason) 'tell the user its disabled, and why
        End 'terminate the program, you could run a batch file here to delete the program also - again totally optional
    End If
    
        'for this example we will be using only the major and minor of the programs internal version, you can add revision to your own projects/server
        'curVer is the string that holds the version of the program that is running this updater
        curVer = App.Major & "." & App.Minor '& "." & App.Revision
        
        'the most current version gets put into a txt document and uploaded here
        updateString = Inet1.OpenURL("http://puddys-programs.com/updateExample/newVersion.txt")
    
     'newVer is the string that holds the most current version number
    If updateString <> "" Then
        newVer = updateString
    End If
    
     'if new version is higher than the running version
    If newVer > curVer Then
    
         'update notes describe what updates have been done (optional)
        updateText = Inet1.OpenURL("http://puddys-programs.com/updateExample/updateText.txt")
         
         'declare the filname to download (must be on the server)
         'name it with NEW and let the batch file rename it later
         'you could hardcode this into the next line, or in a class you could pass this to it
        FileName = "Update ExampleNEW.exe"
        
         'grab some info about the update file - update size and progress details
        Inet1.Execute Trim("http://puddys-programs.com/updateExample/" & FileName), "GET"
    Do While Inet1.StillExecuting = True
        DoEvents
    Loop
        ProgressBar1.Max = 100
        Size = CLng(Inet1.GetHeader("Content-Length")) 'get the size of the update file
        lblSize.Caption = Size & "bytes"
        Remaining = Size
        NowSize = 0
        lblAction.Caption = "Download update? Yes or No": Me.Width = 7455: Me.Height = 1500
            
         'let the user know there is a update available and what the update contains, gives them an option to download or not
        If MsgBox("There is a new version available v " & newVer & vbCrLf & vbCrLf & "This update - " & updateText & vbCrLf & vbCrLf & "Would you like to download the new version?", vbYesNo, "Update available") = vbYes Then
                
            lblAction.Caption = "Downloading update, please wait"
         
         'donot change app.path, it must be downloaded to the same location for the batch file, important
        Open App.Path & "\" & FileName For Binary Access Write As #1
        Do Until Remaining = 0 'download untill finished
    
            'If the user cancels the update mid way through the download
            'then run a sequence of code, messagebox to tell them its aborted, remove the part downloaded file, close the check from and open the main form
            If frmCheck.Tag = "Cancel" Then Inet1.Cancel: Close #1: MsgBox "Update aborted": Kill App.Path & "\" & FileName: frmMain.Show: Unload Me: Exit Sub
        
                 'download in chunks, use the array to determine progress
                If Remaining > 1024 Then
                    Chunk = Inet1.GetChunk(1024, icByteArray)
                    Remaining = Remaining - 1024
                Else
                    Chunk = Inet1.GetChunk(Remaining, icByteArray) 'completing the download
                    Remaining = 0 'download finished, progress complete
                End If
                    
                NowSize = Size - Remaining
                ProgressReal = CInt((100 / Size) * NowSize) 'gets the current progress value
                ProgressBar1.Value = ProgressReal 'show the progress
                Me.Caption = ProgressReal & "%" & " - Downloaded" 'show the progress in title also
            Put #1, , Chunk
        Loop
        Close #1
        
                'tell the user the download is complete, and warn them of the next step
            MsgBox "The update is complete. Press ok to open the new version.", vbInformation, "Update Complete"
            
            Call deleteOLD 'create and run the batch file
                Exit Sub
     
         'if the user selected no to downloading the update, give them a message (optional)
        Else
        
                 'they didnt wana update, open the main program
                MsgBox "You choose not to update this time, you will be asked again next time you open this program.", vbInformation, "Update Pending"
                frmMain.Show
                Unload Me
                Exit Sub
        End If
    
     'the verion running is the same or higher than the current version we put in the online txt document
    Else
            frmMain.Show 'open main program
            Unload Me
            Exit Sub
    End If
    
errHandler:
        MsgBox Err.Description 'if there was an error, get a description and open the main program, debug purpose mostly
        frmMain.Show
        Unload Me
End Sub

Private Sub cmdCancel_Click()
         'cancel the update
        frmCheck.Tag = "Cancel"
End Sub

Private Sub deleteOLD()

        'create the batch file in the same directory as the old and new versions to make this batch smaller
    Open App.Path & IIf(Right(App.Path, 1) <> "\", "\DeleteOLD.bat", "DeleteOLD.bat") For Output As #1 'create the batch file
    
        'open the created batch file and print some commands into it, batch file will look like this
            
            '@Echo off
            ':S
            'Del "(this is the app exe name, we use this incase the user changed the exe name)"   <note: the quotation marks throughout this batch file are nesasary incase your exe name contains spaces>
            'If Exist "(app name again here)" Goto S   <so if its not deleted yet, go back to :S and read on>
            ':D
            'ren "Update ExampleNEW.exe" "Update Example.exe"   <use the batch to change the new version into the same name as the old version>
            'If Exist "Update ExampleNEW.exe" Goto D   <same as three lines above>
            'Update Example   <run the new version, name is now the same as old version>
            'Del DeleteOLD.bat   <delete this batch file>
            
    Print #1, "@Echo off" & vbCrLf & _
              ":S" & vbCrLf & _
              "Del " & Chr(34) & App.EXEName & ".exe" & Chr(34) & vbCrLf & _
              "If Exist " & Chr(34) & App.EXEName & ".exe" & Chr(34) & " Goto S" & vbCrLf & _
              ":D" & vbCrLf & _
              "ren " & Chr(34) & "Update ExampleNEW.exe" & Chr(34) & " " & Chr(34) & "Update Example.exe" & Chr(34) & vbCrLf & _
              "If Exist " & Chr(34) & "Update ExampleNEW.exe" & Chr(34) & " Goto D" & vbCrLf & _
              Chr(34) & "Update Example.exe" & Chr(34) & vbCrLf & "Del DeleteOLD.bat"
    Close #1
    
         'run the batch file, make it run hidden
        Shell "DeleteOLD.bat", vbHide
            End
End Sub

