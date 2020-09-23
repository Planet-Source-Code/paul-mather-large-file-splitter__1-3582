VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Splitter"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCopyToFloppy 
      Caption         =   "Copy to Floppies when finished Splitting"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1800
      Width           =   3375
   End
   Begin FileSplitter.ctlProgress ctlProgress1 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2880
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      Appearance      =   1
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      FillColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkOpenExplorer 
      Caption         =   "Open Explorer Window to Split Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CheckBox chkDeleteParts 
      Caption         =   "Delete Split Parts when finished Joining"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox txtSplitCount 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2760
      TabIndex        =   16
      Top             =   2400
      Width           =   855
   End
   Begin VB.OptionButton optSplit 
      Caption         =   "File Count"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   1095
   End
   Begin VB.OptionButton optSplit 
      Caption         =   "File Size"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Frame fraEst 
      Caption         =   "Estimates"
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   3495
      Begin VB.Label lblEstFileSize 
         AutoSize        =   -1  'True
         Caption         =   "File Size:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   630
      End
      Begin VB.Label lblEstFileCount 
         AutoSize        =   -1  'True
         Caption         =   "File Count:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.ComboBox cmbSplitSize 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Text            =   "cmbSplitSize"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame fraStats 
      Caption         =   "Statistics"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   3495
      Begin VB.Label lblFileCount 
         AutoSize        =   -1  'True
         Caption         =   "File Count:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lblElapsed 
         AutoSize        =   -1  'True
         Caption         =   "Elapsed Time:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label lblFinish 
         AutoSize        =   -1  'True
         Caption         =   "Finish Time:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "Start Time:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "Split File"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblSplitCount 
      AutoSize        =   -1  'True
      Caption         =   "Split Files"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2760
      TabIndex        =   15
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label lblSplitSize 
      AutoSize        =   -1  'True
      Caption         =   "Split Size (KB)"
      Height          =   195
      Left            =   1320
      TabIndex        =   10
      Top             =   2160
      Width           =   990
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      Caption         =   "Filename:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function SplitFiles(ByVal inputFilename As String, newFileSizeBytes As Long) As Boolean
Dim fReadHandle As Long
Dim fWriteHandle As Long
Dim fSuccess As Long
Dim lBytesWritten As Long
Dim lBytesRead As Long
Dim ReadBuffer() As Byte
Dim TotalCount As Long
Dim StartTime As Date
Dim FinishTime As Date
Dim StartTimeDouble As Double
Dim FinishTimeDouble As Double
Dim Count As Integer
Dim Count2 As Integer
Dim NewFileString As String
Dim ret As Integer
Dim copyPass As Boolean
Dim fullFileString As String
    
    ' Determine Total Number of Output Files for User Interface Only
    If CInt(FileLen(inputFilename) / newFileSizeBytes) < FileLen(inputFilename) / newFileSizeBytes Then
        TotalCount = CInt(FileLen(inputFilename) / newFileSizeBytes) + 1
    Else
        TotalCount = CInt(FileLen(inputFilename) / newFileSizeBytes)
    End If

    If TotalCount > 10 Then
        ret = MsgBox("Are you sure you want to break this file up into " & TotalCount & " pieces?", vbYesNo + vbQuestion, "Are you sure?")
        If ret = vbNo Then
            SplitFiles = False
            Exit Function
        End If
    End If
    
    ' User Interface Stuff
    StartTime = Now
    StartTimeDouble = getTime
    Me.MousePointer = vbHourglass
    cmdJoin.Enabled = False
    cmdSplit.Enabled = False
    txtSplitCount.Enabled = False
    cmbSplitSize.Enabled = False
    optSplit(0).Enabled = False
    optSplit(1).Enabled = False
    lblStart = "Start Time: " & StartTime
    ctlProgress1.Max = FileLen(inputFilename)
    ctlProgress1.CaptionStyle = eCap_CaptionPercent
    ctlProgress1.Caption = "Splitting"
    ctlProgress1.value = 0
    
    Count = 1
    ' Resize Byte Array for Read
    ReDim ReadBuffer(0 To newFileSizeBytes)
            
    ' Open Read File Handle
    If DetermineDirectory(inputFilename) <> "" Then
        ChDir (DetermineDirectory(inputFilename))
    End If
    inputFilename = DetermineFilename(inputFilename)
    fReadHandle = CreateFile(inputFilename, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    
    ' If Successful read, continue
    If fReadHandle <> INVALID_HANDLE_VALUE Then
        ' Read First File Block
        fSuccess = ReadFile(fReadHandle, ReadBuffer(0), UBound(ReadBuffer), lBytesRead, 0)
        
        ' Increment ProgressBar
        ctlProgress1.value = ctlProgress1.value + lBytesRead
        ctlProgress1.Refresh
        
        ' Loop while not EOF
        Do While lBytesRead > 0
            ' Update File Count Statistic on User Interface
            lblFileCount.Caption = "Split Count: " & Count & " of " & TotalCount
            lblFileCount.Refresh
            
            ' Open Write File Handle
            If Dir(inputFilename & "." & Count) <> "" Then
                Kill inputFilename & "." & Count
            End If
            fWriteHandle = CreateFile(inputFilename & "." & Count, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
            ' If Successful Write, Continue
            If fWriteHandle <> INVALID_HANDLE_VALUE Then
                ' Write Data Block to File
                fSuccess = WriteFile(fWriteHandle, ReadBuffer(0), lBytesRead, lBytesWritten, 0)
                If fSuccess <> 0 Then
                    ' Required to Write to File
                    fSuccess = FlushFileBuffers(fWriteHandle)
                    ' Close Write File
                    fSuccess = CloseHandle(fWriteHandle)
                Else
                    ' On Failure Quit
                    lblFileCount.Caption = "Split Count: Write Error"
                    SplitFiles = False
                    Exit Function
                End If
            Else
                ' On Failure Quit
                lblFileCount.Caption = "Split Count: Write Error"
                SplitFiles = False
                Exit Function
            End If
            ' Get the next Read Block
            fSuccess = ReadFile(fReadHandle, ReadBuffer(0), UBound(ReadBuffer), lBytesRead, 0)
                       
            ' Increment ProgressBar
            ctlProgress1.value = ctlProgress1.value + lBytesRead
            ctlProgress1.Refresh
            
            ' Increment Count
            Count = Count + 1
        Loop
        ' Close Read File
        fSuccess = CloseHandle(fReadHandle)
    Else
        ' On Failure Quit
        lblFileCount.Caption = "Split Count: Read Error"
        SplitFiles = False
        Call MsgBox(lblFileCount.Caption, vbOKOnly + vbCritical, "File Error")
        Me.MousePointer = vbDefault
        Call txtFileName_Change
        If optSplit(1).value = True Then
            txtSplitCount.Enabled = True
        Else
            cmbSplitSize.Enabled = True
        End If
        optSplit(0).Enabled = True
        optSplit(1).Enabled = True
        Exit Function
    End If
    
    Open inputFilename & ".bat" For Output As #1
        Print #1, "@ECHO OFF"
        Print #1, "ECHO Joining Files..."
        Print #1, "IF EXIST " & """" & inputFilename & """" & " DEL " & """" & inputFilename & """"
        If chkDeleteParts.value = vbChecked Then
            NewFileString = """" & inputFilename & ".1"""
        Else
            NewFileString = """" & inputFilename & """"
            Print #1, "COPY """ & inputFilename & ".1"" """ & inputFilename & """"
        End If
        If Count - 1 > 1 Then
            For Count2 = 2 To Count - 1
                If NewFileString <> "" Then
                    NewFileString = NewFileString & "+"
                End If
                NewFileString = NewFileString & """" & inputFilename & "." & Count2 & """"
            Next Count2
            Print #1, "COPY /B " & NewFileString
            If chkDeleteParts.value = vbChecked Then
                For Count2 = 2 To Count - 1
                    Print #1, "DEL """ & inputFilename & "." & Count2 & """"
                Next Count2
            End If
        End If
        If chkDeleteParts.value = vbChecked Then
            Print #1, "REN """ & inputFilename & ".1"" """ & inputFilename & """"
            Print #1, "DEL """ & inputFilename & ".bat"""
        End If
        Print #1, "ECHO Done!"
    Close #1

    ' User Interface Stuff
    ctlProgress1.Caption = "Finished Splitting"
    ctlProgress1.CaptionStyle = eCap_CaptionOnly
    ctlProgress1.value = ctlProgress1.Max
    lblFileCount.Caption = "Split Count: Finished (" & Count - 1 & " Files)"
    FinishTime = Now
    FinishTimeDouble = getTime
    lblFinish = "Finish Time: " & FinishTime
    lblElapsed = "Elapsed Time: " & Format(FinishTimeDouble - StartTimeDouble, "#.00") & " seconds"
    Call txtFileName_Change
    Me.MousePointer = vbDefault
    SplitFiles = True
    If optSplit(1).value = True Then
        txtSplitCount.Enabled = True
    Else
        cmbSplitSize.Enabled = True
    End If
    optSplit(0).Enabled = True
    optSplit(1).Enabled = True
    If chkCopyToFloppy.value = vbChecked Then
        For Count = 1 To TotalCount
            copyPass = True
            ret = MsgBox("Please enter Disk #" & Count & " in Drive A:", vbOKCancel + vbInformation, "Coppy to Floppy")
            If ret = vbCancel Then
                frmSplash.Hide
                Exit For
            End If
            DoEvents
            Do While True
                frmSplash.ShowMessage ("Cleaning up disk...")
                If CleanDir("A:\") = False Then
                    frmSplash.Hide
                    ret = MsgBox("Unable to find disk in drive or disk is write-protected" & vbCr & "Please insert another disk and press OK", vbOKCancel + vbCritical, "Disk Missing")
                    If ret = vbCancel Then
                        Exit For
                    End If
                Else
                    On Error GoTo e_Copy_Fail
                    frmSplash.ShowMessage ("Copying " & inputFilename & "." & Count & " to " & "A:\" & inputFilename & "." & Count & "...")
                    Call FileCopy(inputFilename & "." & Count, "a:\" & inputFilename & "." & Count)
                    If copyPass = True Then
                        If Count = TotalCount Then
                            Call FileCopy(inputFilename & ".bat", "A:\" & inputFilename & ".bat")
                        End If
                        Exit Do
                    End If
                End If
            Loop
            frmSplash.Hide
            DoEvents
        Next Count
    End If
    frmSplash.Hide
    If chkOpenExplorer.value = vbChecked Then
        Call Shell("explorer " & CurDir, vbNormalFocus)
    End If
    Exit Function
e_Copy_Fail:
    frmSplash.Hide
    Call MsgBox("Unable to copy " & inputFilename & "." & Count & " to " & " A:\" & inputFilename & "." & Count, vbOKOnly + vbCritical, "File Copy Failed")
    copyPass = False
    Resume Next
End Function
Public Function JoinFiles(ByVal inputFilename As String) As Boolean
Dim fReadHandle As Long
Dim fWriteHandle As Long
Dim fSuccess As Long
Dim lBytesWritten As Long
Dim lBytesRead As Long
Dim ReadBuffer() As Byte
Dim TotalCount As Long
Dim StartTime As Date
Dim FinishTime As Date
Dim StartTimeDouble As Double
Dim FinishTimeDouble As Double
Dim Count As Integer
Dim FileName As String
Dim ret As Integer

    ' Check for existing Output File
    If Dir(inputFilename) <> "" Then
        ret = MsgBox("Output file (" & inputFilename & ") already exists." & vbCr & "Are you sure you want to overwrite it?", vbYesNo + vbQuestion, "Overwrite Warning")
        If ret = vbNo Then
            lblFileCount.Caption = "Join Count: Output File Exists"
            JoinFiles = False
            Exit Function
        Else
            Kill inputFilename
        End If
    End If
        
    ' Determine how many split files are contained in the entire set
    Count = 1
    FileName = Dir(inputFilename & ".1")
    
    If FileName = "" Then
        lblFileCount.Caption = "Join Count: No Input Files"
        JoinFiles = False
        Exit Function
    End If
    
    ctlProgress1.CaptionStyle = eCap_CaptionPercent
    ctlProgress1.Caption = "Joining"
    ctlProgress1.value = 0
    ctlProgress1.Max = FileLen(inputFilename & ".1")
    Do While FileName <> ""
        Count = Count + 1
        FileName = Dir(inputFilename & "." & Count)
        If FileName <> "" Then
            ctlProgress1.Max = ctlProgress1.Max + FileLen(inputFilename & "." & Count)
        End If
    Loop
    TotalCount = Count - 1
    
    ' User Interface Stuff
    StartTime = Now
    StartTimeDouble = getTime
    Me.MousePointer = vbHourglass
    cmdJoin.Enabled = False
    cmdSplit.Enabled = False
    txtSplitCount.Enabled = False
    cmbSplitSize.Enabled = False
    optSplit(0).Enabled = False
    optSplit(1).Enabled = False
    lblStart = "Start Time: " & StartTime

    ' Open Write File Handle
    fWriteHandle = CreateFile(inputFilename, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    
    ' If Successful Write, Continue
    If fWriteHandle <> INVALID_HANDLE_VALUE Then
    
        For Count = 1 To TotalCount
            DoEvents
            ' Open Read File Handle
            ReDim ReadBuffer(0 To FileLen(inputFilename & "." & Count))
            fReadHandle = CreateFile(inputFilename & "." & Count, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
            
            ' If Successful read, continue
            If fReadHandle <> INVALID_HANDLE_VALUE Then
                ' Read First File Block
                fSuccess = ReadFile(fReadHandle, ReadBuffer(0), UBound(ReadBuffer), lBytesRead, 0)
                
                ' Write Data Block to File
                fSuccess = WriteFile(fWriteHandle, ReadBuffer(0), UBound(ReadBuffer), lBytesWritten, 0)
                If fSuccess <> 0 Then
                    ' Required to Write to File
                    fSuccess = FlushFileBuffers(fWriteHandle)
                Else
                    ' On Failure Quit
                    lblFileCount.Caption = "Join Count: Write Error"
                    JoinFiles = False
                    Exit Function
                End If
                                
                fSuccess = CloseHandle(fReadHandle)
            
                 ' Increment ProgressBar
                ctlProgress1.value = ctlProgress1.value + lBytesWritten
                ctlProgress1.Refresh
                
                ' Update File Count Statistic on User Interface
                lblFileCount.Caption = "Join Count: " & Count & " of " & TotalCount
                lblFileCount.Refresh
            Else
                ' On Failure Quit
                lblFileCount.Caption = "Join Count: Read Error"
                JoinFiles = False
                Exit Function
            End If
        
        Next Count
        If chkDeleteParts.value = vbChecked Then
            For Count = 1 To TotalCount
                Kill inputFilename & "." & Count
            Next Count
            If Dir(inputFilename & ".bat") <> "" Then
                Kill inputFilename & ".bat"
            End If
        End If
    Else
        ' On Failure Quit
        lblFileCount.Caption = "Join Count: Write Error"
        JoinFiles = False
        Exit Function
    End If
    
    ' Close Write File
    fSuccess = CloseHandle(fWriteHandle)
            
    ' User Interface Stuff
    ctlProgress1.Caption = "Finished Joining"
    ctlProgress1.CaptionStyle = eCap_CaptionOnly
    ctlProgress1.value = ctlProgress1.Max
    lblFileCount.Caption = "File Count: Finished (" & Count - 1 & " Files)"
    FinishTime = Now
    FinishTimeDouble = getTime
    lblFinish = "Finish Time: " & FinishTime
    lblElapsed = "Elapsed Time: " & Format(FinishTimeDouble - StartTimeDouble, "#.00") & " seconds"
    Call txtFileName_Change
    Me.MousePointer = vbDefault
    If optSplit(1).value = True Then
        txtSplitCount.Enabled = True
    Else
        cmbSplitSize.Enabled = True
    End If
    optSplit(0).Enabled = True
    optSplit(1).Enabled = True
    If chkOpenExplorer.value = vbChecked Then
        Call Shell("explorer " & CurDir, vbNormalFocus)
    End If

End Function


Private Sub cmbSplitSize_Change()
    Call ChangeEstimates
End Sub
Private Sub cmbSplitSize_Click()
    Call cmbSplitSize_Change
End Sub

Private Sub cmdBrowse_Click()
Dim CroppedEnd As String
Dim returnedFiles As SelectedFile
    On Error GoTo e_cmdBrowse
    FileDialog.sFilter = "All Files (*.*)" & Chr$(0) & "*.*"
    FileDialog.flags = &H4 + &H1000
    FileDialog.sInitDir = GetSetting(App.Title, "Settings", "InitDir")
    returnedFiles = ShowOpen(Me.hWnd, True)
    
    CroppedEnd = Mid(returnedFiles.sFiles(1), InStrRev(returnedFiles.sFiles(1), ".") + 1)
    If IsNumeric(CroppedEnd) = True Then
        txtFileName = returnedFiles.sLastDirectory & Mid(returnedFiles.sFiles(1), 1, InStrRev(returnedFiles.sFiles(1), ".") - 1)
    Else
        txtFileName = returnedFiles.sLastDirectory & returnedFiles.sFiles(1)
    End If
    Call SaveSetting(App.Title, "Settings", "InitDir", returnedFiles.sLastDirectory)
    Exit Sub
e_cmdBrowse:
    Exit Sub
End Sub

Private Sub cmdJoin_Click()
    If txtFileName <> "" And Dir(txtFileName & ".1") <> "" Then
        Call JoinFiles(txtFileName)
    Else
        Call txtFileName_Change
        Call MsgBox("File Not Found (" & txtFileName & ")", vbOKOnly + vbCritical, "File Read Error")
    End If
End Sub

Private Sub cmdSplit_Click()
    If txtFileName <> "" And Dir(txtFileName) <> "" Then
        If optSplit(0).value = True Then
            If IsNumeric(cmbSplitSize.Text) = False Then
                cmbSplitSize.SetFocus
                Call MsgBox("Please specify a valid split size!", vbOKOnly + vbCritical, "File Read Error")
            Else
                Call SplitFiles(txtFileName, CLng(cmbSplitSize.Text) * CLng(1024))
            End If
        Else
            If IsNumeric(txtSplitCount.Text) = False Then
                txtSplitCount.SetFocus
                Call MsgBox("Please specify a valid number of files!", vbOKOnly + vbCritical, "File Read Error")
            Else
                Call SplitFiles(txtFileName, FileLen(txtFileName) / CLng(txtSplitCount))
            End If
        End If
    Else
        Call txtFileName_Change
        Call MsgBox("File Not Found (" & txtFileName & ")", vbOKOnly + vbCritical, "File Read Error")
    End If
End Sub

Private Sub Form_Load()
Dim commandString As String

    Me.Caption = "File Splitter v" & App.Major & "." & App.Minor & "." & App.Revision
    
    Call SaveSetting("*\shell\open2", "", "", "Split File", HKEY_CLASSES_ROOT, "")
    Call SaveSetting("*\shell\open2\command", "", "", App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & App.EXEName & ".exe %1", HKEY_CLASSES_ROOT, "")

    On Error GoTo e_next
    commandString = Command
    If commandString <> "" Then
        If Dir(commandString) <> "" Then
            txtFileName = commandString
        End If
    End If

e_next:
    cmbSplitSize.AddItem ("5")
    cmbSplitSize.AddItem ("10")
    cmbSplitSize.AddItem ("50")
    cmbSplitSize.AddItem ("100")
    cmbSplitSize.AddItem ("500")
    cmbSplitSize.AddItem ("1000")
    cmbSplitSize.AddItem ("1200")
    cmbSplitSize.AddItem ("1440")
    cmbSplitSize.AddItem ("2000")
    cmbSplitSize.AddItem ("3000")
    cmbSplitSize.Text = GetSetting(App.Title, "Settings", "SplitSize", "1440")
    txtSplitCount = GetSetting(App.Title, "Settings", "SplitCount", 2)
    optSplit(1).value = GetSetting(App.Title, "Settings", "SplitOption", False)
    chkDeleteParts.value = GetSetting(App.Title, "Settings", "DeleteParts", vbChecked)
    chkOpenExplorer.value = GetSetting(App.Title, "Settings", "OpenExplorer", vbChecked)
    chkCopyToFloppy.value = GetSetting(App.Title, "Settings", "CopyToFloppy", vbUnchecked)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSetting(App.Title, "Settings", "SplitSize", cmbSplitSize.Text)
    Call SaveSetting(App.Title, "Settings", "SplitCount", txtSplitCount)
    Call SaveSetting(App.Title, "Settings", "SplitOption", optSplit(1).value)
    Call SaveSetting(App.Title, "Settings", "DeleteParts", chkDeleteParts.value)
    Call SaveSetting(App.Title, "Settings", "OpenExplorer", chkOpenExplorer.value)
    Call SaveSetting(App.Title, "Settings", "CopyToFloppy", chkCopyToFloppy.value)
    
    Unload frmSplash
End Sub


Private Sub optSplit_Click(Index As Integer)
    If Index = 0 Then
        lblSplitCount.Enabled = False
        txtSplitCount.Enabled = False
        lblSplitSize.Enabled = True
        cmbSplitSize.Enabled = True
    Else
        lblSplitCount.Enabled = True
        txtSplitCount.Enabled = True
        lblSplitSize.Enabled = False
        cmbSplitSize.Enabled = False
    End If
    Call ChangeEstimates
End Sub

Private Sub txtFileName_Change()
    If txtFileName <> "" And Dir(txtFileName) <> "" Then
        cmdSplit.Enabled = True
    Else
        cmdSplit.Enabled = False
    End If
    If txtFileName <> "" And Dir(txtFileName & ".1") <> "" Then
        cmdJoin.Enabled = True
    Else
        cmdJoin.Enabled = False
    End If
    Call ChangeEstimates
End Sub
Private Sub ChangeEstimates()
    lblEstFileCount.Caption = "File Count:"
    lblEstFileSize.Caption = "File Size:"
    If txtFileName = "" Then Exit Sub
    If Dir(txtFileName) = "" Then
        cmdSplit.Enabled = False
        Exit Sub
    End If
    If optSplit(0).value = True Then
        If IsNumeric(cmbSplitSize.Text) = True Then
            cmdSplit.Enabled = True
            If CInt(FileLen(txtFileName) / (CLng(cmbSplitSize.Text) * 1024)) < FileLen(txtFileName) / (CLng(cmbSplitSize.Text) * 1024) Then
                lblEstFileCount.Caption = "File Count: " & CInt(FileLen(txtFileName) / (CLng(cmbSplitSize.Text) * 1024)) + 1
            Else
                lblEstFileCount.Caption = "File Count: " & CInt(FileLen(txtFileName) / (CLng(cmbSplitSize.Text) * 1024))
            End If
            lblEstFileSize.Caption = "File Size: " & cmbSplitSize.Text & " Kilobytes"
        Else
            cmdSplit.Enabled = False
        End If
    Else
        If IsNumeric(txtSplitCount.Text) = True Then
            cmdSplit.Enabled = True
            txtSplitCount = Trim(txtSplitCount)
            If txtSplitCount = "" Then Exit Sub
            If IsNumeric(txtSplitCount) = False Then Exit Sub
            If txtSplitCount = 0 Then Exit Sub
            lblEstFileCount.Caption = "File Count: " & txtSplitCount.Text
            lblEstFileSize.Caption = "File Size: " & Format((FileLen(txtFileName) / CDbl(txtSplitCount)) / 1024, "0.0") & " Kilobytes"
        Else
            cmdSplit.Enabled = False
        End If
    End If
End Sub

Private Sub txtSplitCount_Change()
    Call ChangeEstimates
End Sub
