Attribute VB_Name = "basFloppyCopy"
Option Explicit

Public Function CleanDir(ByVal dirPath As String) As Boolean
Dim FileName As String   ' Walking filename variable.
Dim DirName As String    ' SubDirectory Name.
Dim dirNames() As String ' Buffer for directory name entries.
Dim nDir As Integer      ' Number of directories in this path.
Dim i As Integer         ' For-loop counter.

    On Error GoTo sysFileERR
    If Right(dirPath, 1) <> "\" Then dirPath = dirPath & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    DirName = Dir(dirPath, vbDirectory Or vbHidden)  ' Even if hidden.
    Do While Len(DirName) > 0
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
        ' Check for directory with bitwise comparison.
            If GetAttr(dirPath & DirName) And vbDirectory Then
                dirNames(nDir) = DirName
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
                Kill (dirPath & DirName)
'                MkDir destPath & DirName
            End If

sysFileERRCont:
        End If
        DirName = Dir()  ' Get next subdirectory.
    Loop

    ' Search through this directory and sum file sizes.
    FileName = Dir(dirPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
    While Len(FileName) <> 0
        Call Kill(dirPath & FileName)
        FileName = Dir()  ' Get next file.
    Wend
    
'    ' If there are sub-directories..
'    If nDir > 0 Then
'        ' Recursively walk into them
'        For i = 0 To nDir - 1
'            Call CopyFilesRecursive(sourcePath & dirNames(i) & "\", destPath & dirNames(i) & "\", progressBarObj)
'        Next i
'    End If
    CleanDir = True
    Exit Function
    
AbortFunction:
    CleanDir = False
    Exit Function
sysFileERR:
    If Right(DirName, 4) = ".sys" Then
        Resume sysFileERRCont ' Known issue with pagefile.sys
    Else
        Resume AbortFunction
    End If
End Function
