VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.UserControl ctlFTP 
   CanGetFocus     =   0   'False
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   900
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1320
   ScaleWidth      =   900
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "FTP"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "ctlFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-------------------------------------------------------------------------------------------'
'   This code is developed by Ronald Kas (r.kas@kaycys.com)                                 '
'   from Kaycys (http://www.kaycys.com).                                                    '
'                                                                                           '
'   You may use this for all purposes except from making profit with it.                    '
'   Check our site regulary for updates.                                                    '
'-------------------------------------------------------------------------------------------'


Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long

Enum EnumBusyWith
    bwNothing = 0
    bwListing = 1
    bwDownloading = 2
    bwUploading = 3
    bwGetCurDir = 4
    bwGetResult = 5
End Enum



Public FtpUrl As String
Public Username As String
Public PassWord As String
Public TimeOut As Long
Public ProxyServer As String
Public ProxyPort As Long
Public InitDir As String
Public Files As New Collection
Public LogFile As String
Dim LoggingState As Boolean
Dim LastAction As EnumBusyWith
Dim CurDir As String
Dim sResponseString As String
Dim lResponseCode As Long

Event Error(ErrorCode As Long)

Public Sub Connect()
On Error Resume Next
     Inet1.Cancel
    
On Error GoTo LocalError
        
    If ProxyServer = "" Then
        Inet1.AccessType = icUseDefault
    Else
        Inet1.AccessType = icNamedProxy
        Inet1.Proxy = ProxyServer & ":" & Trim(CStr(ProxyPort))
    End If
    
    Inet1.URL = FtpUrl
    Inet1.Username = Username
    Inet1.PassWord = PassWord
    If TimeOut = 0 Then
        Inet1.RequestTimeout = 40
    Else
        Inet1.RequestTimeout = TimeOut
    End If
    If LoggingState = True Then
        Print #1, Now & "  Connecting with this information:"
        Print #1, "         URL: " & FtpUrl
        Print #1, "         Username: " & Username
        Print #1, "         Password: " & PassWord
        Print #1, "         Proxy: " & ProxyServer & ":" & ProxyPort
    End If
    
    
    
    Inet1.Execute , "DIR"
    Do While Inet1.StillExecuting
         DoEvents
    Loop
    
    
    Exit Sub
LocalError:
    If LoggingState = True Then
        Print #1, Now & "  Error: " & CStr(Err) & "-" & Error(Err)
    End If
    Inet1.Cancel
    RaiseEvent Error(Err)
End Sub


Public Sub GetFiles(strDir As String)
On Error GoTo LocalError
    Dim i As Integer
    For i = 1 To Files.Count
        Files.Remove (1)
    Next i
    
    LastAction = bwListing
    If LoggingState = True Then
        Print #1, Now & "  Listing files of: " & strDir
    End If
    
    Inet1.Execute , "DIR" & " " & strDir
    Do While Inet1.StillExecuting
         DoEvents
    Loop
    If LoggingState = True Then
        Print #1, Now & "  Response on listing files: " & Inet1.ResponseInfo
    End If
        
    Exit Sub
LocalError:
    RaiseEvent Error(Err)
    If LoggingState = True Then
        Print #1, Now & "  Error: " & CStr(Err) & "-" & Error(Err)
    End If

End Sub


Private Sub Inet1_StateChanged(ByVal State As Integer)
On Error GoTo LocalError
    Dim DirData As String
    Dim tmpData As String
    Dim bDone As Boolean
    Dim strEntry As String
    Dim i As Integer
    Dim k As Integer
    sResponseString = ""
    lResponseCode = 0
    If State = icResponseCompleted Then
        Debug.Print "Response complete.."
        If LastAction = bwListing Then
            Do Until bDone = True
                tmpData = Inet1.GetChunk(1024, icString)
                DoEvents
                If Len(tmpData) = 0 Then
                    bDone = True
                Else
                    DirData = DirData & tmpData
                End If
            Loop
             
             
            For i = 1 To Len(DirData) - 1
                 k = InStr(i, DirData, vbCrLf)
                 strEntry = Mid(DirData, i, k - i)
                 If Right(strEntry, 1) = "/" Then
                      strEntry = Left(strEntry, Len(strEntry) - 1) & "/"
                 End If
                 If Trim(strEntry) <> "" Then
                      Files.Add strEntry
                 End If
                 i = k + 1
                 DoEvents
            Next i
            If LoggingState = True Then
                Print #1, Now & "  Listing complete."
            End If
        
            LastAction = bwNothing
        ElseIf LastAction = bwGetCurDir Then
            CurDir = Inet1.GetChunk(10000, icString)
            LastAction = bwNothing
        ElseIf LastAction = bwGetResult Then
            sResponseString = Inet1.GetChunk(10000, icString)
            Debug.Print Trim(sResponseString)
            LastAction = bwNothing
        End If
    ElseIf State = icError Then
        Debug.Print "Error..."
        sResponseString = Inet1.ResponseInfo
        lResponseCode = CLng(Inet1.ResponseCode)
        If LoggingState = True Then
            Print #1, Now & "  Error: " & CStr(Inet1.ResponseCode) & "-" & Inet1.ResponseInfo
        End If
        
    ElseIf State = icResolvingHost Then
        Debug.Print "Resolving host.."
        If LoggingState = True Then
            Print #1, Now & "  Resolving host.."
        End If
    ElseIf State = icHostResolved Then
        Debug.Print "Host resolved..."
        If LoggingState = True Then
            Print #1, Now & "  Host resolved"
        End If
    ElseIf State = icConnecting Then
        Debug.Print "Connecting..."
        If LoggingState = True Then
            Print #1, Now & "  Connectiong..."
        End If
    ElseIf State = icConnected Then
        Debug.Print "Connected..."
        If LoggingState = True Then
            Print #1, Now & "  Connected"
        End If
    ElseIf State = icRequesting Then
        Debug.Print "Requesting..."
        If LoggingState = True Then
            Print #1, Now & "  Sending request..."
        End If
    ElseIf State = icRequestSent Then
        Debug.Print "Request send..."
        If LoggingState = True Then
            Print #1, Now & "  Request send"
        End If
    ElseIf State = icReceivingResponse Then
        Debug.Print "Receiving response..."
        If LoggingState = True Then
            Print #1, Now & "  Receiving data..."
        End If
    
    ElseIf State = icResponseReceived Then
        Debug.Print "response received"
        If LoggingState = True Then
            Print #1, Now & "  Response received"
        End If
    ElseIf State = icDisconnecting Then
        If LoggingState = True Then
            Print #1, Now & "  Disconnecting.."
        End If
    ElseIf State = icDisconnected Then
        If LoggingState = True Then
            Print #1, Now & "  Disconnected"
        End If
    End If
    
    Exit Sub
LocalError:
    RaiseEvent Error(Err)
    If LoggingState = True Then
        Print #1, Now & "  Error: " & CStr(Err) & "-" & Error(Err)
    End If
    
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = Label1.Width
    UserControl.Height = Label1.Height
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    Inet1.Execute , "CLOSE"
    Inet1.Cancel
    Call StopLogging
End Sub

Public Sub CloseSession()
    On Error Resume Next
    Inet1.Execute , "CLOSE"
    Inet1.Cancel
End Sub

Public Sub DownLoadFile(SourceFile As String, DestFile As String)
On Error GoTo LocalError

    LastAction = bwDownloading
    If LoggingState = True Then
        Print #1, Now & "  Downloading file: " & SourceFile & " to " & DestFile
    End If
    
    Inet1.Execute , "GET " & SourceFile & " " & DestFile
    Do While Inet1.StillExecuting = True
        DoEvents
    Loop
    If LoggingState = True Then
        Print #1, Now & "  Response on download: " & Inet1.ResponseInfo
    End If
    
    Exit Sub
LocalError:
    RaiseEvent Error(Err)
    If LoggingState = True Then
        Print #1, Now & "  Error: " & CStr(Err) & "-" & Error(Err)
    End If
End Sub

Public Sub UploadFile(LocalFile As String, RemoteFile As String)
On Error GoTo LocalError
    LastAction = bwUploading
    If LoggingState = True Then
        Print #1, Now & "  Uploading file from: " & LocalFile & " to " & RemoteFile
    End If
    
    Inet1.Execute , "PUT " & LocalFile & " " & RemoteFile
    Do While Inet1.StillExecuting = True
        DoEvents
    Loop
    If LoggingState = True Then
        Print #1, Now & "  Response on upload: " & Inet1.ResponseInfo
    End If
    Exit Sub
LocalError:
    RaiseEvent Error(Err)
    If LoggingState = True Then
        Print #1, Now & "  Error: " & CStr(Err) & "-" & Error(Err)
    End If
End Sub

Public Sub MkDir(DirSpec As String)
On Error GoTo LocalError
    If LoggingState = True Then
        Print #1, Now & "  Creating directory: " & DirSpec
    End If
    Inet1.Execute , "MKDIR " & DirSpec
    Do While Inet1.StillExecuting = True
        DoEvents
    Loop
    If LoggingState = True Then
        Print #1, Now & "  Response on creating dir: " & Inet1.ResponseInfo
    End If
    
    Exit Sub
LocalError:
    RaiseEvent Error(Err)
    If LoggingState = True Then
        Print #1, Now & "  Error: " & CStr(Err) & "-" & Error(Err)
    End If
End Sub

Public Sub ChDir(DirSpec As String)
On Error GoTo LocalError
    LastAction = bwGetResult
    If LoggingState = True Then
        Print #1, Now & "  Change directory to: " & DirSpec
    End If
    Inet1.Execute , "CD " & DirSpec
    Do While Inet1.StillExecuting = True
        DoEvents
    Loop
    If LoggingState = True Then
        Print #1, Now & "  Response on changedir: " & Inet1.ResponseInfo
    End If
    
    Exit Sub
LocalError:
    RaiseEvent Error(Err)
    If LoggingState = True Then
        Print #1, Now & "  Error: " & CStr(Err) & "-" & Error(Err)
    End If

End Sub

Public Sub DeleteFile(FileSpec As String)
On Error GoTo LocalError
    If LoggingState = True Then
        Print #1, Now & "  Deleting file: " & FileSpec
    End If
    Inet1.Execute , "DELETE " & FileSpec
    Do While Inet1.StillExecuting = True
        DoEvents
    Loop
    If LoggingState = True Then
        Print #1, Now & "  Respsone on delete file: " & Inet1.ResponseInfo
    End If
    
    Exit Sub
LocalError:
    RaiseEvent Error(Err)
    If LoggingState = True Then
        Print #1, Now & "  Error: " & CStr(Err) & "-" & Error(Err)
    End If
End Sub

Public Sub DeletDir(FileSpec As String)
On Error GoTo LocalError
    If LoggingState = True Then
        Print #1, Now & "  Delete directory: " & FileSpec
    End If
    Inet1.Execute , "RMDIR " & FileSpec
    Do While Inet1.StillExecuting = True
        DoEvents
    Loop
    If LoggingState = True Then
        Print #1, Now & "  Response on delete directory: " & Inet1.ResponseInfo
    End If
    
    Exit Sub
LocalError:
    RaiseEvent Error(Err)
    If LoggingState = True Then
        Print #1, Now & "  Error: " & CStr(Err) & "-" & Error(Err)
    End If

End Sub

Public Function GetCurDir() As String
On Error GoTo LocalError
    LastAction = bwGetCurDir
    
    If LoggingState = True Then
        Print #1, Now & "  Getting current remote dir.."
    End If
    Inet1.Execute , "PWD"
    Do While Inet1.StillExecuting = True
        DoEvents
    Loop
    If LoggingState = True Then
        Print #1, Now & "  Response on getcurdir: " & Inet1.ResponseInfo
    End If
    
    GetCurDir = Trim(CurDir)
    Exit Function
LocalError:
    RaiseEvent Error(Err)
    If LoggingState = True Then
        Print #1, Now & "  Error: " & CStr(Err) & "-" & Error(Err)
    End If
End Function

Public Sub Rename(SourceFile As String, DestFile As String)
On Error GoTo LocalError
    If LoggingState = True Then
        Print #1, Now & "  Rename file: " & SourceFile & " to " & DestFile
    End If
    Inet1.Execute , "RENAME " & SourceFile & " " & DestFile
    Do While Inet1.StillExecuting = True
        DoEvents
    Loop
    If LoggingState = True Then
        Print #1, Now & "  Response on rename file: " & Inet1.ResponseInfo
    End If
    Exit Sub
LocalError:
    RaiseEvent Error(Err)
    If LoggingState = True Then
        Print #1, Now & "  Error: " & CStr(Err) & "-" & Error(Err)
    End If
End Sub

Private Function IsNetConnected() As Boolean
    IsNetConnected = InternetGetConnectedState(0, 0)
End Function

Public Property Get ResponseInfo() As String
    ResponseInfo = Inet1.ResponseInfo
End Property

Public Property Get ResponseCode() As Long
    ResponseCode = Inet1.ResponseCode
End Property

Public Sub StartLogging()
    On Error GoTo LocalError
    If LoggingState = False Then
        Open LogFile For Output As #1
        Print #1, Now & "  Logging started"
        LoggingState = True
    End If
Exit Sub
LocalError:
    RaiseEvent Error(Err)
    If LoggingState = True Then
        Print #1, Now & "  Error: " & CStr(Err) & "-" & Error(Err)
    End If
End Sub

Public Sub StopLogging()
    If LoggingState = True Then
        Print #1, Now & "  End logging"
        Close #1
        LoggingState = False
    End If
End Sub

Public Sub AddLogEntry(strLogEntry As String)
    If LoggingState = True Then
        Print #1, Now & "  " & strLogEntry
    End If
End Sub

Public Function StillExecuting() As Boolean
    StillExecuting = Inet1.StillExecuting
End Function

Public Sub SetTo(s As String)
    LastAction = bwGetResult
    Inet1.Execute , "TYPE " & UCase(s)
    Do While Inet1.StillExecuting = True
        DoEvents
    Loop
    
End Sub
