VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8100
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   11235
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11235
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog cdFileOpen 
      Left            =   8550
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraMain 
      Height          =   6360
      Left            =   90
      TabIndex        =   15
      Top             =   855
      Width           =   11010
      Begin VB.CommandButton cmdCopy 
         Caption         =   "&Copy to clipboard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   9090
         TabIndex        =   27
         Top             =   5130
         Width           =   1635
      End
      Begin VB.Frame fraHash 
         Height          =   1440
         Index           =   2
         Left            =   8880
         TabIndex        =   22
         Top             =   1935
         Width           =   2010
         Begin VB.CheckBox chkExtraInfo 
            Caption         =   "Return hashed data as lowercase"
            Height          =   435
            Left            =   180
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   900
            Width           =   1695
         End
         Begin VB.PictureBox picDataType 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1110
            Index           =   0
            Left            =   90
            ScaleHeight     =   1110
            ScaleWidth      =   1815
            TabIndex        =   23
            Top             =   180
            Width           =   1815
            Begin VB.OptionButton optDataType 
               Caption         =   "File"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   915
               TabIndex        =   5
               Top             =   375
               Width           =   555
            End
            Begin VB.OptionButton optDataType 
               Caption         =   "String"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   60
               TabIndex        =   4
               Top             =   375
               Value           =   -1  'True
               Width           =   720
            End
            Begin VB.Label lblAlgo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   405
               TabIndex        =   24
               Top             =   90
               Width           =   735
            End
         End
      End
      Begin VB.Frame fraHash 
         Height          =   1770
         Index           =   1
         Left            =   8880
         TabIndex        =   19
         Top             =   120
         Width           =   2010
         Begin VB.ComboBox cboHash 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmMain.frx":030A
            Left            =   90
            List            =   "frmMain.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   450
            Width           =   1785
         End
         Begin VB.ComboBox cboRounds 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmMain.frx":030E
            Left            =   90
            List            =   "frmMain.frx":0310
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1215
            Width           =   1785
         End
         Begin VB.Label lblAlgo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hash Algorithm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   21
            Top             =   180
            Width           =   1530
         End
         Begin VB.Label lblAlgo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Number of rounds"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   60
            TabIndex        =   20
            Top             =   945
            Width           =   1530
         End
      End
      Begin VB.Frame fraHash 
         Height          =   6150
         Index           =   0
         Left            =   105
         TabIndex        =   16
         Top             =   105
         Width           =   8655
         Begin VB.PictureBox picProgressBar 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   90
            ScaleHeight     =   240
            ScaleWidth      =   8370
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   5715
            Width           =   8430
         End
         Begin RichTextLib.RichTextBox txtOutput 
            Height          =   855
            Left            =   90
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   4695
            Width           =   8430
            _ExtentX        =   14870
            _ExtentY        =   1508
            _Version        =   393217
            BackColor       =   14737632
            ScrollBars      =   2
            TextRTF         =   $"frmMain.frx":0312
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtInputData 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   0
            Text            =   "frmMain.frx":0395
            Top             =   375
            Width           =   7860
         End
         Begin VB.CommandButton cmdBrowse 
            Height          =   375
            Left            =   8040
            Picture         =   "frmMain.frx":03A4
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   360
            Width           =   465
         End
         Begin VB.Label lblHash 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   18
            Top             =   4425
            Width           =   5820
         End
         Begin VB.Label lblHash 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   17
            Top             =   135
            Width           =   5820
         End
      End
   End
   Begin VB.CommandButton cmdChoice 
      Height          =   640
      Index           =   2
      Left            =   9705
      Picture         =   "frmMain.frx":04A6
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Display credits"
      Top             =   7335
      Width           =   640
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   2
      Left            =   4515
      Picture         =   "frmMain.frx":07B0
      ScaleHeight     =   435
      ScaleWidth      =   2205
      TabIndex        =   13
      Top             =   15
      Width           =   2205
   End
   Begin VB.CommandButton cmdChoice 
      Height          =   640
      Index           =   1
      Left            =   9000
      Picture         =   "frmMain.frx":0C6B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7335
      Width           =   640
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   10425
      Picture         =   "frmMain.frx":10AD
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   180
      Width           =   480
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   300
      Picture         =   "frmMain.frx":13B7
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   180
      Width           =   480
   End
   Begin VB.CommandButton cmdChoice 
      Height          =   640
      Index           =   3
      Left            =   10425
      Picture         =   "frmMain.frx":16C1
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Terminate this application"
      Top             =   7335
      Width           =   640
   End
   Begin VB.CommandButton cmdChoice 
      Height          =   640
      Index           =   0
      Left            =   9000
      Picture         =   "frmMain.frx":19CB
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7335
      Width           =   640
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDisclaimer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   135
      TabIndex        =   14
      Top             =   7515
      Width           =   4035
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5062
      TabIndex        =   10
      Top             =   540
      Width           =   1245
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ***************************************************************************
' Module Variables
' ***************************************************************************
  Private mlngRounds        As Long
  Private mlngDisplay       As Long
  Private mlngHashAlgo      As Long
  Private mstrFolder        As String
  Private mstrFilename      As String
  Private mblnStringData    As Boolean
  Private mblnHashLowercase As Boolean
  Private mobjKeyEdit       As cKeyEdit
  
  ' 29-Jan-2010 Add events to track hash progress
  Private WithEvents mobjHash As kiHash.cHash
Attribute mobjHash.VB_VarHelpID = -1

Private Sub cboHash_Click()

    Dim lngIdx As Long
    
    txtOutput.Text = vbNullString
    cmdCopy.Enabled = False
    mlngHashAlgo = cboHash.ListIndex
    
    ' Multiple hash rounds only available
    ' for testing output values
    Select Case cboHash.ListIndex
           
           ' MD5, RipeMD family, SHA family, Whirlpool family
           Case 0 To 13, 21 To 24
                With cboRounds
                    .Clear
                    For lngIdx = 1 To 10
                        .AddItem CStr(lngIdx)
                    Next lngIdx
                    .ListIndex = 0  ' Default rounds = 1
                End With
                    
           Case Else   ' Tiger family
                With cboRounds
                    .Clear
                    For lngIdx = 3 To 15
                        .AddItem CStr(lngIdx)
                    Next lngIdx
                    .ListIndex = 0  ' Default rounds = 3
                End With
    End Select
    
    mlngRounds = CLng(Trim$(Left$(cboRounds.Text, 2)))
    
End Sub

Private Sub cboRounds_Click()

    txtOutput.Text = vbNullString
    cmdCopy.Enabled = False
    mlngRounds = CLng(Trim$(Left$(cboRounds.Text, 2)))
    
End Sub

Private Sub chkExtraInfo_Click()

    ' Hash processing
    ' Checked   - Return hashed data in lowercase format
    ' Unchecked - Return hashed data in uppercase format
    
    mblnHashLowercase = CBool(chkExtraInfo.Value)

End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtOutput.Text
End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub cmdBrowse_Click()
    
    Dim strFilters  As String

    mstrFilename = vbNullString
    mstrFolder = vbNullString
    txtOutput.Text = vbNullString
    txtInputData.Text = vbNullString
    cmdCopy.Enabled = False
    
    strFilters = "All Files (*.*)|*.*"

    On Error GoTo Cancel_Selected

    ' Get the file location. Display the File Open dialog box
    With cdFileOpen
         .CancelError = True  ' Set CancelError is True
         .DialogTitle = "Select file to process"
         .DefaultExt = "*.*"
         .Filter = strFilters
         .Flags = cdlOFNLongNames Or cdlOFNExplorer
         .FilterIndex = 1  ' Specify default filter
         .FileName = vbNullString
         .ShowOpen         ' Display the Open dialog box
    End With

    ' Save the name of the item selected
    mstrFilename = TrimStr(cdFileOpen.FileName)

    ' separate path from filename
    If Len(mstrFilename) > 0 Then
        txtInputData.Text = ShrinkToFit(mstrFilename, 70)  ' Original file name
    End If

CleanUp:
    Exit Sub

Cancel_Selected:
    On Error GoTo 0
    GoTo CleanUp

End Sub

Private Sub cmdChoice_Click(Index As Integer)

    Select Case Index
    
           Case 0  ' OK button
                Screen.MousePointer = vbHourglass
                gblnStopProcessing = False
                ResetProgressBar
                cmdChoice_GotFocus 1
                
                DoEvents
                LockDownCtrls
                Hash_Processing
                
                DoEvents
                SaveLastPath
                UnLockCtrls
                ResetProgressBar
                cmdChoice_GotFocus 0
                Screen.MousePointer = vbDefault

           Case 1  ' Stop button
                DoEvents
                Screen.MousePointer = vbDefault
                gblnStopProcessing = True
                mobjHash.StopProcessing = True
                DoEvents
                
                SaveLastPath
                ResetProgressBar
                UnLockCtrls
                cmdChoice_GotFocus 0
                
           Case 2  ' Show About form
                frmMain.Hide
                frmAbout.DisplayAbout

           Case Else  ' EXIT button
                Screen.MousePointer = vbDefault
                gblnStopProcessing = True
                ResetProgressBar
                
                DoEvents
                SaveLastPath
                TerminateProgram
    End Select

CleanUp:
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdChoice_GotFocus(Index As Integer)

    Select Case Index
           Case 0
                cmdChoice(0).Enabled = True
                cmdChoice(0).Visible = True
                cmdChoice(1).Visible = False
                cmdChoice(1).Enabled = False
           Case 1
                cmdChoice(0).Visible = False
                cmdChoice(0).Enabled = False
                cmdChoice(1).Enabled = True
                cmdChoice(1).Visible = True
    End Select

    Refresh

End Sub

Private Sub Form_Load()

    Set mobjKeyEdit = New cKeyEdit
    Set mobjHash = New cHash
    
    gblnStopProcessing = False
    GetLastPath
    LoadComboBox
    mlngDisplay = 0  ' Display in hex format
    
    With frmMain
        .Caption = gstrVersion
        .lblHash(0).Caption = "Data to be hashed"
        .lblHash(1).Caption = "Hashed results"
        .lblDisclaimer.Caption = "This is a freeware product." & vbNewLine & _
                                 "No warranties or guarantees implied or intended."
        optDataType_Click 0
        ResetProgressBar
        UnLockCtrls
        
        .txtOutput.BackColor = &HE0E0E0   ' Light gray
        .txtOutput.Text = vbNullString
        .cmdCopy.Enabled = False
     
        ' Center the form on the screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show vbModeless   ' reduce flicker
        .Refresh
    End With

    cmdChoice_GotFocus 0
    mobjKeyEdit.CenterCaption frmMain
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set mobjHash = Nothing
    Set mobjKeyEdit = Nothing
    
    Screen.MousePointer = vbDefault

    If UnloadMode = 0 Then
        TerminateProgram
    End If
    
End Sub

' 29-Jan-2010 Add events to track hash progress
Private Sub mobjHash_HashProgress(ByVal lngProgress As Long)
    
    ProgressBar picProgressBar, lngProgress
    DoEvents
    
End Sub

Private Sub optDataType_Click(Index As Integer)

    Select Case Index

           Case 0   ' string data
                mblnStringData = True
    
                With frmMain
                    .optDataType(0).Value = True
                    .optDataType(1).Value = False
                    .cmdBrowse.Enabled = False
                    .cmdBrowse.Visible = False
                    .txtOutput.Text = vbNullString
                    
                    With .txtInputData
                        .Locked = False
                        .Height = 4000
                        .Top = 375
                        .Width = 8430
                        .Text = vbNullString
                    End With
                End With
                                
           Case 1   ' file data
                mblnStringData = False
    
                With frmMain
                    .optDataType(0).Value = False
                    .optDataType(1).Value = True
                    .cmdBrowse.Enabled = True
                    .cmdBrowse.Visible = True
                    .txtOutput.Text = vbNullString
                    
                    With .txtInputData
                        .Locked = False
                        .Height = 330
                        .Top = 375
                        .Width = 7860
                        .Text = ShrinkToFit(mstrFilename, 70)
                        .Locked = True
                    End With
                End With
    End Select

End Sub

' ***************************************************************************
' Data functions
' ***************************************************************************
Private Sub LoadComboBox()

    Dim lngIdx As Long
    
    With frmMain
        ' Hash algorithms
        With .cboHash
            .Clear
            .AddItem "MD4"              ' 0
            .AddItem "MD5"              ' 1
            .AddItem "SHA-1"            ' 2
            .AddItem "SHA-224"          ' 3
            .AddItem "SHA-256"          ' 4
            .AddItem "SHA-384"          ' 5
            .AddItem "SHA-512"          ' 6
            .AddItem "SHA-512/224"      ' 7
            .AddItem "SHA-512/256"      ' 8
            .AddItem "SHA-512/320"      ' 9
            .AddItem "RipeMD-128"       ' 10
            .AddItem "RipeMD-160"       ' 11
            .AddItem "RipeMD-256"       ' 12
            .AddItem "RipeMD-320"       ' 13
            .AddItem "Tiger-128"        ' 14
            .AddItem "Tiger-160"        ' 15
            .AddItem "Tiger-192"        ' 16
            .AddItem "Tiger-224"        ' 17
            .AddItem "Tiger-256"        ' 18
            .AddItem "Tiger-384"        ' 19
            .AddItem "Tiger-512"        ' 20
            .AddItem "Whirlpool-224"    ' 21
            .AddItem "Whirlpool-256"    ' 22
            .AddItem "Whirlpool-384"    ' 23
            .AddItem "Whirlpool-512"    ' 24
            .ListIndex = 4
        End With
            
        With .cboRounds
            .Clear
            For lngIdx = 1 To 10
                .AddItem CStr(lngIdx)
            Next lngIdx
            .ListIndex = 0
        End With
    End With
    
End Sub

Private Sub Hash_Processing()
                    
    Dim strOutput  As String
    Dim abytData() As Byte
    Dim abytHash() As Byte
    
    Screen.MousePointer = vbHourglass
    
    Erase abytData()    ' Always start with empty arrays
    Erase abytHash()
    
    strOutput = vbNullString
    txtOutput.Text = vbNullString
    cmdCopy.Enabled = False
    
    If mblnStringData Then
        ' Test for string data to process
        If Len(Trim$(txtInputData.Text)) = 0 Then
            InfoMsg "Need some data to process"
            txtInputData.SetFocus
            GoTo Hash_Processing_CleanUp
        End If
    Else
        ' Test for file name to process
        If Len(Trim$(mstrFilename)) = 0 Then
            InfoMsg "Path\File name missing"
            txtInputData.SetFocus
            GoTo Hash_Processing_CleanUp
        End If
    End If
    
    With mobjHash
        .StopProcessing = False                ' Reset stop flag
        .HashMethod = mlngHashAlgo             ' Hash algorithm selected
        .HashRounds = mlngRounds               ' Number of passes
        .ReturnLowercase = mblnHashLowercase   ' TRUE = Return as lowercase
                                               ' FALSE = Return as uppercase
        ' Hash string data
        If optDataType(0).Value Then
            
            abytData() = StringToByteArray(txtInputData.Text)   ' Convert to byte array
            abytHash() = .HashString(abytData())                ' Hash string data
            gblnStopProcessing = .StopProcessing                ' See if processing aborted
            strOutput = ByteArrayToString(abytHash())           ' Convert byte array to string
            
        Else
            ' Hash a file
            If IsPathValid(mstrFilename) Then
                abytData() = StringToByteArray(mstrFilename)    ' Convert to byte array
                abytHash() = .HashFile(abytData())              ' Hash file
                gblnStopProcessing = .StopProcessing            ' See if processing aborted
                strOutput = ByteArrayToString(abytHash())       ' Convert byte array to string
            Else
                InfoMsg "Cannot locate Path\File." & vbNewLine & mstrFilename
                txtInputData.SetFocus
                GoTo Hash_Processing_CleanUp
            End If
        End If
    
    End With
    
    DoEvents
    If gblnStopProcessing Then
        GoTo Hash_Processing_CleanUp
    End If
    
    txtOutput.Text = TrimStr(strOutput)
    cmdCopy.Enabled = True
    
Hash_Processing_CleanUp:
    Erase abytData()    ' Always empty arrays when not needed
    Erase abytHash()
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub ResetProgressBar()

    ' Resets progressbar to zero
    ' with all white background
    ProgressBar picProgressBar, 0, vbWhite
    
End Sub

' ***************************************************************************
' Routine:       ProgessBar
'
' Description:   Fill a picturebox as if it were a horizontal progress bar.
'
' Parameters:    objProgBar - name of picture box control
'                lngPercent - Current percentage value
'                lngForeColor - Optional-The progression color. Default = Black.
'                           can use standard VB colors or long Integer
'                           values representing a color.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2001  Randy Birch  http://vbnet.mvps.org/index.html
'              Routine created
' 14-FEB-2005  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 01-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for the speed enhancement.
'              This way the progress bar is only initialized once.
' 05-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' ***************************************************************************
Private Sub ProgressBar(ByRef objProgBar As PictureBox, _
                        ByVal lngPercent As Long, _
               Optional ByVal lngForeColor As Long = vbBlue)

    Dim strPercent As String
    
    Const MAX_PERCENT As Long = 100
    
    ' Called by ResetProgressBar() routine
    ' to reinitialize progress bar properties.
    ' If forecolor is white then progressbar
    ' is being reset to a starting position.
    If lngForeColor = vbWhite Then
        
        With objProgBar
            .AutoRedraw = True      ' Required to prevent flicker
            .BackColor = &HFFFFFF   ' White
            .DrawMode = 10          ' Not Xor Pen
            .FillStyle = 0          ' Solid fill
            .FontName = "Arial"     ' Name of font
            .FontSize = 11          ' Font point size
            .FontBold = True        ' Font is bold.  Easier to see.
            Exit Sub                ' Exit this routine
        End With
    
    End If
        
    ' If no progress then leave
    If lngPercent < 1 Then
        Exit Sub
    End If
    
    ' Verify flood display has not exceeded 100%
    If lngPercent <= MAX_PERCENT Then

        With objProgBar
        
            ' Error trap in case code attempts to set
            ' scalewidth greater than the max allowable
            If lngPercent > .ScaleWidth Then
                lngPercent = .ScaleWidth
            End If
               
            .Cls                        ' Empty picture box
            .ForeColor = lngForeColor   ' Reset forecolor
         
            ' set picture box ScaleWidth equal to maximum percentage
            .ScaleWidth = MAX_PERCENT
            
            ' format percent into a displayable value (ex: 25%)
            strPercent = Format$(CLng((lngPercent / .ScaleWidth) * 100)) & "%"
            
            ' Calculate X and Y coordinates within
            ' picture box and and center data
            .CurrentX = (.ScaleWidth - .TextWidth(strPercent)) \ 2
            .CurrentY = (.ScaleHeight - .TextHeight(strPercent)) \ 2
                
            objProgBar.Print strPercent   ' print percentage string in picture box
            
            ' Print flood bar up to new percent position in picture box
            objProgBar.Line (0, 0)-(lngPercent, .ScaleHeight), .ForeColor, BF
        
        End With
                
        DoEvents   ' allow flood to complete drawing
    
    End If

End Sub

Private Sub GetLastPath()

    mstrFilename = GetSetting("kiHash", "Settings", "Filename", App.Path & "\TestFile.txt")
    mstrFolder = GetSetting("kiHash", "Settings", "LastPath", App.Path & "\")
    
End Sub

Private Sub SaveLastPath()

    SaveSetting "kiHash", "Settings", "Filename", mstrFilename
    SaveSetting "kiHash", "Settings", "LastPath", mstrFolder

End Sub

Private Sub txtInputData_GotFocus()
    ' Highlight contents in text box
    mobjKeyEdit.TextBoxFocus txtInputData
End Sub

Private Sub txtInputData_KeyDown(KeyCode As Integer, Shift As Integer)
    ' key control (Ex:   Ctrl+C, etc.)
    mobjKeyEdit.TextBoxKeyDown txtInputData, KeyCode, Shift
End Sub

Private Sub txtInputData_KeyPress(KeyAscii As Integer)
        
    ' edit data input
    Select Case KeyAscii
           Case 9
                ' Tab key
                KeyAscii = 0
                SendKeys "{TAB}"
                
           Case 8, 13, 32 To 126
                ' Backspace, ENTER key and
                ' other valid data keys
                
           Case Else  ' Everything else (invalid)
                KeyAscii = 0
    End Select

End Sub

Private Sub LockDownCtrls()

    With frmMain
        .cmdCopy.Enabled = False
        .cmdChoice(2).Enabled = False
        .cmdChoice(3).Enabled = False
        .fraHash(1).Enabled = False
        .fraHash(2).Enabled = False
    End With

End Sub
                
Private Sub UnLockCtrls()

    With frmMain
        .cmdCopy.Enabled = True
        .cmdChoice(2).Enabled = True
        .cmdChoice(3).Enabled = True
        .fraHash(1).Enabled = True
        .fraHash(2).Enabled = True
    End With

End Sub
                
                
