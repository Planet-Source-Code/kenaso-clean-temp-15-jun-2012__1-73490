VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clean Up"
   ClientHeight    =   3075
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   4200
   Icon            =   "frmCleanUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame fraMain 
      Height          =   1410
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   615
      Width           =   3945
      Begin VB.CheckBox chkChkBox 
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
         Left            =   240
         TabIndex        =   13
         Top             =   270
         Width           =   210
      End
      Begin VB.CheckBox chkChkBox 
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
         Left            =   240
         TabIndex        =   12
         Top             =   990
         Width           =   210
      End
      Begin VB.CheckBox chkChkBox 
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
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   630
         Width           =   210
      End
      Begin VB.PictureBox picNA 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   180
         Picture         =   "frmCleanUp.frx":0442
         ScaleHeight     =   210
         ScaleWidth      =   330
         TabIndex        =   16
         Top             =   990
         Width           =   330
      End
      Begin VB.PictureBox picNA 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   180
         Picture         =   "frmCleanUp.frx":07B4
         ScaleHeight     =   210
         ScaleWidth      =   330
         TabIndex        =   15
         Top             =   270
         Width           =   330
      End
      Begin VB.PictureBox picNA 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   180
         Picture         =   "frmCleanUp.frx":0B26
         ScaleHeight     =   210
         ScaleWidth      =   330
         TabIndex        =   14
         Top             =   630
         Width           =   330
      End
      Begin VB.PictureBox picRedX 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   180
         Picture         =   "frmCleanUp.frx":0E98
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   19
         Top             =   270
         Width           =   210
      End
      Begin VB.PictureBox picRedX 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   180
         Picture         =   "frmCleanUp.frx":1214
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   18
         Top             =   630
         Width           =   210
      End
      Begin VB.PictureBox picRedX 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   180
         Picture         =   "frmCleanUp.frx":1590
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   17
         Top             =   990
         Width           =   210
      End
      Begin VB.PictureBox picGrnChk 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   0
         Left            =   90
         Picture         =   "frmCleanUp.frx":190C
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   22
         Top             =   135
         Width           =   450
      End
      Begin VB.PictureBox picGrnChk 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   1
         Left            =   90
         Picture         =   "frmCleanUp.frx":1CB9
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   21
         Top             =   495
         Width           =   450
      End
      Begin VB.PictureBox picGrnChk 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   2
         Left            =   90
         Picture         =   "frmCleanUp.frx":2066
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   20
         Top             =   855
         Width           =   450
      End
      Begin VB.Label lblWait 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Be patient"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2250
         TabIndex        =   10
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label lblFolders 
         BackStyle       =   0  'Transparent
         Caption         =   "Most Recently Used"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   540
         TabIndex        =   9
         Top             =   990
         Width           =   1800
      End
      Begin VB.Label lblFolders 
         BackStyle       =   0  'Transparent
         Caption         =   "Recycle Bin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   540
         TabIndex        =   8
         Top             =   630
         Width           =   1800
      End
      Begin VB.Label lblFolders 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Temp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   540
         TabIndex        =   7
         Top             =   270
         Width           =   1800
      End
   End
   Begin VB.PictureBox picManifest 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3465
      Left            =   0
      ScaleHeight     =   3465
      ScaleWidth      =   4200
      TabIndex        =   0
      Top             =   0
      Width           =   4200
      Begin VB.CommandButton cmdChoice 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   2520
         TabIndex        =   3
         Top             =   2145
         Width           =   690
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   3285
         TabIndex        =   2
         Top             =   2145
         Width           =   690
      End
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Select / Deselect All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   1
         Top             =   2160
         Width           =   2010
      End
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         Caption         =   "lblDisclaimer"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Clean Up Temp Files"
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
         Left            =   45
         TabIndex        =   5
         Top             =   90
         Width           =   4095
      End
      Begin VB.Label lblAuthor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Kenneth Ives"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1620
         TabIndex        =   4
         Top             =   390
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        frmMain
'
' Description:   This is the user interface to select what areas need to be
'                cleaned.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 16-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 07-Mar-2010  Kenneth Ives  kenaso@tx.rr.com
'              Renamed controls for easier maintenance.
' ***************************************************************************
Option Explicit

  Private mblnWindowsTemp As Boolean
  Private mblnRecycleBin  As Boolean
  Private mblnMostRecent  As Boolean
  Private mblnCallFlag    As Boolean
  
Private Sub chkChkBox_Click(Index As Integer)

    ' Determine which option has been checked
    Select Case Index
           Case 0: mblnWindowsTemp = CBool(chkChkBox(0).Value)
           Case 1: mblnRecycleBin = CBool(chkChkBox(1).Value)
           Case 2: mblnMostRecent = CBool(chkChkBox(2).Value)
    End Select
                
    ' Flag is to prevent looping
    If mblnCallFlag Then
        mblnCallFlag = False
        Exit Sub
    End If
    
    ' If all options have been checked
    ' then place checkmark in "Select all"
    If mblnWindowsTemp And _
       mblnRecycleBin And _
       mblnMostRecent Then
        
        mblnCallFlag = True
        chkSelectAll.Value = vbChecked
    Else
        ' At least one item was not checked.
        ' Remove "Select all" checkmark.
        mblnCallFlag = True
        chkSelectAll.Value = vbUnchecked
    End If
    
End Sub

Private Sub chkSelectAll_Click()

    Dim blnSelectAll As Boolean
    Dim intIndex     As Integer
    
    ' Has select all checkbox been checked
    blnSelectAll = CBool(chkSelectAll.Value)
    
    ' Flag is to prevent looping
    If mblnCallFlag Then
        mblnCallFlag = False
        Exit Sub
    End If
    
    ' If so then place a checkmark
    ' in all the checkboxes
    If blnSelectAll Then
        For intIndex = 0 To 2
            mblnCallFlag = True
            chkChkBox(intIndex).Value = vbChecked
        Next intIndex
    Else
        For intIndex = 0 To 2
            mblnCallFlag = True
            chkChkBox(intIndex).Value = vbUnchecked
        Next intIndex
    End If
        
    mblnCallFlag = False   ' Not needed
        
End Sub

Private Sub cmdChoice_Click(Index As Integer)

    Select Case Index
    
           Case 0
                ' disable buttons and display wait message
                cmdChoice(0).Enabled = False
                cmdChoice(1).Enabled = False
                lblWait.Visible = True
                
                BeginProcessing   ' Start the clean up
                ResetControls     ' reset the controls
                 
           Case Else
                TerminateProgram  ' shutdown this application
    End Select
  
End Sub

Private Sub Form_Load()

    mblnWindowsTemp = False  ' Preset boolean switches to FALSE
    mblnRecycleBin = False
    mblnMostRecent = False
    ResetControls            ' reset checkboxes and picture controls
    
    With frmMain
        .Caption = gstrVersion
        .lblDisclaimer.Caption = "This is a freeware product." & vbNewLine & _
                                 "No warranties or guarantees implied or intended."
        .chkSelectAll.Value = vbChecked
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2  ' Center form on screen
        .Show
    End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' "X" in upper right corner was selected
    If UnloadMode = 0 Then
        TerminateProgram
    End If
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail   ' Have a problem?  Notify the author.
End Sub

Private Sub ResetControls()

    Dim intIndex As Integer
    
    With frmMain
        
        .cmdChoice(0).Enabled = True
        .cmdChoice(1).Enabled = True
        .lblWait.Visible = False       ' hide the wait messge
        
        ' Hide all the picture checkboxes
        For intIndex = 0 To 2
            .picNA(intIndex).Visible = False
            .picRedX(intIndex).Visible = False
            .picGrnChk(intIndex).Visible = False
        Next intIndex
        
        ' Make checkboxes visible
        For intIndex = 0 To 2
            .chkChkBox(intIndex).Visible = True
        Next intIndex
        
    End With
    
End Sub

Private Sub BeginProcessing()

    Dim intIndex As Integer
    
    With frmMain
        
        ' hide all folder and picture checkboxes
        For intIndex = 0 To 2
            .chkChkBox(intIndex).Visible = False
            .picNA(intIndex).Visible = False
            .picRedX(intIndex).Visible = False
            .picGrnChk(intIndex).Visible = False
        Next intIndex
         
        ' Empty windows temp folder
        If mblnWindowsTemp Then
            If EmptyWindowsTemp Then
                .picGrnChk(0).Visible = True  ' Show green checkmark - Successful completion
            Else
                .picRedX(0).Visible = True    ' Show red "X" - Error in processing
            End If
        Else
            .picNA(0).Visible = True          ' display "n/a"
        End If
                        
        ' Empty recycle bin
        If mblnRecycleBin Then
            If EmptyRecycleBin Then
                .picGrnChk(1).Visible = True  ' Show green checkmark - Successful completion
            Else
                .picRedX(1).Visible = True    ' Show red "X" - Error in processing
            End If
        Else
            .picNA(1).Visible = True          ' display "n/a"
        End If
                         
        ' Empty most recent folder
        If mblnMostRecent Then
            If EmptyMostRecent Then
                .picGrnChk(2).Visible = True  ' Show green checkmark - Successful completion
            Else
                .picRedX(2).Visible = True    ' Show red "X" - Error in processing
            End If
        Else
            .picNA(2).Visible = True          ' display "n/a"
        End If
         
    End With
    
    Wait 2   ' pause for 2 seconds

End Sub


