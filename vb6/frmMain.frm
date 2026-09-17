VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "MarkInfoV2.1"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6585
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   3000
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraPlayer 
      Caption         =   "Player Maintenance"
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Width           =   6330
      Begin VB.CommandButton cmdMove 
         Caption         =   "&Move Player"
         Height          =   500
         Left            =   1812
         TabIndex        =   6
         Top             =   360
         Width           =   1200
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Player"
         Height          =   500
         Left            =   4824
         TabIndex        =   8
         Top             =   360
         Width           =   1200
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Player"
         Height          =   500
         Left            =   3318
         TabIndex        =   7
         Top             =   360
         Width           =   1200
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Player"
         Default         =   -1  'True
         Height          =   500
         Left            =   306
         TabIndex        =   5
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import Shomatch File"
      Height          =   995
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   210
      Width           =   1095
   End
   Begin VB.Frame fraSession 
      Caption         =   "Session"
      Height          =   1095
      Left            =   3960
      TabIndex        =   12
      Top             =   120
      Width           =   1095
      Begin VB.TextBox txtSession 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   550
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   2
         Top             =   400
         Width           =   315
      End
      Begin MSComCtl2.UpDown udnSession 
         Height          =   315
         Left            =   220
         TabIndex        =   13
         Top             =   400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         Alignment       =   0
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtSession"
         BuddyDispid     =   196616
         OrigLeft        =   600
         OrigTop         =   405
         OrigRight       =   840
         OrigBottom      =   720
         Max             =   16
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
   End
   Begin VB.ComboBox cboTeams 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   770
      Width           =   2055
   End
   Begin VB.ComboBox cboLeagues 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame fraTeam 
      Caption         =   "Selected Team"
      Height          =   1095
      Left            =   135
      TabIndex        =   9
      Top             =   120
      Width           =   3550
      Begin VB.Label lblTeam 
         Alignment       =   1  'Right Justify
         Caption         =   "Team :"
         Height          =   200
         Left            =   400
         TabIndex        =   11
         Top             =   670
         Width           =   700
      End
      Begin VB.Label lblLeague 
         Alignment       =   1  'Right Justify
         Caption         =   "League :"
         Height          =   200
         Left            =   400
         TabIndex        =   10
         Top             =   290
         Width           =   700
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgrSquad 
      Height          =   3450
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   6085
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewSeason 
         Caption         =   "&New Season"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMaint 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuImport 
         Caption         =   "&Import Database"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLeagueAdd 
         Caption         =   "New &League"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuLeagueDel 
         Caption         =   "&Delete League"
      End
      Begin VB.Menu mnuLeagueRename 
         Caption         =   "&Rename League"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTeamAdd 
         Caption         =   "New &Team"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuTeamDel 
         Caption         =   "De&lete Team"
      End
      Begin VB.Menu mnuTeamRename 
         Caption         =   "Re&name Team"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                                  '
' MarkInfo - A Player Database management program for Spellbinder Games' Kickabout '
' Copyright (C) 2000-2001  Tim Dodge                                               '
'                                                                                  '
' This program is free software; you can redistribute it and/or                    '
' modify it under the terms of the GNU General Public License                      '
' as published by the Free Software Foundation; either version 2                   '
' of the License, or (at your option) any later version.                           '
'                                                                                  '
' This program is distributed in the hope that it will be useful,                  '
' but WITHOUT ANY WARRANTY; without even the implied warranty of                   '
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                    '
' GNU General Public License for more details.                                     '
'                                                                                  '
' You should have received a copy of the GNU General Public License                '
' along with this program; if not, write to the Free Software                      '
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.      '
'                                                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub cboLeagues_Click()
    mdlMain.PopulateTeams cboLeagues.Text
End Sub

Private Sub cboTeams_Click()
    mdlMain.PopulateSquad cboTeams.Text, cboLeagues.Text
End Sub

Private Sub cmdAdd_Click()

    If (cboTeams <> "NO TEAMS") Then
    
        'Ensure that the session number has been saved
        vp_DBManager_o.Session = CInt(txtSession.Text)
        
        'Pre-populate then show the player form
        frmPlayer.txtAge = frmPlayer.udnAge.Min
        frmPlayer.txtLevel = frmPlayer.udnLevel.Min
        frmPlayer.cmdSave.Caption = "&Add"
        frmPlayer.cboNewTeams.Enabled = False
        frmPlayer.Show vbModal

    End If

End Sub

Private Sub cmdDelete_Click()

    Dim vl_Player_s As String
    
    If (fgrSquad.Row > 0) Then
    
        'Get player name
        vl_Player_s = vp_Squad_o.Item(fgrSquad.Row).PlayerName

        'Only delete if the user confirms the delete
        If (MsgBox("Really delete " & vl_Player_s & "?", vbYesNo, "Confirm Delete") = vbYes) Then
            vp_DBManager_o.DelPlayer vl_Player_s, cboTeams.Text, cboLeagues.Text
            mdlMain.PopulateSquad cboTeams.Text, cboLeagues.Text
        End If

    End If

End Sub

Private Sub cmdEdit_Click()

    If (fgrSquad.Row > 0) Then

        'Ensure that the session number has been saved
        vp_DBManager_o.Session = CInt(txtSession.Text)
    
        'Pre-populate then show the player form
        frmPlayer.txtName = vp_Squad_o.Item(fgrSquad.Row).PlayerName
        frmPlayer.txtAge = vp_Squad_o.Item(fgrSquad.Row).PlayerAge
        frmPlayer.txtLevel = vp_Squad_o.Item(fgrSquad.Row).PlayerLevel
        frmPlayer.cboCat.Text = vp_Squad_o.Item(fgrSquad.Row).PlayerCat
        frmPlayer.cboPos.Text = vp_Squad_o.Item(fgrSquad.Row).PlayerPos
        frmPlayer.cboType.Text = vp_Squad_o.Item(fgrSquad.Row).PlayerType
        frmPlayer.txtNotes = vp_Squad_o.Item(fgrSquad.Row).PlayerNotes
        frmPlayer.txtName.Locked = True
        frmPlayer.cboNewTeams.Enabled = False
        frmPlayer.Show vbModal

    End If

End Sub

Private Sub cmdImport_Click()

On Error GoTo ErrorHandler

    Dim vl_TempFolder_s As String

    'Make sure there is a league selected - ie it isn't "NO LEAGUES"
    If Not (cboLeagues.Text = "NO LEAGUES") Then

        'Ensure that the session number has been saved
        vp_DBManager_o.Session = CInt(txtSession.Text)

        'Open the dialog box
        With cdlMain
    
            .DialogTitle = "Choose Shomatch File"
            .InitDir = App.Path
            .Filter = "ShoMatch files|*.zip|Text files|*.txt|All files|*.*"
            .CancelError = True
            .ShowOpen
    
            MousePointer = vbHourglass
            ProcInputFile .FileName
            MousePointer = vbNormal

        End With

    Else
        MsgBox "You have to add a league first!"
    End If

    Exit Sub

ErrorHandler:

    If (Err.Number = cdlCancel) Then
        'Exit the subroutine - cancel was selected
        Exit Sub
    ElseIf (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "frmMain.cmdImport_Click", Err.Description
    End If

End Sub

Private Sub cmdMove_Click()

    If (fgrSquad.Row > 0) Then
        
        'Ensure that the session number has been saved
        vp_DBManager_o.Session = CInt(txtSession.Text)
    
        'Pre-populate then show the player form
        frmPlayer.txtName = vp_Squad_o.Item(fgrSquad.Row).PlayerName
        frmPlayer.txtAge = vp_Squad_o.Item(fgrSquad.Row).PlayerAge
        frmPlayer.txtLevel = vp_Squad_o.Item(fgrSquad.Row).PlayerLevel
        frmPlayer.cboCat.Text = vp_Squad_o.Item(fgrSquad.Row).PlayerCat
        frmPlayer.cboPos.Text = vp_Squad_o.Item(fgrSquad.Row).PlayerPos
        frmPlayer.cboType.Text = vp_Squad_o.Item(fgrSquad.Row).PlayerType
        frmPlayer.txtNotes = vp_Squad_o.Item(fgrSquad.Row).PlayerNotes
        frmPlayer.cboNewTeams.RemoveItem cboTeams.ListIndex
        frmPlayer.cboNewTeams.Text = frmPlayer.cboNewTeams.List(0)
        frmPlayer.cmdSave.Caption = "&Move"
        frmPlayer.txtName.Locked = True
        frmPlayer.Show vbModal

    End If

End Sub

Private Sub fgrSquad_DblClick()
    cmdEdit_Click
End Sub

Private Sub Form_Load()

    'Set flexgrid column widths
    fgrSquad.ColWidth(0) = 1620
    fgrSquad.ColWidth(1) = 660
    fgrSquad.ColWidth(2) = 660
    fgrSquad.ColWidth(3) = 660
    fgrSquad.ColWidth(4) = 660
    fgrSquad.ColWidth(5) = 660
    fgrSquad.ColWidth(6) = 660
    fgrSquad.ColWidth(7) = 660

    'Set flexgrid alignment
    fgrSquad.ColAlignment(0) = flexAlignLeftCenter
    fgrSquad.ColAlignment(1) = flexAlignCenterCenter
    fgrSquad.ColAlignment(2) = flexAlignCenterCenter
    fgrSquad.ColAlignment(3) = flexAlignCenterCenter
    fgrSquad.ColAlignment(4) = flexAlignCenterCenter
    fgrSquad.ColAlignment(5) = flexAlignCenterCenter
    fgrSquad.ColAlignment(6) = flexAlignCenterCenter
    fgrSquad.ColAlignment(7) = flexAlignCenterCenter

    'Put Column Headings into flexgrid
    fgrSquad.Row = 0
    fgrSquad.Col = 0
    fgrSquad.Text = "Name"
    fgrSquad.Col = 1
    fgrSquad.Text = "Age"
    fgrSquad.Col = 2
    fgrSquad.Text = "Level"
    fgrSquad.Col = 3
    fgrSquad.Text = "Cat"
    fgrSquad.Col = 4
    fgrSquad.Text = "Pos"
    fgrSquad.Col = 5
    fgrSquad.Text = "Type"
    fgrSquad.Col = 6
    fgrSquad.Text = "Session"
    fgrSquad.Col = 7
    fgrSquad.Text = "Notes?"

    'Get List of Leagues and show the first
    mdlMain.PopulateLeagues
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Make sure session number is saved to database
    vp_DBManager_o.Session = CInt(txtSession)

    'Clear up all objects
    Set vp_DBManager_o = Nothing
    Set vp_Squad_o = Nothing

End Sub

Private Sub mnuAbout_Click()
    MsgBox "Copyright 2000, 2001 Tim Dodge" & vbCrLf & vbCrLf & _
           "MarkInfo comes with ABSOLUTELY NO WARRANTY." & vbCrLf & vbCrLf & _
           "This is free software, you are welcome to redistribute it" & vbCrLf & _
           "under certain conditions; see the file COPYING for details.", _
           vbInformation, "MarkInfo v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuExit_Click()

    Dim vl_Form_o As Form
    
    For Each vl_Form_o In Forms
        Unload vl_Form_o
    Next

End Sub

Private Sub mnuImport_Click()

On Error GoTo ErrorHandler

    With cdlMain
        .DialogTitle = "Choose a MarkInfo Database"
        .InitDir = App.Path
        .Filter = "MarkInfo Databases|MarkInfo*.mdb"
        .CancelError = True
        .ShowOpen
    End With

    ImportData cdlMain.FileName
    PopulateLeagues
    Unload frmImport

    Exit Sub

ErrorHandler:

    If (Err.Number = cdlCancel) Then
        'do nothing - cancel was selected
    ElseIf (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "frmMain.cmdImport_Click", Err.Description
    End If

End Sub

Private Sub mnuLeagueAdd_Click()

    Dim vl_League_s As String
    
    vl_League_s = InputBox("Enter the name of the new League", "New League")
    
    If (vl_League_s <> vbNullString And Len(vl_League_s) <= 20) Then
        vp_DBManager_o.AddLeague UCase$(vl_League_s)
        PopulateLeagues
    End If

End Sub

Private Sub mnuLeagueDel_Click()

    If (cboLeagues.Text = "NO LEAGUES") Then
    
        MsgBox "You have to add a league first!"

    Else

        If (MsgBox("Really delete " & cboLeagues.Text & "?", vbYesNo, "Confirm Delete") = vbYes) Then
            vp_DBManager_o.DelLeague cboLeagues.Text
            PopulateLeagues
        End If

    End If

End Sub

Private Sub mnuLeagueRename_Click()
    
    Dim vl_League_s As String
    
    If (cboLeagues.Text = "NO LEAGUES") Then
    
        MsgBox "No Leagues!"

    Else

        vl_League_s = InputBox("Enter the new name for " & cboLeagues.Text, "Rename " & cboLeagues.Text)
    
        If (vl_League_s <> vbNullString And Len(vl_League_s) <= 20) Then
            vp_DBManager_o.RenameLeague cboLeagues.Text, UCase$(vl_League_s)
            PopulateLeagues
        End If

    End If

End Sub

Private Sub mnuNewSeason_Click()
    frmNewSess.Show
End Sub

Private Sub mnuTeamAdd_Click()

    Dim vl_Team_s As String
    
    If (cboLeagues.Text = "NO LEAGUES") Then
    
        MsgBox "You have to add a league first!"

    Else
    
        vl_Team_s = InputBox("Enter the name of the new Team", "New Team in " & cboLeagues.Text)
    
        If (vl_Team_s <> vbNullString And Len(vl_Team_s) <= 20) Then
            vp_DBManager_o.AddTeam UCase$(vl_Team_s), cboLeagues.Text
            PopulateTeams cboLeagues.Text
        End If

    End If

End Sub

Private Sub mnuTeamDel_Click()

    If (cboTeams.Text = "NO TEAMS") Then
    
        MsgBox "No Teams!"

    Else

        If (MsgBox("Really delete " & cboTeams.Text & "?", vbYesNo, "Confirm Delete") = vbYes) Then
            vp_DBManager_o.DelTeam cboTeams.Text, cboLeagues.Text
            PopulateTeams cboLeagues.Text
        End If

    End If
    
End Sub

Private Sub mnuTeamRename_Click()

    Dim vl_Team_s As String
    
    If (cboTeams.Text = "NO TEAMS") Then
    
        MsgBox "No Teams!"

    Else

        vl_Team_s = InputBox("Enter the new name for " & cboTeams.Text, "Rename " & cboTeams.Text)
    
        If (vl_Team_s <> vbNullString And Len(vl_Team_s) <= 20) Then
            vp_DBManager_o.RenameTeam cboTeams.Text, UCase$(vl_Team_s), cboLeagues.Text
            PopulateTeams cboLeagues.Text
        End If

    End If

End Sub
