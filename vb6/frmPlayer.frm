VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlayer 
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraNewTeam 
      Caption         =   "New Team"
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Width           =   2295
      Begin VB.ComboBox cboNewTeams 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   495
      Left            =   3590
      TabIndex        =   9
      Top             =   3275
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   2515
      TabIndex        =   8
      Top             =   3275
      Width           =   975
   End
   Begin VB.Frame fraNotes 
      Caption         =   "Notes"
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   4445
      Begin VB.TextBox txtNotes 
         Height          =   800
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   4200
      End
   End
   Begin VB.Frame fraType 
      Caption         =   "Type"
      Height          =   735
      Left            =   3150
      TabIndex        =   17
      Top             =   960
      Width           =   1415
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1175
      End
   End
   Begin VB.Frame fraPosition 
      Caption         =   "Position"
      Height          =   735
      Left            =   1635
      TabIndex        =   16
      Top             =   960
      Width           =   1415
      Begin VB.ComboBox cboPos 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1175
      End
   End
   Begin VB.Frame fraCategory 
      Caption         =   "Category"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   1415
      Begin VB.ComboBox cboCat 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Frame fraLevel 
      Caption         =   "Level"
      Height          =   735
      Left            =   3590
      TabIndex        =   14
      Top             =   120
      Width           =   975
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   2
         Top             =   240
         Width           =   315
      End
      Begin MSComCtl2.UpDown udnLevel 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Alignment       =   0
         BuddyControl    =   "txtLevel"
         BuddyDispid     =   196622
         OrigLeft        =   120
         OrigTop         =   240
         OrigRight       =   360
         OrigBottom      =   555
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame fraAge 
      Caption         =   "Age"
      Height          =   735
      Left            =   2515
      TabIndex        =   13
      Top             =   120
      Width           =   975
      Begin VB.TextBox txtAge 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   1
         Top             =   240
         Width           =   315
      End
      Begin MSComCtl2.UpDown udnAge 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   17
         Alignment       =   0
         BuddyControl    =   "txtAge"
         BuddyDispid     =   196624
         OrigLeft        =   120
         OrigTop         =   240
         OrigRight       =   360
         OrigBottom      =   555
         Max             =   99
         Min             =   17
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame fraName 
      Caption         =   "Name"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmPlayer"
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

Private Sub cboCat_Validate(Cancel As Boolean)

    If (cboCat.Text = "---") Then
        cboPos.Text = "GK"
    End If

End Sub

Private Sub cboPos_Validate(Cancel As Boolean)

    If (cboPos.Text = "GK") Then
        cboCat.Text = "---"
    End If

    If (cboPos.Text <> "GK" And cboCat.Text = "---") Then
        cboCat.Text = "?"
    End If

End Sub

Private Sub cboType_Validate(Cancel As Boolean)

    If (cboType.Text = "STAR" And CInt(txtLevel) < 12) Then
        If (CInt(txtAge) = 17) Then
            cboType = "SBY"
        Else
            cboType.Text = "---"
        End If
    End If

    If (cboType.Text = "SBY" And CInt(txtAge) > 18) Then
        cboType.Text = "---"
    End If
    
    If ((cboType.Text = "APP" Or cboType.Text = "FUT") And CInt(txtAge) <> 18) Then
        cboType.Text = "---"
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload frmPlayer
End Sub

Private Sub cmdSave_Click()

    Dim vl_Player_o As Player

    If (txtName = vbNullString) Then
        MsgBox "Please name the player!"
        Exit Sub
    End If

    'Ask if changes made are okay
    If (MsgBox(Mid$(cmdSave.Caption, 2) & " player " & txtName & "?", vbYesNo, "Confirm Changes") = vbYes) Then

        'Create the player object
        Set vl_Player_o = New Player
        vl_Player_o.Create frmMain.cboLeagues.Text, cboNewTeams.Text, txtName, _
                           CInt(txtAge), CInt(txtLevel), cboCat.Text, cboPos.Text, _
                           cboType.Text, 0, txtNotes

        'Add, Change or Move the player as necessary
        If (cboNewTeams.Enabled) Then  'Move
            vp_DBManager_o.MovePlayer vl_Player_o, frmMain.cboTeams.Text
        Else 'Add or Change
            vp_DBManager_o.AddPlayer vl_Player_o
        End If

        PopulateSquad frmMain.cboTeams.Text, frmMain.cboLeagues.Text
        Set vl_Player_o = Nothing

        Unload frmPlayer

    End If

End Sub

Private Sub Form_Load()

    Dim vl_Index_i As Integer
    Dim vl_Cats_s() As String
    Dim vl_Posns_s() As String
    Dim vl_Types_s() As String
    
    'Set the form caption
    Caption = frmMain.cboTeams & " - " & frmMain.cboLeagues & " League"
    
    'Populate the League combo
    For vl_Index_i = 0 To frmMain.cboTeams.ListCount - 1
        cboNewTeams.AddItem frmMain.cboTeams.List(vl_Index_i)
    Next
    cboNewTeams.Text = frmMain.cboTeams.Text

    'Populate the Categories combo
    vl_Cats_s = vp_DBManager_o.GetCats
    For vl_Index_i = LBound(vl_Cats_s) To UBound(vl_Cats_s)
        cboCat.AddItem vl_Cats_s(vl_Index_i)
    Next
    cboCat.Text = cboCat.List(0)
    
    'Populate the Positions combo
    vl_Posns_s = vp_DBManager_o.GetPosns
    For vl_Index_i = LBound(vl_Posns_s) To UBound(vl_Posns_s)
        cboPos.AddItem vl_Posns_s(vl_Index_i)
    Next
    cboPos.Text = cboPos.List(0)
    
    'Populate the Types combo
    vl_Types_s = vp_DBManager_o.GetTypes
    For vl_Index_i = LBound(vl_Types_s) To UBound(vl_Types_s)
        cboType.AddItem vl_Types_s(vl_Index_i)
    Next
    cboType.Text = cboType.List(0)

End Sub

Private Sub txtAge_Validate(Cancel As Boolean)
    AgeValidate
End Sub

Private Sub txtLevel_Validate(Cancel As Boolean)
    LevelValidate
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    txtName.Text = UCase$(txtName.Text)
End Sub

Private Sub AgeValidate()

    If (CInt(txtAge) = 17) Then
        cboType.Text = "SBY"
    End If

    If (CInt(txtAge) > 18 And cboType.Text = "SBY") Then
        cboType.Text = "---"
    End If
    
    If (CInt(txtAge) <> 18 And (cboType.Text = "APP" Or cboType.Text = "FUT")) Then
        cboType.Text = "---"
    End If

    If (CInt(txtAge) < 19) Then
        If (CInt(txtLevel) > 12) Then
            txtLevel = 12
            cboType.Text = "---"
        End If
        udnLevel.Max = 12
    Else
        udnLevel.Max = 99
    End If

End Sub

Private Sub LevelValidate()

    If (CInt(txtLevel) > 12) Then
        cboType.Text = "STAR"
    End If

    If (CInt(txtLevel) < 12 And cboType.Text = "STAR") Then
        cboType.Text = "---"
    End If

End Sub

Private Sub udnAge_DownClick()
    AgeValidate
End Sub

Private Sub udnAge_UpClick()
    AgeValidate
End Sub

Private Sub udnLevel_DownClick()
    LevelValidate
End Sub

Private Sub udnLevel_UpClick()
    LevelValidate
End Sub
