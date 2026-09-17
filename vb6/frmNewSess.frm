VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNewSess 
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2910
   Icon            =   "frmNewSess.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.Frame fraNewSess 
      Caption         =   "Player Pruning"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.TextBox txtNewSess 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   2
         Top             =   925
         Width           =   315
      End
      Begin MSComCtl2.UpDown udnNewSess 
         Height          =   315
         Left            =   1070
         TabIndex        =   1
         Top             =   925
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   2
         Alignment       =   0
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtNewSess"
         BuddyDispid     =   196612
         OrigLeft        =   1060
         OrigTop         =   925
         OrigRight       =   1300
         OrigBottom      =   1240
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "sessions"
         Height          =   255
         Left            =   1720
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "more than"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "have not been updated for"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   660
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Remove players whose details"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmNewSess"
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

Private Sub cmdCancel_Click()
    Unload frmNewSess
End Sub

Private Sub cmdOK_Click()
    
    'Disable OK Button, pointer to hourglass
    cmdOK.Enabled = False
    MousePointer = vbHourglass
    
    'Do the new season processing
    vp_DBManager_o.NewSeason frmMain.cboLeagues.Text, CInt(txtNewSess)
    PopulateSquad frmMain.cboTeams.Text, frmMain.cboLeagues.Text
    
    'Enable OK etc
    MousePointer = vbNormal
    cmdOK.Enabled = True
    
    Unload frmNewSess
    
End Sub

Private Sub Form_Load()
    Caption = "New Season - " & frmMain.cboLeagues.Text & " League"
    txtNewSess = "2"
End Sub
