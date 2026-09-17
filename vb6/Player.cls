VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Private vm_PlayerLeague_s As String
Private vm_PlayerTeam_s As String
Private vm_PlayerName_s As String
Private vm_PlayerAge_i As Integer
Private vm_PlayerLevel_i As Integer
Private vm_PlayerCat_s As String
Private vm_PlayerPos_s As String
Private vm_PlayerType_s As String
Private vm_PlayerSess_i As Integer
Private vm_PlayerNotes_s As String

Public Sub Create(ByVal pr_PlayerLeague_s As String, _
                  ByVal pr_PlayerTeam_s As String, _
                  ByVal pr_PlayerName_s As String, _
                  ByVal pr_PlayerAge_i As Integer, _
                  ByVal pr_PlayerLevel_i As Integer, _
                  ByVal pr_PlayerCat_s As String, _
                  ByVal pr_PlayerPos_s As String, _
                  ByVal pr_PlayerType_s As String, _
                  ByVal pr_PlayerSess_i As Integer, _
                  ByVal pr_PlayerNotes_s As String)

    vm_PlayerLeague_s = pr_PlayerLeague_s
    vm_PlayerTeam_s = pr_PlayerTeam_s
    vm_PlayerName_s = pr_PlayerName_s
    vm_PlayerAge_i = pr_PlayerAge_i
    vm_PlayerLevel_i = pr_PlayerLevel_i
    vm_PlayerCat_s = pr_PlayerCat_s
    vm_PlayerPos_s = pr_PlayerPos_s
    vm_PlayerType_s = pr_PlayerType_s
    vm_PlayerSess_i = pr_PlayerSess_i
    vm_PlayerNotes_s = pr_PlayerNotes_s

End Sub


Public Property Get PlayerLeague() As String
    PlayerLeague = vm_PlayerLeague_s
End Property

Public Property Let PlayerLeague(ByVal pr_PlayerLeague_s As String)
    vm_PlayerLeague_s = pr_PlayerLeague_s
End Property

Public Property Get PlayerTeam() As String
    PlayerTeam = vm_PlayerTeam_s
End Property

Public Property Let PlayerTeam(ByVal pr_PlayerTeam_s As String)
    vm_PlayerTeam_s = pr_PlayerTeam_s
End Property

Public Property Get PlayerName() As String
    PlayerName = vm_PlayerName_s
End Property

Public Property Let PlayerName(ByVal pr_PlayerName_s As String)
    vm_PlayerName_s = pr_PlayerName_s
End Property

Public Property Get PlayerAge() As Integer
    PlayerAge = vm_PlayerAge_i
End Property

Public Property Let PlayerAge(ByVal pr_PlayerAge_i As Integer)
    vm_PlayerAge_i = pr_PlayerAge_i
End Property

Public Property Get PlayerLevel() As Integer
    PlayerLevel = vm_PlayerLevel_i
End Property

Public Property Let PlayerLevel(ByVal pr_PlayerLevel_i As Integer)
    vm_PlayerLevel_i = pr_PlayerLevel_i
End Property

Public Property Get PlayerCat() As String
    PlayerCat = vm_PlayerCat_s
End Property

Public Property Let PlayerCat(ByVal pr_PlayerCat_s As String)
    vm_PlayerCat_s = pr_PlayerCat_s
End Property

Public Property Get PlayerPos() As String
    PlayerPos = vm_PlayerPos_s
End Property

Public Property Let PlayerPos(ByVal pr_PlayerPos_s As String)
    vm_PlayerPos_s = pr_PlayerPos_s
End Property

Public Property Get PlayerType() As String
    PlayerType = vm_PlayerType_s
End Property

Public Property Let PlayerType(ByVal pr_PlayerType_s As String)
    vm_PlayerType_s = pr_PlayerType_s
End Property

Public Property Get PlayerSess() As Integer
    PlayerSess = vm_PlayerSess_i
End Property

Public Property Let PlayerSess(ByVal pr_PlayerSess_i As Integer)
    vm_PlayerSess_i = pr_PlayerSess_i
End Property

Public Property Get PlayerNotes() As String
    PlayerNotes = vm_PlayerNotes_s
End Property

Public Property Let PlayerNotes(ByVal pr_PlayerNotes_s As String)
    vm_PlayerNotes_s = pr_PlayerNotes_s
End Property
