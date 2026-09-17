VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Squad"
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

Private vm_Squad_o As Collection

Public Sub Add(ByVal pr_PlayerLeague_s As String, _
               ByVal pr_PlayerTeam_s As String, _
               ByVal pr_PlayerName_s As String, _
               ByVal pr_PlayerAge_i As Integer, _
               ByVal pr_PlayerLevel_i As Integer, _
               ByVal pr_PlayerCat_s As String, _
               ByVal pr_PlayerPos_s As String, _
               ByVal pr_PlayerType_s As String, _
               Optional ByVal pr_PlayerSess_i As Integer = 0, _
               Optional ByVal pr_PlayerNotes_s As String = vbNullString)

    Dim vl_Player_o As Player
    Set vl_Player_o = New Player

    vl_Player_o.Create pr_PlayerLeague_s, pr_PlayerTeam_s, pr_PlayerName_s, _
                       pr_PlayerAge_i, pr_PlayerLevel_i, pr_PlayerCat_s, _
                       pr_PlayerPos_s, pr_PlayerType_s, pr_PlayerSess_i, pr_PlayerNotes_s
    
    vm_Squad_o.Add vl_Player_o
    Set vl_Player_o = Nothing

End Sub

Public Sub AddPlayer(ByVal pr_Player_o As Player)
    vm_Squad_o.Add pr_Player_o
End Sub

Public Function Count() As Long
    Count = vm_Squad_o.Count
End Function

Public Function Item(ByVal pr_Index_l As Long) As Player
Attribute Item.VB_UserMemId = 0
    Set Item = vm_Squad_o.Item(pr_Index_l)
End Function

Public Sub Remove(ByVal pr_Index_l As Long)
    vm_Squad_o.Remove pr_Index_l
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = vm_Squad_o.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set vm_Squad_o = New Collection
End Sub

Private Sub Class_Terminate()
    Set vm_Squad_o = Nothing
End Sub
